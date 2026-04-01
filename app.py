"""
Tech Estrutural - Interface Web para Automação de Engenharia
Uso: python app.py
"""
import os
import sys
import json
import queue
import threading
import subprocess
import webbrowser
from pathlib import Path

import uvicorn
from fastapi import FastAPI
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

# Diálogos nativos do SO via tkinter (sem exibir janela)
import tkinter as tk
from tkinter import filedialog

app = FastAPI(title="Tech Estrutural Automação", docs_url=None, redoc_url=None)
app.mount("/assets", StaticFiles(directory="assets"), name="assets")

_BASE = Path(__file__).parent
_SCRIPTS = _BASE / "scripts"
_log_queue: queue.Queue = queue.Queue()
_is_running = threading.Event()


# ─────────────────────────────────────────────────────────────
# Rotas de UI
# ─────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
def index():
    return HTMLResponse((_BASE / "templates" / "index.html").read_text(encoding="utf-8"))


# ─────────────────────────────────────────────────────────────
# Diálogos nativos
# ─────────────────────────────────────────────────────────────

def _native_dialog(fn, *args, **kwargs):
    """Executa um diálogo tkinter em thread separada e retorna o resultado."""
    result = {"value": ""}
    done = threading.Event()

    def _run():
        root = tk.Tk()
        root.withdraw()
        root.lift()
        root.attributes("-topmost", True)
        result["value"] = fn(*args, parent=root, **kwargs) or ""
        root.destroy()
        done.set()

    threading.Thread(target=_run, daemon=True).start()
    done.wait(timeout=120)
    return result["value"]


@app.get("/api/browse/folder")
def browse_folder(title: str = "Selecionar Pasta"):
    path = _native_dialog(filedialog.askdirectory, title=title)
    return JSONResponse({"path": path})


@app.get("/api/browse/file")
def browse_file(title: str = "Selecionar Planilha"):
    path = _native_dialog(
        filedialog.askopenfilename,
        title=title,
        filetypes=[("Planilha Excel com Macro", "*.xlsm"), ("Excel", "*.xlsx")],
    )
    return JSONResponse({"path": path})


# ─────────────────────────────────────────────────────────────
# Execução da automação
# ─────────────────────────────────────────────────────────────

@app.post("/api/executar")
def executar(body: dict):
    if _is_running.is_set():
        return JSONResponse({"error": "Processo já em execução."}, status_code=400)

    pastas: dict = body.get("pastas", {})
    excel: str = body.get("excel", "")

    if not excel or not Path(excel).exists():
        return JSONResponse({"error": "Planilha não encontrada."}, status_code=400)

    pastas_validas = {k: v for k, v in pastas.items() if v and Path(v).exists()}
    if not pastas_validas:
        return JSONResponse({"error": "Nenhuma pasta de desenhos válida selecionada."}, status_code=400)

    # Limpar fila de log anterior
    while not _log_queue.empty():
        try:
            _log_queue.get_nowait()
        except queue.Empty:
            break

    threading.Thread(target=_worker, args=(pastas_validas, excel), daemon=True).start()
    return JSONResponse({"status": "started"})


def _worker(pastas: dict, excel: str):
    _is_running.set()
    python = sys.executable
    converter = str(_SCRIPTS / "converter_dwg_dxf.py")
    extractor = str(_SCRIPTS / "extrair_dados_dxf.py")
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    total = len(pastas)
    erros = []

    try:
        for idx, (nome, pasta) in enumerate(pastas.items(), 1):
            _log(f"\n{'━' * 56}")
            _log(f"  [{idx}/{total}]  {nome.upper()}")
            _log(f"  {pasta}")
            _log(f"{'━' * 56}")

            # Etapa 1: Converter DWG → DXF
            _log(f"\n  [1/2] Convertendo DWG → DXF...")
            ok = _run_proc([python, converter, pasta], flags)
            if not ok:
                erros.append(nome)
                _log(f"  ✗ Falha na conversão de {nome}.")
                continue
            _log(f"  ✓ Conversão concluída.")

            # Etapa 2: Extrair dados
            _log(f"\n  [2/2] Extraindo dados e atualizando planilha...")
            ok = _run_proc([python, extractor, pasta, excel], flags)
            if not ok:
                erros.append(nome)
                _log(f"  ✗ Falha na extração de {nome}.")
                continue
            _log(f"  ✓ {nome} processado.")

        _log(f"\n{'━' * 56}")
        if erros:
            _log(f"  CONCLUÍDO COM ERROS em: {', '.join(erros)}")
        else:
            _log(f"  AUTOMAÇÃO CONCLUÍDA COM SUCESSO!")
        _log(f"{'━' * 56}\n")

    except Exception as exc:
        _log(f"\n  ✗ ERRO CRÍTICO: {exc}")
    finally:
        _is_running.clear()
        _log("__DONE__")


def _log(msg: str):
    _log_queue.put(msg)


def _run_proc(cmd: list, flags: int) -> bool:
    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=flags,
        )
        for line in proc.stdout:
            _log("  " + line.rstrip())
        proc.wait()
        return proc.returncode == 0
    except Exception as exc:
        _log(f"  ✗ {exc}")
        return False


# ─────────────────────────────────────────────────────────────
# SSE — stream de log em tempo real
# ─────────────────────────────────────────────────────────────

@app.get("/api/log/stream")
def log_stream():
    def _generate():
        while True:
            try:
                msg = _log_queue.get(timeout=25)
                yield f"data: {json.dumps({'text': msg})}\n\n"
                if msg == "__DONE__":
                    return
            except queue.Empty:
                yield "data: {\"ping\":true}\n\n"

    return StreamingResponse(
        _generate(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/api/status")
def status():
    return JSONResponse({"running": _is_running.is_set()})


# ─────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    def _open_browser():
        import time
        time.sleep(1.5)
        webbrowser.open("http://localhost:8000")

    threading.Thread(target=_open_browser, daemon=True).start()
    print("Tech Estrutural · Automação iniciada → http://localhost:8000")
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="warning")
