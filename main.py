import os
import logging
import uuid
import sys
import subprocess
import time
import base64
import tempfile
import asyncio
import json
from contextlib import asynccontextmanager
from logging.handlers import RotatingFileHandler
from typing import Optional, List, Dict

import win32print
import win32ui
import win32con
import win32api
import socket
from fastapi import FastAPI, HTTPException, BackgroundTasks, Depends, Header, Request, WebSocket, WebSocketDisconnect
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, field_validator
import uvicorn
from dotenv import load_dotenv
from fastapi.middleware.cors import CORSMiddleware

# --- CONFIGURAÇÃO E LOGGING ---
load_dotenv()
LOG_FILE = "printer_api.log"

def setup_logger():
    logger = logging.getLogger("PrinterAPI")
    logger.setLevel(logging.INFO)
    if logger.hasHandlers():
        logger.handlers.clear()
    
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    # File Handler
    file_handler = RotatingFileHandler(LOG_FILE, maxBytes=1000000, backupCount=5, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    logger.propagate = False
    return logger

logger = setup_logger()

# --- MODELOS ---
class ImpressaoRequest(BaseModel):
    num_pedido: str
    impressora: str
    conteudo: str

    @field_validator('impressora', 'num_pedido', mode='before')
    @classmethod
    def sanitize_strings(cls, v):
        if isinstance(v, str):
            return v.strip().strip('"').strip("'").strip()
        return v

# --- SERVIÇOS ---
class ConnectionManager:
    """Gerencia conexões WebSocket para atualizações real-time."""
    def __init__(self):
        self.active_connections: List[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        if websocket in self.active_connections:
            self.active_connections.remove(websocket)

    async def broadcast(self, message: dict):
        for connection in self.active_connections:
            try:
                await connection.send_json(message)
            except Exception:
                pass

class PrinterService:
    """Encapsula toda a interação com a API de impressão do Windows."""
    
    @staticmethod
    def parse_status(status_code: int) -> List[str]:
        FLAGS = {
            win32print.PRINTER_STATUS_PAUSED: "Pausada",
            win32print.PRINTER_STATUS_ERROR: "Erro",
            win32print.PRINTER_STATUS_PAPER_OUT: "Sem Papel",
            win32print.PRINTER_STATUS_OFFLINE: "Offline",
            win32print.PRINTER_STATUS_BUSY: "Ocupada",
            win32print.PRINTER_STATUS_PRINTING: "Imprimindo",
        }
        if status_code == 0: return ["Pronta"]
        return [label for flag, label in FLAGS.items() if status_code & flag] or [f"Status {status_code}"]

    def get_all_printers(self) -> List[dict]:
        data = []
        try:
            enum_flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            printers = win32print.EnumPrinters(enum_flags)
            for p in printers:
                try:
                    h = win32print.OpenPrinter(p[2])
                    info = win32print.GetPrinter(h, 2)
                    status_list = self.parse_status(info['Status'])
                    data.append({
                        "nome": str(p[2]),
                        "status_list": status_list,
                        "status_principal": status_list[0],
                        "jobs": info['cJobs'],
                        "driver": str(info['pDriverName']),
                        "porta": str(info['pPortName'])
                    })
                    win32print.ClosePrinter(h)
                except: continue
        except Exception as e:
            logger.error(f"Erro ao listar impressoras: {e}")
        return data

    def detect_type(self, conteudo: str) -> str:
        # Remover prefixo data URI antes de detectar
        c = conteudo.split(",", 1)[1] if "," in conteudo else conteudo
        c = c.strip()
        if c.startswith("JVBERi"):
            return "pdf"
        if c.strip().startswith("^XA"):
            return "zebra"
        return "comum"

    async def print_common(self, printer: str, text: str, order_id: str):
        try:
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer)
            hdc.StartDoc(f"Pedido_{order_id}")
            hdc.StartPage()
            font = win32ui.CreateFont({"name": "Arial", "height": 40, "weight": 400})
            hdc.SelectObject(font)
            y = 100
            for line in text.split("\n"):
                hdc.TextOut(100, y, line)
                y += 60
            hdc.EndPage(); hdc.EndDoc(); hdc.DeleteDC()
            logger.info(f"Pedido {order_id}: GDI OK")
        except Exception as e: logger.error(f"Pedido {order_id} GDI Error: {e}")

    async def print_zebra(self, printer: str, zpl: str, order_id: str):
        try:
            h = win32print.OpenPrinter(printer)
            win32print.StartDocPrinter(h, 1, (f"Pedido_{order_id}", None, "RAW"))
            win32print.StartPagePrinter(h)
            win32print.WritePrinter(h, zpl.encode("utf-8"))
            win32print.EndPagePrinter(h); win32print.EndDocPrinter(h); win32print.ClosePrinter(h)
            logger.info(f"Pedido {order_id}: ZPL OK")
        except Exception as e: logger.error(f"Pedido {order_id} ZPL Error: {e}")

    async def print_pdf(self, printer: str, temp_path: str, order_id: str):
        try:
            if hasattr(sys, '_MEIPASS'):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))

            sumatra = os.path.join(base_dir, "SumatraPDF.exe")

            if not os.path.exists(sumatra):
                raise FileNotFoundError(f"SumatraPDF.exe não encontrado em: {sumatra}")

            cmd = [
                sumatra,
                "-print-to", printer,
                "-print-settings", "noscale",
                "-silent",
                temp_path
            ]
            result = subprocess.run(cmd, timeout=30, capture_output=True)
            if result.returncode != 0:
                logger.error(f"Pedido {order_id} SumatraPDF erro: {result.stderr.decode(errors='replace')}")
            else:
                logger.info(f"Pedido {order_id}: PDF OK via SumatraPDF")
            await asyncio.sleep(5)
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception as e:
            logger.error(f"Pedido {order_id} PDF Error: {e}")

# --- SINGLETONS ---
ws_manager = ConnectionManager()
printer_service = PrinterService()

# --- APP SETUP ---
@asynccontextmanager
async def lifespan(app: FastAPI):
    async def broadcast_loop():
        while True:
            if ws_manager.active_connections:
                data = {
                    "impressoras": printer_service.get_all_printers(),
                    "logs": [line.strip() for line in open(LOG_FILE, "r", encoding='utf-8', errors='replace').readlines()[-25:]] if os.path.exists(LOG_FILE) else []
                }
                await ws_manager.broadcast(data)
            await asyncio.sleep(1)
    
    task = asyncio.create_task(broadcast_loop())
    yield
    task.cancel()

app = FastAPI(title="Printer API Gateway", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

templates = Jinja2Templates(directory=resource_path("templates"))

# --- AUTH ---
API_KEY = os.getenv("API_KEY", "minha_chave_segura_123")
async def verify_auth(x_api_key: str = Header(None)):
    if x_api_key != API_KEY:
        raise HTTPException(status_code=403, detail="API Key Inválida")
    return x_api_key

# --- ENDPOINTS ---
@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await ws_manager.connect(websocket)
    try:
        while True: await websocket.receive_text()
    except WebSocketDisconnect: ws_manager.disconnect(websocket)

@app.get("/admin", response_class=HTMLResponse)
async def admin_page(request: Request):
    return templates.TemplateResponse("admin.html", {"request": request, "api_key": API_KEY})

@app.get("/api/status")
async def get_api_status():
    return {
        "impressoras": printer_service.get_all_printers(),
        "logs": [line.strip() for line in open(LOG_FILE, "r", encoding='utf-8', errors='replace').readlines()[-25:]] if os.path.exists(LOG_FILE) else []
    }

@app.post("/imprimir")
async def post_imprimir(pedido: ImpressaoRequest, auth=Depends(verify_auth), background_tasks: BackgroundTasks = BackgroundTasks()):
    # Validação imediata
    try:
        h = win32print.OpenPrinter(pedido.impressora)
        win32print.ClosePrinter(h)
    except:
        raise HTTPException(status_code=404, detail=f"Impressora '{pedido.impressora}' não encontrada.")

    tipo = printer_service.detect_type(pedido.conteudo)
    
    if tipo == "zebra":
        background_tasks.add_task(printer_service.print_zebra, pedido.impressora, pedido.conteudo, pedido.num_pedido)
    elif tipo == "pdf":
        try:
            conteudo = pedido.conteudo
            # Remover prefixo data URI se vier com ele
            if "," in conteudo:
                conteudo = conteudo.split(",", 1)[1]
            # Limpar espaços e quebras de linha que invalidam o base64
            conteudo = conteudo.strip().replace(" ", "+")
            pdf_bytes = base64.b64decode(conteudo)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(pdf_bytes)
                background_tasks.add_task(printer_service.print_pdf, pedido.impressora, tmp.name, pedido.num_pedido)
        except Exception as ex:
            raise HTTPException(status_code=400, detail=f"Base64 de PDF inválido: {ex}")
    else:
        background_tasks.add_task(printer_service.print_common, pedido.impressora, pedido.conteudo, pedido.num_pedido)

    return {"status": "ok", "pedido": pedido.num_pedido, "engine": tipo}

@app.get("/impressoras")
def listar_impressoras(x_api_key: str = Header(...)):
    if x_api_key != API_KEY:
        raise HTTPException(status_code=401)
    return {"impressoras": printer_service.get_all_printers()}

@app.get("/hostname")
def get_hostname(x_api_key: str = Header(...)):
    if x_api_key != API_KEY:
        raise HTTPException(status_code=401)
    return {"hostname": socket.gethostname()}

if __name__ == "__main__":
    port = int(os.getenv("PORTA_API", 5000))
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=port,
        log_config=None  # ← desativa o log config padrão do uvicorn
    )
