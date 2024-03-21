import uvicorn
import os
import shutil
from io import BytesIO
from asyncio import Queue
from fastapi import FastAPI, WebSocket, WebSocketDisconnect, Request
from fastapi.responses import FileResponse
from base62 import *


class ConnectionManager:
    def __init__(self):
        self.active_connections: list[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        self.active_connections.remove(websocket)

    async def send_text(self, text: str, websocket: WebSocket):
        text_bytes = text.encode("utf-8")
        n_bytes = len(text_bytes)
        await websocket.send_bytes(n_bytes.to_bytes(4, "little"))
        await websocket.send_bytes(text_bytes)

    async def receive_text(self, websocket: WebSocket):
        data = await websocket.receive_bytes()
        return data.decode("utf-16")

    async def send_bytes(self, data: bytes, websocket: WebSocket):
        n_bytes = len(data)
        await websocket.send_bytes(n_bytes.to_bytes(4, "little"))
        bis = BytesIO(data)
        chunk_size = 1024
        while True:
            chunk = bis.read(chunk_size)
            if not chunk:
                break
            await websocket.send_bytes(chunk)

    async def receive_bytes(self, websocket: WebSocket):
        return await websocket.receive_bytes()


app = FastAPI()
manager = ConnectionManager()
msg_queue: Queue


@app.on_event("startup")
async def startup_event():
    global msg_queue
    msg_queue = Queue(100)


@app.on_event("shutdown")
def shutdown_event():
    shutil.rmtree("temp")


@app.websocket("/ws/{client_id}")
async def websocket_endpoint(websocket: WebSocket, client_id: str):
    global msg_queue
    await manager.connect(websocket)
    role = await manager.receive_text(websocket)
    try:
        if role == "judge":
            judge = client_id
            if os.path.exists(f"temp/{judge}"):
                shutil.rmtree(f"temp/{judge}")
            os.mkdir(f"temp/{judge}")
            while True:
                department = await manager.receive_text(websocket)
                data = await websocket.receive_bytes()
                save_path = f"temp/{judge}/{department}.xlsx"
                with open(save_path, "wb") as f:
                    f.write(data)
                web_path = base62_encode(save_path)
                await msg_queue.put((judge, department, web_path))
        elif role == "merger":
            while True:
                await manager.receive_text(websocket)
                if not msg_queue.empty():
                    await manager.send_text("true", websocket)
                    judge, department, web_path = await msg_queue.get()
                    await manager.send_text(judge, websocket)
                    await manager.send_text(department, websocket)
                    await manager.send_text(web_path, websocket)
                else:
                    await manager.send_text("false", websocket)
    except WebSocketDisconnect:
        manager.disconnect(websocket)
        if role == "judge":
            print(f"评委 {client_id} 已下线")
        else:
            msg_queue = Queue(100)
            shutil.rmtree(f"temp")
            os.mkdir("temp")
            print(f"评分汇总完成！")


@app.get("/rating_table/{table_path}")
async def download_rating_table(table_path: str):
    table_path = base62_decode(table_path)
    file_path = os.path.abspath(table_path)
    return FileResponse(file_path)


if __name__ == "__main__":
    if os.path.exists("temp"):
        shutil.rmtree("temp")
    os.mkdir("temp")
    uvicorn.run(app, host="0.0.0.0", port=5422)
