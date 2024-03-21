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
        # WinHTTP 的 WebSocket 实现无法直接接收 UTF-8 文本
        # 所以直接发送二进制，并使用 VBA 程序解码为文本
        text_bytes = text.encode("utf-8")
        n_bytes = len(text_bytes)
        await websocket.send_bytes(n_bytes.to_bytes(4, "little"))
        await websocket.send_bytes(text_bytes)

    async def receive_text(self, websocket: WebSocket):
        # WinHTTP 的 WebSocket 发送的文本消息为 UTF-16 编码
        # 无法直接被 FastAPI 解码，所以直接接收二进制并解码为文本
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
    # 消息队列，格式为 (judge, department, web_path)
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
            # 不清空同名评委的 temp 文件夹，因为该评委可能因为网络问题断线重连
            # 清空 temp 文件夹会导致该评委已经上传的评分表丢失
            if not os.path.exists(f"temp/{judge}"):
                os.mkdir(f"temp/{judge}")
            while True:
                # 接收评分表信息，并添加到消息队列
                department = await manager.receive_text(websocket)
                data = await websocket.receive_bytes()
                save_path = f"temp/{judge}/{department}.xlsx"
                with open(save_path, "wb") as f:
                    f.write(data)
                web_path = base62_encode(save_path)
                await msg_queue.put((judge, department, web_path))
        elif role == "merger":
            while True:
                # 汇总者每隔 1s 发送一条 available 消息，以检查是否有新上传的评分表
                await manager.receive_text(websocket)  # 接收 available 消息
                if not msg_queue.empty():
                    # 有新上传的评分表，返回 true 并发送评分表信息
                    await manager.send_text("true", websocket)
                    judge, department, web_path = await msg_queue.get()
                    await manager.send_text(judge, websocket)
                    await manager.send_text(department, websocket)
                    await manager.send_text(web_path, websocket)
                else:
                    # 没有新上传的评分表，返回 false
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
