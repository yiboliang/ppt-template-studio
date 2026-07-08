"""
PPT Template Studio — FastAPI 后端

接口:
  GET  /api/themes            列出可用主题
  GET  /api/layouts           列出可用版式（含占位符定义）
  POST /api/generate          提交 deck spec JSON，返回生成的 .pptx 文件
  GET  /healthz                健康检查（部署平台探活用）
  GET  /                       前端静态页面
"""

import io
from datetime import datetime

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from app.engine import render_deck, list_themes, list_layouts, TemplateError

app = FastAPI(title="PPT Template Studio", version="1.0.0")


class SlideSpec(BaseModel):
    layout: str
    content: dict = Field(default_factory=dict)


class DeckSpec(BaseModel):
    theme: str
    slides: list[SlideSpec]


@app.get("/healthz")
def healthz():
    return {"status": "ok"}


@app.get("/api/themes")
def api_themes():
    return list_themes()


@app.get("/api/layouts")
def api_layouts():
    return list_layouts()


@app.post("/api/generate")
def api_generate(deck: DeckSpec):
    try:
        prs = render_deck(deck.model_dump())
    except TemplateError as e:
        raise HTTPException(status_code=400, detail=str(e))

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)

    filename = f"deck_{deck.theme}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=headers,
    )


# 静态前端（放在最后挂载，避免遮蔽 /api/* 路由）
app.mount("/", StaticFiles(directory="app/static", html=True), name="static")
