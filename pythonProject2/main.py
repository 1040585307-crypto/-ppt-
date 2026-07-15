import json
import re
import os
import tempfile
import time
from typing import Dict

import requests
from fastapi import FastAPI, HTTPException, BackgroundTasks, Request
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# 引入你的 PPT 排版引擎
from ppt_engine import create_ppt_from_json

# ================= 凭证配置 =================
COZE_TOKEN = ""
WORKFLOW_ID = ""
COZE_API_URL = "https://api.coze.com/v1/workflow/run"

app = FastAPI(title="AI PPT 生成引擎API")

# 跨域设置，兼容主流浏览器请求
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class TopicReq(BaseModel):
    topic: str


# ================= 任务书合规性：安全/合规/拦截拦截器 =================
IP_TRACKER = {}
COOLDOWN_SECONDS = 15  # 限流：同IP 15秒内仅允许生成一次
SENSITIVE_WORDS = ["暴力", "赌博", "毒品", "枪支", "非法", "测试敏感词"]  # 敏感词库


@app.post("/generate_ppt")
def generate_ppt(req: TopicReq, background_tasks: BackgroundTasks, request: Request):
    topic = req.topic.strip()

    # 1. 安全合规检验：敏感词过滤拦截
    for word in SENSITIVE_WORDS:
        if word in topic:
            raise HTTPException(status_code=400, detail=f"生成请求已被拦截：检测到敏感/违规词汇 [{word}]")

    # 2. 访问频率控制：防滥用限流拦截
    client_ip = request.client.host
    current_time = time.time()
    if client_ip in IP_TRACKER:
        time_passed = current_time - IP_TRACKER[client_ip]
        if time_passed < COOLDOWN_SECONDS:
            raise HTTPException(status_code=429,
                                detail=f"系统保护中：请等待 {int(COOLDOWN_SECONDS - time_passed)} 秒后再试")
    IP_TRACKER[client_ip] = current_time

    # 3. 构建请求发送至 Coze 国际版的工作流
    headers = {
        "Authorization": f"Bearer {COZE_TOKEN}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    body = {
        "workflow_id": WORKFLOW_ID,
        "parameters": {
            "input": topic
        }
    }

    try:
        resp = requests.post(COZE_API_URL, headers=headers, json=body, timeout=90, verify=False)
    except requests.RequestException as e:
        raise HTTPException(status_code=502, detail=f"调用 AI 管道服务失败: {e}")

    if resp.status_code != 200:
        raise HTTPException(status_code=502, detail=f"AI 管道异常：{resp.status_code}")

    # 4. 解析大模型传回的结构化 JSON
    try:
        resp_json = resp.json()
        if resp_json.get("code") != 0:
            raise ValueError(resp_json.get("msg", "业务异常"))

        data_str = resp_json.get("data", "{}")
        data_dict = json.loads(data_str)
        output_str = data_dict.get("output", "{}")

        # 剥离外部可能包夹的 Markdown 标记
        m = re.search(r"\{.*\}", output_str, re.S)
        if not m:
            raise ValueError("未能解出合法的 JSON 报文格式")

        final_json_data = json.loads(m.group(0))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"AI 数据解析失败: {e}")

    # 5. 任务书合规：即用即销，使用临时文件
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    tmp_path = tmp.name
    tmp.close()

    try:
        create_ppt_from_json(final_json_data, tmp_path)
    except Exception as e:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise HTTPException(status_code=500, detail=f"PPT 自动排版渲染失败: {e}")

    # 数据安全：生成后台清理任务，文件下载顺着网线过去后，立即将本地文件物理删除
    def _cleanup(path: str):
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception:
            pass

    background_tasks.add_task(_cleanup, tmp_path)

    return FileResponse(
        tmp_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="最新版_AI_PPT.pptx",
        background=background_tasks
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)