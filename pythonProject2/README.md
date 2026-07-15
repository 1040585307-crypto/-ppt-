# AI一键生成PPT演示文稿 - 系统交付与部署文档



## 1. 系统架构与技术栈规范
* **前端 (Frontend):** 原生 HTML5 + CSS3 + ES6 JavaScript。极致轻量，响应式支持，完全兼容最新的 Chrome, Edge, Safari 浏览器。
* **后端 (Backend):** Python 3.9 + FastAPI。依靠异步高性能 ASGI 架构，满足 10 个以上并发用户平滑调度。
* **排版渲染层:** 基于开源 `python-pptx` 构建自动布局，不含任何未商用授权资产，字体、版式均开源。
* **AI 驱动管道:** 依托国际版 Coze 工作流 (基于 GPT-4o 系列模型)。

## 2. 部署指南

### 本地部署 (Windows / Linux / macOS)
1. 确保安装好 Python 3.9+ 环境。
2. 在项目根目录下打开终端，装载包依赖：
   ```bash
   pip install -r requirements.txt
