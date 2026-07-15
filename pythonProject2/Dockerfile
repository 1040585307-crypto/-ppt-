# 选用精简轻量级的 Python 容器镜像
FROM python:3.9-slim

# 设置容器内部运行工作目录
WORKDIR /app

# 将当前项目目录下的所有源码、前端文件、模板依赖全部拷入容器中
COPY . /app

# 一键装载依赖环境
RUN pip install --no-cache-dir -r requirements.txt

# 宣告容器所对外监听的API通信端口
EXPOSE 8000

# 容器启动运行指令
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]