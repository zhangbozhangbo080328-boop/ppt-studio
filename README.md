# PPT Studio

一个本地运行的 PPT 生成器（网页应用）。输入标题、副标题和内容，选择主题色与风格，点击「生成并下载 PPTX」，即可得到一份原生可编辑的 PowerPoint。

## 功能

- 封面、内容页、结尾页自动排版
- 空行分隔页面，每页第一行为标题，其余行为要点
- 6 种主题色 + 浅色 / 深色封面
- 输出为原生可编辑 `.pptx`（真正的形状与文本，不是图片）
- 全程本地运行，数据不出本机

## 运行

本应用复用已装好的 Python 环境（含 Flask 与 python-pptx），没有则自动创建并安装。在终端执行：

```bash
./run.sh
```

然后浏览器打开 http://localhost:5000。

如果想换端口：

```bash
PORT=8080 ./run.sh
```

## 依赖

- Python 3.10+
- Flask
- python-pptx

## 部署到公网

应用已按生产环境准备好，包含 `requirements.txt`、`Procfile`、`render.yaml` 与 `Dockerfile`。

### 方式 A：Render（推荐，免费）

1. 把本目录推到 GitHub 仓库；
2. 在 [render.com](https://render.com) 新建 Blueprint，选择仓库，Render 会自动读取 `render.yaml` 完成部署；
3. 完成后会得到 `https://ppt-studio.onrender.com` 形式的公网地址。

### 方式 B：Docker 任意平台

```bash
docker build -t ppt-studio .
docker run -p 8000:8000 ppt-studio
```

可部署到 Fly.io、Railway、阿里云、腾讯云等任何支持容器的平台。

### 注意事项

- 免费平台首次访问可能需约 30 秒冷启动；
- 当前为公开生成工具，未加登录鉴权，请勿处理敏感内容。
