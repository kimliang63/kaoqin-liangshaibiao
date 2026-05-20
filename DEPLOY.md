# 公网部署说明（Streamlit 免费版）

## 1. 创建代码仓库并上传

1. 打开 GitHub 新建仓库：`https://github.com/new`
2. 把本项目代码推送到仓库，确保包含：
   - `app.py`
   - `generate_report.py`
   - `requirements.txt`
   - `.streamlit/config.toml`

## 2. 发布到 Streamlit Community Cloud

1. 打开控制台：`https://share.streamlit.io/`
2. 选择 `New app`
3. 配置：
   - Repository：你的 GitHub 仓库
   - Branch：`main`（或你的默认分支）
   - Main file path：`app.py`
4. 点击 `Deploy`

## 3. 公网访问地址

发布成功后将自动生成公网 URL，格式：

- `https://<your-app-name>.streamlit.app`

例如（示例）：

- `https://attendance-report-demo.streamlit.app`

## 4. 交付验收

- 上传 4 个 Excel 后可成功生成报表
- 页面可在线预览报表
- 可下载 HTML 与 JSON
