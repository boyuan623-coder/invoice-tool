<p align="center">
  <img src="https://img.shields.io/badge/Version-v2.5-blue?style=for-the-badge&logo=python" alt="Version">
  <img src="https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/Deploy-Tencent%20Lighthouse-00A4FF?style=for-the-badge&logo=tencentqq" alt="Deploy">
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" alt="License">
</p>

<h1 align="center">电子发票识别工具</h1>

<p align="center">
  <b>上传 PDF / 图片 → 自动识别 → 一键导出 Excel 报销明细</b>
</p>

<p align="center">
  文本发票秒级提取 · 扫描件 OCR · 智能跳过非发票页 · 表格转 Excel · 本机 / 分享 / 云服务器一键部署
</p>

<p align="center">
  <a href="#-快速开始本机">快速开始</a> ·
  <a href="#-三种使用方式">使用方式</a> ·
  <a href="#-腾讯云轻量部署">云部署</a> ·
  <a href="#-功能说明">功能说明</a> ·
  <a href="./使用说明.md">详细使用说明</a>
</p>

---

## ✨ 功能一览

| 场景 | 能力 |
|------|------|
| 报销录入 | 自动提取发票号码、日期、购销方、价税合计等字段 |
| 扫描件 / 截图 | PaddleOCR 识别图片型 PDF 与 JPG/PNG |
| 混装报销材料 | 自动跳过报销单、明细表等非发票页面 |
| 批量处理 | 多文件 / 多文件夹上传，合并导出 |
| 表格截图 | 「表格转 Excel」模式，每张图生成同名表格 |
| 对外使用 | 本机、临时外网链接、腾讯云轻量常驻部署 |
| 安全访问 | 可选访问口令（公网部署强烈建议开启） |

---

## 🚀 快速开始（本机）

> Windows 10/11，首次需联网安装环境。

### 1. 安装（仅第一次）

双击 `setup.cmd`，等待出现 `Setup complete!`（约 10–20 分钟）。

### 2. 日常使用

双击 `invoice_tool.bat` → 浏览器打开 `http://127.0.0.1:5000` → 上传 → 识别 → 下载 Excel。

### 3. 命令行（可选）

```bash
pip install -r requirements.txt
python app.py
```

更细的步骤、排错与隧道说明见 **[使用说明.md](./使用说明.md)**。

---

## 🌐 三种使用方式

| 方式 | 怎么启动 | 适用 |
|------|----------|------|
| **本机** | `invoice_tool.bat` | 自己用 |
| **临时分享** | `invoice_tool_share.bat` 或 `python app.py --share` | 发给同事临时链接（关闭窗口即失效） |
| **云服务器** | `bash deploy/start_cloud.sh` | 腾讯云轻量等，公网长期访问 |

```bash
# 本机
python app.py

# 临时公网链接（Cloudflare 隧道）
python app.py --share

# 云服务器（不弹浏览器，可配合口令与低内存模式）
python app.py --cloud --low-mem --host 0.0.0.0 --port 5000
```

---

## ☁️ 腾讯云轻量部署

适合把服务挂在公网，给其他人浏览器访问。

### 控制台

1. 开通 **轻量应用服务器**（建议 ≥ 2 核 4G；2G 可用但图片 OCR 较慢）
2. **防火墙** 放行 TCP `5000`
3. 用 SSH 登录，上传本仓库到例如 `~/invoice-tool`

### 服务器命令

```bash
cd ~/invoice-tool
chmod +x deploy/*.sh

# 2G 机器强烈建议先加 Swap
sudo bash deploy/setup_swap.sh

# 配置口令与低内存
cp deploy/env.example deploy/invoice-tool.env
nano deploy/invoice-tool.env
# 修改：INVOICE_TOOL_ACCESS_TOKEN=你的口令
# 确认：INVOICE_TOOL_LOW_MEM=1

# 安装依赖并后台启动
bash deploy/start_cloud.sh --daemon
```

浏览器访问：`http://你的公网IP:5000`，按提示输入口令。

### 开机自启（可选）

```bash
# 按实际路径/用户编辑 service 文件后：
sudo cp deploy/invoice-tool.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now invoice-tool
```

| 环境变量 | 说明 |
|----------|------|
| `INVOICE_TOOL_ACCESS_TOKEN` | 访问口令（公网务必设置） |
| `INVOICE_TOOL_LOW_MEM=1` | 2G 小机：跳过 OCR 预热、限制并发 |
| `INVOICE_TOOL_CLOUD=1` | 云模式 |
| `INVOICE_TOOL_PORT` | 端口，默认 `5000` |

> 文本型电子发票通常很快；多页扫描件 OCR 在 2G CPU 上可能需要数分钟，属正常现象。

---

## 📊 功能说明

### Excel 输出（三个工作表）

| 工作表 | 内容 |
|--------|------|
| **发票明细** | 号码、日期、购销方、金额、价税合计、手机号、账期等 |
| **报销汇总** | 按人员汇总张数与金额 |
| **报销文件名** | 自动生成规范报销命名建议 |

### 支持的发票类型

| 类型 | 引擎 | 速度 |
|------|------|------|
| 增值税电子发票（文本） | pdfplumber | 秒级 |
| 运营商话费发票 | 文本解析 | 秒级 |
| 财政 / 医疗票据（图片） | PaddleOCR | 约数秒/页 |
| 其他扫描件 PDF / 图片 | PaddleOCR | 首次需下载模型 ≈500MB |

### 识别模式

页面可选：**速度 / 平衡 / 精度**。云上小规格机器建议优先「速度」。

### 表格转 Excel

界面切换到「表格转 Excel」，上传表格截图，按图生成同名 Excel。

---

## 🧰 技术栈

<p align="center">
  <img src="https://img.shields.io/badge/Flask-3.0+-000000?style=flat-square&logo=flask" alt="Flask">
  <img src="https://img.shields.io/badge/PaddleOCR-3.0+-0066CC?style=flat-square&logo=paddlepaddle" alt="PaddleOCR">
  <img src="https://img.shields.io/badge/Waitress-Production-333333?style=flat-square" alt="Waitress">
  <img src="https://img.shields.io/badge/openpyxl-Excel-217346?style=flat-square&logo=microsoftexcel" alt="openpyxl">
  <img src="https://img.shields.io/badge/PyMuPDF-PDF-FF6B00?style=flat-square" alt="PyMuPDF">
</p>

---

## 📁 项目结构

```text
invoice-tool/
├── app.py                  # Web 服务入口（支持 --cloud / --share / --low-mem）
├── invoice_extractor.py    # 发票识别核心
├── table_extractor.py      # 表格转 Excel
├── requirements.txt        # 依赖（CPU）
├── requirements-gpu.txt    # GPU 版说明
├── invoice_tool.bat        # 本机启动
├── invoice_tool_share.bat  # 临时外网分享
├── setup.cmd / setup.bat   # 首次环境安装
├── deploy/                 # 云部署脚本与 systemd / Nginx 示例
│   ├── start_cloud.sh
│   ├── setup_swap.sh
│   ├── env.example
│   └── invoice-tool.service
├── templates/index.html    # 前端页面
├── docs/DESIGN_v3.md       # 设计说明
└── 使用说明.md              # 完整使用手册
```

---

## 🔧 环境要求

| 项目 | 建议 |
|------|------|
| 本机系统 | Windows 10/11 64-bit |
| 本机内存 | ≥ 4 GB（OCR 更稳） |
| 云服务器 | Ubuntu 22.04+，建议 2 核 4G；2G 需开 Swap + 低内存模式 |
| 磁盘 | ≥ 3 GB（含 OCR 模型） |
| 网络 | 首次安装与下载模型需要 |

---

## ❓ 常见问题

<details>
<summary><b>PaddleOCR 初始化失败？</b></summary>
<br>
Windows 可运行 <code>修复OCR缓存.bat</code> 后重试。
</details>

<details>
<summary><b>首次识别图片很慢？</b></summary>
<br>
首次需下载约 500MB 模型，之后会快很多。多页扫描件按「约数秒～十余秒/页」估算。
</details>

<details>
<summary><b>云上页面打不开？</b></summary>
<br>
检查轻量防火墙是否放行 5000；服务器上执行 <code>curl http://127.0.0.1:5000/api/health</code> 与 <code>tail -f app.log</code>。
</details>

<details>
<summary><b>前端进度好像不动，但日志在跑？</b></summary>
<br>
多页 OCR 时后端可能已到第 N 页，页面文案刷新可能滞后；以服务器 <code>app.log</code> 为准，完成后可下载 Excel。
</details>

<details>
<summary><b>Excel 乱码？</b></summary>
<br>
请用 Microsoft Excel 打开；部分 WPS 对编码支持不佳。
</details>

更多说明见 **[使用说明.md](./使用说明.md)**。

---

## 📝 版本历史

| 版本 | 日期 | 更新 |
|------|------|------|
| **v2.5** | 2026-07 | 云部署脚本、访问口令、低内存模式、表格转 Excel、分享隧道完善 |
| **v2.4** | 2026-07-07 | 联通发票 / 医疗门诊票据增强、OCR 字段容错 |
| **v2.3** | 2026-06-28 | 逐页混合 OCR、CPU/GPU 适配 |
| **v2.2** | 2026-06-27 | Waitress、任务管理与安全加固 |
| **v2.1** | 2026-06-26 | 智能页面过滤 |
| **v2.0** | 2026-06-25 | 网页版初始发布 |

---

## 📄 License

MIT © [涂泊远](https://github.com/boyuan623-coder)

<p align="center">
  <sub>问题与建议欢迎提 <a href="https://github.com/boyuan623-coder/invoice-tool/issues">Issue</a> · 如果觉得有用欢迎 Star</sub>
</p>
