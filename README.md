<p align="center">
  <img src="https://img.shields.io/badge/Version-v2.4-blue?style=for-the-badge&logo=python" alt="Version">
  <img src="https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" alt="License">
  <img src="https://img.shields.io/badge/Platform-Windows%2010%2B-lightgrey?style=for-the-badge&logo=windows" alt="Platform">
</p>

<h1 align="center">🧾 电子发票识别工具 — 通用版</h1>

<p align="center">
  <b>一键提取 PDF 发票信息，自动生成 Excel 报销明细表</b>
</p>

<p align="center">
  <sub>🏆 支持文本型 / 图片型 PDF · 智能过滤非发票页面 · 多文件夹批量处理 · 网页操作界面</sub>
</p>

---

## ✨ 为什么选择这个工具？

| 💡 痛点 | ✅ 解决方案 |
|--------|----------|
| 手动抄发票费时费力 | **自动识别** 发票号码、日期、金额、销售方/购买方等全部字段 |
| 扫描件/截图无法复制文字 | **PaddleOCR 深度识别**，支持图片型 PDF |
| 发票夹在报销单里要手动挑 | **智能页面过滤** — 自动跳过报销单、明细表等非发票页面 |
| 发票太多统计容易错 | **一键导出 Excel** — 含明细表、汇总表、报销文件名三个工作表 |
| 需要安装各种库太麻烦 | **一键 setup** → 双击即用，无需懂编程 |

---

## 📸 界面预览

<p align="center">
  <sub>🎨 简洁的网页操作界面，上传 → 识别 → 下载，三步完成</sub>
</p>

```
┌──────────────────────────────────────────────────┐
│  🧾 电子发票识别工具                               │
│  上传 PDF 发票 → 自动识别 → 导出 Excel             │
├──────────────────────────────────────────────────┤
│                                                   │
│   ┌──────────┐          ┌──────────┐              │
│   │  📄 上传  │          │  📁 上传  │              │
│   │ PDF 文件 │          │ PDF 文件夹│              │
│   └──────────┘          └──────────┘              │
│                                                   │
│   ████████████░░░░░░░░  67%                       │
│   识别中 (8/12): 2026年1月话费.pdf                 │
│                                                   │
│   ┌──────────────────────────┐                    │
│   │ 发票号码  │ 日期  │ 金额  │ ...                │
│   │ 251320... │ 01-15 │ 151.9 │ ...               │
│   │ 250678... │ 02-12 │ 187.6 │ ...               │
│   └──────────────────────────┘                    │
│                                                   │
│   [⬇ 下载 Excel]                                  │
└──────────────────────────────────────────────────┘
```

---

## 🚀 快速开始

### 第一次使用（需联网安装环境）

```bash
# 1. 双击运行
setup.cmd
# 自动安装 Python + 依赖，约需 10-20 分钟
```

### 日常使用

```bash
# 2. 双击启动
invoice_tool.bat
# 浏览器自动打开 http://127.0.0.1:5000
```

### 命令行启动（可选）

```bash
pip install -r requirements.txt
python app.py
```

---

## 📊 Excel 输出包含三个工作表

| 工作表 | 内容 | 示例 |
|--------|------|------|
| 📋 **发票明细** | 每张发票的全部字段，可筛选、排序、求和 | 发票号码、开票日期、购买方、销售方、金额、价税合计… |
| 📈 **报销汇总** | 按人名汇总金额和发票张数 | 张三年 — 3 张 — ¥451.70 |
| 📝 **报销文件名** | 自动生成规范的报销文件名 | 张三_2026年1-3月_通讯费_¥451.70 |

---

## 🔍 支持的发票类型

<table>
<tr>
<th>PDF 类型</th><th>识别引擎</th><th>速度</th><th>准确率</th>
</tr>
<tr>
<td>📄 增值税电子发票（文本型）</td><td>pdfplumber</td><td>⚡ 秒级</td><td>99%+</td>
</tr>
<tr>
<td>📱 运营商话费发票（联通/移动/电信）</td><td>pdfplumber + 智能排版解析</td><td>⚡ 秒级</td><td>99%+</td>
</tr>
<tr>
<td>🏥 财政/医疗门诊收费票据（图片型）</td><td>PaddleOCR</td><td>每页约 2-5 秒</td><td>95%+</td>
</tr>
<tr>
<td>🖼️ 其他扫描件 / 图片型 PDF</td><td>PaddleOCR</td><td>🔰 首次需下载模型 (≈500MB)</td><td>95%+</td>
</tr>
</table>

### 🆕 智能页面过滤 (v2.1+)

混在 PDF 里的报销单、明细表、银行卡流水等非发票页面会被**自动识别并跳过**，只提取真正的电子发票或财政票据。

### 🆕 v2.4 识别增强

- **联通等运营商发票**：修复发票号码在标签上方、同行双「名称」等排版问题，文本型 PDF 无需 OCR 即可识别
- **医疗门诊收费票据**：支持浙江省财政电子票据（体检费等），自动提取票据号码、交款人、收款单位、金额
- **OCR 容错**：交款人/开票日期/收款单位等字段增强误识别修复（如 `开东日期`、`收款单住`、`文秋人杨飞玲`）
- **话费字段**：支持 `业务号码`、`账期` 自动提取

---

## 🧰 技术栈

<p align="center">
  <img src="https://img.shields.io/badge/Flask-3.0+-000000?style=flat-square&logo=flask" alt="Flask">
  <img src="https://img.shields.io/badge/PaddleOCR-3.0+-0066CC?style=flat-square&logo=paddlepaddle" alt="PaddleOCR">
  <img src="https://img.shields.io/badge/Waitress-Production-333333?style=flat-square" alt="Waitress">
  <img src="https://img.shields.io/badge/openpyxl-Excel%20Export-217346?style=flat-square&logo=microsoftexcel" alt="openpyxl">
  <img src="https://img.shields.io/badge/PyMuPDF-PDF%20Render-FF6B00?style=flat-square" alt="PyMuPDF">
</p>

| 技术 | 用途 |
|------|------|
| **Flask** | Web 后端服务 |
| **Waitress** | 生产级 WSGI 服务器 |
| **pdfplumber** | 文本型 PDF 内容提取 |
| **PyMuPDF** | PDF 渲染为图片 |
| **PaddleOCR** | 深度学习 OCR 文字识别 |
| **openpyxl** | Excel 生成与格式化 |

---

## 📁 项目结构

```
通用版/
├── app.py                  # Flask Web 主程序
├── invoice_extractor.py    # 核心识别逻辑
├── requirements.txt        # Python 依赖清单
├── invoice_tool.bat        # 日常启动脚本（双击即用）
├── setup.cmd               # 首次环境自动安装
├── setup.bat               # 备用安装脚本
├── 修复OCR缓存.bat          # OCR 异常时一键修复
├── templates/
│   └── index.html          # 网页前端界面
├── uploads/                # 上传文件暂存（自动清理）
├── excel/                  # 识别结果输出
└── 使用说明.md              # 详细使用文档
```

---

## 🔧 系统要求

| 项目 | 最低配置 |
|------|----------|
| 🖥️ 操作系统 | Windows 10 64-bit |
| 💾 内存 | 4 GB |
| 💿 硬盘可用空间 | 3 GB |
| 🌐 网络 | 首次安装需要（下载 Python 和模型） |

---

## ❓ 常见问题

<details>
<summary><b>🔧 PaddleOCR 初始化失败 / 缓存损坏？</b></summary>
<br>
运行 <code>修复OCR缓存.bat</code> 一键修复，然后重新上传 PDF 即可。
</details>

<details>
<summary><b>📦 首次识别图片非常慢？</b></summary>
<br>
首次使用 PaddleOCR 需要自动下载约 500MB 的模型文件，耗时 2-5 分钟。后续识别为秒级。
</details>

<details>
<summary><b>🔤 Excel 打开乱码？</b></summary>
<br>
请使用 <b>Microsoft Excel</b> 打开，不要用 WPS，WPS 对 UTF-8 编码支持不佳。
</details>

<details>
<summary><b>👤 用户名是中文会影响识别吗？</b></summary>
<br>
程序会自动检测中文用户名并将模型缓存重定向到纯英文路径，无需手动操作。
</details>

<details>
<summary><b>🔄 如何升级到最新版？</b></summary>
<br>
<b>从 v2.3 升级到 v2.4</b>（无需重装环境）：
<ol>
<li>关闭正在运行的程序</li>
<li>用新版文件覆盖旧文件夹（至少替换 <code>invoice_extractor.py</code>）</li>
<li>双击 <code>invoice_tool.bat</code> 重新启动</li>
</ol>
<b>从 v2.0 / v2.1 升级</b>：覆盖全部文件后执行：
<pre><code>pip install waitress -i https://pypi.tuna.tsinghua.edu.cn/simple</code></pre>
</details>

---

## 📝 版本历史

| 版本 | 日期 | 更新内容 |
|------|------|----------|
| **v2.4** | 2026-07-07 | 联通发票排版修复、财政/医疗门诊票据支持、OCR 字段容错增强、业务号码/账期提取 |
| **v2.3** | 2026-06-28 | 跨电脑兼容：逐页混合 OCR、CPU/GPU 自动适配、OCR 格式兼容、图片尺寸优化 |
| **v2.2** | 2026-06-27 | Waitress 生产服务器、OCR 自动重试、报销汇总修正、XSS 防护、任务内存管理 |
| **v2.1** | 2026-06-26 | 智能页面过滤：自动跳过报销单/明细表等非发票表单 |
| **v2.0** | 2026-06-25 | 初始发布：文本型+图片型 PDF、多文件夹批量识别、网页版 |

---

## 📄 License

MIT © [涂泊远](https://github.com/boyuan623-coder)

---

<p align="center">
  <sub>如有问题或建议，欢迎提 <a href="https://github.com/boyuan623-coder/invoice-tool/issues">Issue</a> ⭐</sub>
</p>
