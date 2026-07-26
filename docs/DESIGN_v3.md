# 文档识别与 Excel 转换工具 — v3.0 设计文档

> 版本：Draft 1.0  
> 日期：2026-07-07  
> 状态：待评审  
> 基于：通用版 v2.4（发票识别）扩展

---

## 1. 背景与目标

### 1.1 现状（v2.4）

当前工具是**面向报销场景的发票识别工具**，核心能力：

- 支持 PDF（文本型 / 图片型）及 JPG/PNG 等图片
- 逐页混合识别（pdfplumber 优先，OCR 兜底）
- 提取发票/财政票据固定字段，导出三 Sheet Excel（明细、汇总、报销文件名）
- Web 上传 + 批量文件夹 + 嵌套目录扫描

**局限：**

- 输出列固定，用户不可选
- 仅「发票模式」，无通用表格转换
- PDF 内表格未作为结构化数据一等公民处理
- 失败文件无集中反馈
- 多子文件夹时合并策略不够灵活
- 无「每张图片单独 Excel」选项

### 1.2 v3.0 目标

在**不破坏现有发票模式**的前提下，升级为**多模式文档识别与 Excel 转换平台**：

| 目标 | 说明 |
|------|------|
| 模式化 | 发票 / 图片转 Excel / 通用模式 三轨并行 |
| 表格增强 | PDF + 图片的表格结构识别与还原 |
| 可配置导出 | 用户勾选列、排序，动态生成 Excel |
| 类型化汇总 | 按文档类型生成统计 Sheet（可选） |
| 可追溯 | 失败文件名单 + 转换成功率 |
| 批量增强 | 深层嵌套文件夹 + 合并/拆分导出策略 |

### 1.3 非目标（v3.0 不做）

- 不迁移到 Java 重写
- 不做云端 SaaS / 多租户权限
- 不做发票验真（税务局接口）
- 不做全自动「猜文档类型」作为主路径（采用用户预选类型）
- 不承诺第一期覆盖所有文档类型

---

## 2. 需求清单与优先级

| ID | 需求 | 优先级 | 阶段 |
|----|------|--------|------|
| R7 | 发票模式保持现有功能不变 | P0 | 0 |
| R5 | 未识别文件单独展示 | P0 | 1 |
| R6 | 嵌套文件夹批量 + 合并/分文件夹导出 | P0 | 1 |
| R8 | 图片模式：每张图是否单独 Excel + 输出目录结构 | P0 | 1 |
| R3 | 三模式选择（发票 / 图片转 Excel / 通用） | P0 | 1 |
| R1 | PDF 表格识别升级为图片 + 表格识别 | P1 | 2 |
| R3b | 图片转 Excel 模式汇报转换成功率 | P1 | 2 |
| R2 | 自定义 Excel 列（可选列来自识别字段） | P1 | 3 |
| R4 | 通用模式：类型选择 + 可选汇总 Sheet | P1 | 3 |
| R4b | 成绩表（平均、取整）、工资表（平均、两位小数） | P2 | 3 |

---

## 3. 技术选型

### 3.1 结论：继续 Python

| 维度 | Python | Java |
|------|--------|------|
| 现有资产 | v2.4 完整可复用 | 需全部重写 |
| OCR/PDF 生态 | PaddleOCR、pdfplumber、PyMuPDF 成熟 | 表格 OCR 方案弱，集成成本高 |
| Excel | openpyxl 足够 | Apache POI 可行 |
| Web | Flask → 可演进 FastAPI | Spring Boot 过重 |
| 交付周期 | 8～14 周 | 15～25 周（含迁移） |

**架构原则：** 识别引擎、导出引擎、Web 层解耦；未来若需 Java 网关，Python 识别服务可独立部署。

### 3.2 依赖规划

| 组件 | 用途 | 变更 |
|------|------|------|
| pdfplumber | PDF 文本 + 表格 | 保留，表格路径加强 |
| PyMuPDF | PDF 渲染 | 保留 |
| PaddleOCR | 文字 + 表格识别 | 评估启用 Table 系列模型 |
| openpyxl | Excel 读写 | 保留，扩展动态列 |
| Flask / Waitress | Web | 保留 v3.0，接口扩展 |

---

## 4. 系统架构

### 4.1 分层架构

```
┌─────────────────────────────────────────────────────────┐
│  表现层  templates/index.html  +  REST API (app.py)    │
├─────────────────────────────────────────────────────────┤
│  任务层  TaskManager / 进度回调 / 批次与输出策略          │
├─────────────────────────────────────────────────────────┤
│  模式层  ModeHandler (策略模式)                          │
│    ├── InvoiceModeHandler      ← 封装现有 v2.4 逻辑      │
│    ├── Image2ExcelModeHandler  ← 表格 OCR + 成功率       │
│    └── GeneralModeHandler      ← 类型 + 可选列 + 汇总    │
├─────────────────────────────────────────────────────────┤
│  识别层  Recognizer                                        │
│    ├── TextExtractor (pdfplumber)                        │
│    ├── OcrEngine (PaddleOCR 单例)                        │
│    └── TableExtractor (pdfplumber + OCR Table)           │
├─────────────────────────────────────────────────────────┤
│  导出层  ExcelExporter                                     │
│    ├── FixedTemplateExporter   (发票三 Sheet)             │
│    ├── DynamicColumnExporter   (用户选列)                 │
│    └── TableSheetExporter      (纯表格还原)               │
└─────────────────────────────────────────────────────────┘
```

### 4.2 目录结构（目标）

```
通用版/
├── app.py                      # Web 入口，路由扩展
├── invoice_extractor.py        # 保留，发票模式底层（逐步瘦身）
├── core/
│   ├── __init__.py
│   ├── models.py               # RecognitionResult, TaskConfig, ...
│   ├── discovery.py            # 文件扫描（PDF/图片/嵌套文件夹）
│   ├── task_runner.py          # 统一任务编排
│   └── modes/
│       ├── base.py             # ModeHandler 抽象基类
│       ├── invoice.py          # 发票模式（Adapter → v2.4）
│       ├── image2excel.py      # 图片/PDF → 表格 Excel
│       └── general.py          # 通用模式
├── recognizers/
│   ├── ocr_engine.py           # PaddleOCR 封装（从 invoice_extractor 抽出）
│   ├── table_extractor.py      # 表格识别
│   └── field_extractors/       # 按文档类型的字段提取
│       ├── invoice.py
│       ├── grade_sheet.py
│       ├── payroll.py
│       └── financial.py
├── exporters/
│   ├── invoice_exporter.py     # 现有三 Sheet 逻辑
│   ├── dynamic_exporter.py     # 可选列导出
│   └── summary_builders/       # 各类型汇总 Sheet
├── templates/
│   └── index.html              # 多模式 UI
├── docs/
│   └── DESIGN_v3.md            # 本文档
└── tests/
    └── samples/                # 各类型测试样本（gitignore 大文件）
```

> **迁移策略：** 阶段 0 新建 `core/`，发票模式通过 Adapter 调用现有 `invoice_extractor.py`，保证零回归；后续逐步搬迁 OCR/导出逻辑。

---

## 5. 三种模式设计

### 5.1 模式总览

```
用户上传文件
    ↓
选择模式 ─────────────────────────────────────────────
    │                    │                          │
    ▼                    ▼                          ▼
发票模式            图片转 Excel 模式            通用模式
(Invoice)           (Image2Excel)               (General)
    │                    │                          │
现有字段提取          表格结构还原              预选文档类型
三 Sheet 固定输出     成功率统计                用户选列
                      可选每文件一 Excel        可选汇总 Sheet
```

### 5.2 模式 A：发票模式（Invoice）

**原则：行为与 v2.4 完全一致，代码路径隔离。**

| 项 | 说明 |
|----|------|
| 输入 | PDF、图片（jpg/png/...） |
| 识别 | 调用现有 `extract_all_invoices()` |
| 输出 | 发票明细 + 报销汇总 + 报销文件名 |
| 失败 | 纳入统一 `failed_files` 列表（v3 新增） |
| 批量 | 支持 R6 文件夹策略（不影响字段逻辑） |

**不允许的变更：** 字段含义、Sheet 名称、汇总算法、发票判定规则（除非明确 bugfix）。

### 5.3 模式 B：图片 / PDF → 纯 Excel（Image2Excel）

**目标：** 将图片或 PDF 中的**表格**还原为 Excel 行列数据。

| 项 | 说明 |
|----|------|
| 输入 | 图片、PDF（每页视为一张表或按检测拆分） |
| 识别流程 | ① 渲染/读图 → ② 表格检测 → ③ 单元格 OCR → ④ 结构化 `rows[][]` |
| 输出选项 | 见 5.3.1 |
| 成功率 | 必须汇报（见 5.3.2） |

#### 5.3.1 输出策略

| 策略 | 行为 |
|------|------|
| `merge_single` | 所有文件合并为一个 Excel（每源文件/页一个 Sheet） |
| `one_file_one_excel` | 每个源文件一个 Excel |
| `folder_mirror` | 输出目录镜像输入文件夹结构；同文件夹下多个 Excel 放在对应子目录 |

**R8 实现要点：**

- 用户在上传前或识别前选择输出策略
- 输出根目录：`excel/{batch_id}/` 或 `excel/{batch_id}/{相对路径}/`

#### 5.3.2 成功率汇报

任务完成后返回：

```json
{
  "total_files": 20,
  "success_count": 17,
  "failed_count": 3,
  "success_rate": 0.85,
  "failed_files": [
    { "file_name": "scan_03.jpg", "reason": "未检测到表格结构" },
    { "file_name": "blur.png", "reason": "OCR 无有效文本" }
  ]
}
```

前端单独区域展示失败名单；Excel 可选增加「转换报告」Sheet。

### 5.4 模式 C：通用模式（General）

**目标：** 用户预选文档类型 → 识别字段 → 勾选导出列 → 可选汇总 Sheet。

#### 5.4.1 文档类型（第一期）

| type_id | 名称 | 默认可选字段 | 可选汇总 |
|---------|------|-------------|----------|
| `invoice` | 发票 | 同发票模式 | 按人汇总金额（同 v2.4） |
| `grade_sheet` | 成绩表 | 姓名、学号、科目、分数、排名… | 总分、平均分（**取整**） |
| `payroll` | 工资表 | 姓名、工号、应发、扣款、实发… | 合计、平均（**保留 2 位小数**） |
| `financial` | 财务报表 | 科目、摘要、借方、贷方、余额… | 收入/支出合计 |
| `generic_table` | 通用表格 | 动态列（表头 OCR 结果） | 无（或仅行数统计） |
| `other` | 其他 | 全部 OCR 文本块 | 无 |

> 类型「全面」为长期目标；v3.0 MVP 先上线上表 4～5 种。

#### 5.4.2 交互流程

```
1. 选择「通用模式」
2. 选择文档类型（下拉）
3. 上传文件/文件夹
4. [可选] 选择文件夹导出策略（R6）
5. 识别中…
6. 识别完成 → 展示：
   - 检测到的字段列表（checkbox，默认全选）
   - 列顺序（可选拖拽，P2）
   - 「是否生成汇总 Sheet？」（按类型显示不同选项）
7. 确认 → 生成 Excel
8. 展示成功/失败统计 + 失败文件名
```

#### 5.4.3 汇总 Sheet 规则

| 类型 | 汇总内容 | 数值格式 |
|------|----------|----------|
| 成绩表 | 每人总分、全班平均分 | 平均：**四舍五入取整** |
| 工资表 | 应发/实发合计、平均值 | 平均：**保留 2 位小数** |
| 发票 | 按购买方汇总金额、张数 | 金额 2 位小数 |
| 财务报表 | 借方/贷方合计 | 2 位小数 |

---

## 6. 核心数据模型

### 6.1 TaskConfig（任务配置）

```python
@dataclass
class TaskConfig:
    mode: Literal["invoice", "image2excel", "general"]
    batch_id: str
    input_dir: str
    output_dir: str

    # R6 文件夹策略
    folder_export: Literal["merge_one", "per_folder", "per_file"]

    # R8 图片模式输出
    image_output: Literal["merge_single", "one_file_one_excel", "folder_mirror"]

    # 通用模式
    doc_type: Optional[str] = None          # grade_sheet, payroll, ...
    selected_columns: Optional[list[str]] = None
    enable_summary: bool = False
    summary_options: Optional[dict] = None
```

### 6.2 FileResult（单文件结果）

```python
@dataclass
class FileResult:
    file_name: str
    source_folder: str
    status: Literal["success", "failed", "skipped"]
    reason: Optional[str] = None

    # 发票/通用字段模式
    fields: Optional[dict[str, Any]] = None

    # 表格模式
    tables: Optional[list[list[list[str]]]] = None  # [table][row][col]

    extraction_method: Optional[str] = None  # pdfplumber / OCR / table
```

### 6.3 BatchResult（批次结果）

```python
@dataclass
class BatchResult:
    task_id: str
    mode: str
    total_files: int
    success_count: int
    failed_count: int
    success_rate: float
    file_results: list[FileResult]
    output_files: list[str]       # 生成的 Excel 路径
    invoices: list[dict]          # 兼容现有前端（发票模式）
```

### 6.4 与现有 InvoiceInfo 的关系

- `InvoiceInfo` **保留**，仅发票模式 / 通用模式-发票类型使用
- 新模式逐步转向 `FileResult.fields` 字典，导出层做适配
- 避免一次性删除 `InvoiceInfo` 导致回归

---

## 7. 表格识别方案（R1）

### 7.1 双路径策略

```
PDF 文件
  ├── 路径 A：pdfplumber.extract_tables()  → 有结构化表格 → 直接使用
  └── 路径 B：渲染为图片 → PaddleOCR 表格模型 → 单元格坐标 + 文字

图片文件
  └── 路径 B only
```

### 7.2 表格质量评估

对每个识别结果计算 `table_confidence`：

| 信号 | 权重 |
|------|------|
| 检测到表格边框/结构 | 高 |
| 行列数 ≥ 2×2 | 中 |
| 单元格非空率 | 中 |
| 表头含中文/数字混合 | 低 |

低于阈值 → 标记 `failed`，原因写入 `failed_files`。

### 7.3 降级

- 表格识别失败时，可选降级为「纯文本 Sheet」（OCR 全文），不算 success，计入 partial（P2）

---

## 8. 文件扫描与批量（R6）

### 8.1 目录发现

扩展现有 `_discover_folders()`：

```
input/
├── 张三/
│   ├── a.pdf
│   └── b.jpg
├── 李四/
│   └── 话费/
│       └── c.pdf
└── root.pdf
```

返回：

```python
{
  ".": ["root.pdf"],
  "张三": ["a.pdf", "b.jpg"],
  "李四/话费": ["c.pdf"]
}
```

### 8.2 导出策略

| 策略 | 输出 |
|------|------|
| `merge_one` | 一个 Excel，每来源文件夹一个 Sheet（或全部明细一个 Sheet + 来源列） |
| `per_folder` | 每个一级/叶子文件夹一个 Excel |
| `per_file` | 每个源文件一个 Excel（配合 R8） |

**前端：** 上传后弹窗或分步向导让用户选择（默认可记住上次选择）。

---

## 9. API 设计

### 9.1 现有接口（保留）

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/` | 页面 |
| POST | `/api/upload-files` | 上传（扩展参数） |
| POST | `/api/upload-folder` | 上传文件夹 |
| GET | `/api/status/{task_id}` | 轮询状态 |
| GET | `/api/download/{task_id}` | 下载（多文件时 zip，P2） |

### 9.2 新增 / 扩展参数

**上传时（FormData）：**

| 字段 | 类型 | 说明 |
|------|------|------|
| `mode` | string | invoice / image2excel / general |
| `folder_export` | string | merge_one / per_folder / per_file |
| `image_output` | string | merge_single / one_file_one_excel / folder_mirror |
| `doc_type` | string | 通用模式文档类型 |
| `selected_columns` | JSON string | 通用模式列选择 |
| `enable_summary` | bool | 是否汇总 Sheet |

**通用模式两阶段（可选）：**

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/preview/{task_id}` | 识别完成，返回可用字段列表 |
| POST | `/api/export/{task_id}` | 用户确认列后正式导出 |

> 两阶段可避免「识别完才发现列不对要重来」；MVP 也可一步完成（默认全列）。

### 9.3 status 响应扩展

```json
{
  "status": "done",
  "mode": "image2excel",
  "progress": 100,
  "success_rate": 0.85,
  "failed_files": [
    { "file_name": "x.jpg", "reason": "未检测到表格" }
  ],
  "available_columns": ["姓名", "金额"],
  "output_files": ["发票识别结果_xxx.xlsx"]
}
```

---

## 10. 前端 UI 设计

### 10.1 页面结构

```
┌─ 步骤 1：选择模式 ─────────────────────────────┐
│  ○ 发票模式   ○ 图片转 Excel   ○ 通用模式      │
└──────────────────────────────────────────────┘
┌─ 步骤 2：模式相关选项 ─────────────────────────┐
│  [发票] 无额外选项                              │
│  [图片] 输出方式：合并 / 每文件一 Excel / 镜像目录 │
│  [通用] 文档类型下拉 + 汇总勾选                  │
└──────────────────────────────────────────────┘
┌─ 步骤 3：上传 ─────────────────────────────────┐
│  上传文件 / 上传文件夹                          │
│  文件夹导出：合并一个表 / 每文件夹一个表          │
└──────────────────────────────────────────────┘
┌─ 步骤 4：结果 ─────────────────────────────────┐
│  成功率、统计卡片                               │
│  ⚠ 失败文件列表（单独高亮区域）                  │
│  数据预览表格                                   │
│  [通用] 列选择面板                              │
│  下载 Excel / 下载 ZIP                          │
└──────────────────────────────────────────────┘
```

### 10.2 失败文件展示（R5）

- 独立红色/橙色区域，列表展示 `文件名 + 失败原因`
- 支持复制名单
- 发票模式与新模式统一组件

---

## 11. 发票模式隔离方案（R7 保障）

```python
class InvoiceModeHandler(ModeHandler):
    def run(self, config: TaskConfig) -> BatchResult:
        # 仅调用 invoice_extractor.extract_all_invoices
        # 仅调用 invoice_exporter.write_to_excel
        # 不经过 TableExtractor / DynamicColumnExporter
        ...
```

**回归测试集：** 固定 v2.4 测试 PDF（联通发票、医疗票据、16 页体检）必须通过。

---

## 12. 实施计划

### 阶段 0：架构准备（1～1.5 周）

| 任务 | 产出 |
|------|------|
| 创建 `core/` 模块骨架 | models, modes/base, task_runner |
| InvoiceModeHandler Adapter | 发票模式零变更接入 |
| 统一 BatchResult / failed_files | 后端数据结构 |
| 单元测试框架 | tests/ + 发票回归样本 |

**验收：** 发票模式输出与 v2.4 字节级一致（字段级一致即可）。

### 阶段 1：体验增强（1.5～2 周）

| 任务 | 需求 |
|------|------|
| 模式选择 UI | R3 骨架 |
| 失败文件名单 | R5 |
| 文件夹导出策略 | R6 |
| 图片输出策略 | R8 |
| status API 扩展 | 成功率 |

**验收：** 三模式可选；发票模式不变；失败名单可见；嵌套文件夹可合并/拆分导出。

### 阶段 2：表格识别（2～3 周）

| 任务 | 需求 |
|------|------|
| TableExtractor 双路径 | R1 |
| Image2ExcelModeHandler | R3 图片模式 |
| 转换成功率报告 | R3b |
| 表格质量评估与失败原因 | R5 增强 |

**验收：** 规则清晰表格（有边框）识别率 ≥ 85%；成功率正确统计。

### 阶段 3：通用模式 MVP（3～4 周）

| 任务 | 需求 |
|------|------|
| 文档类型注册表 | R4 |
| grade_sheet / payroll 字段提取 | R4b |
| 可选列 UI + DynamicExporter | R2 |
| SummaryBuilder 按类型 | R4 |
| 两阶段 preview/export（可选） | R2 体验 |

**验收：** 成绩表、工资表样本可导出 + 汇总 Sheet 格式正确。

### 阶段 4： polish（1～2 周，可选）

- 列配置保存
- 多 Excel 打包 ZIP 下载
- 识别预览与手动修正
- 更多文档类型模板

---

## 13. 风险与对策

| 风险 | 影响 | 对策 |
|------|------|------|
| 复杂表格 OCR 不准 | 图片模式成功率低 | 分类型调参；明确告知适用场景；保留失败名单 |
| 范围膨胀 | 延期 | 严格分阶段；通用类型分批上线 |
| 发票回归破坏 | 用户投诉 | Adapter 隔离 + 固定回归集 |
| OCR 性能 | 大批量慢 | 进度细化；CPU 降 DPI；可选 GPU |
| 前端复杂度 | 交互混乱 | 分步向导；模式间 UI 条件渲染 |

---

## 14. 测试策略

| 类型 | 内容 |
|------|------|
| 回归 | v2.4 发票样本（联通、医疗、混合 PDF） |
| 表格 | 5～10 份/类型的扫描表格样本 |
| 批量 | 3 层嵌套文件夹 + 混合 PDF/图片 |
| 失败 | 故意模糊图、非表格图，验证 failed_files |
| 导出 | 各策略下 Excel 文件数量与目录结构 |

---

## 15. 开放问题（评审确认）

1. **「Excel 图片」** 是否指 xlsx 内嵌图片？若是，需增加 openpyxl 抽图流程。
2. **通用模式** 第一期上几种类型？建议：发票、成绩表、工资表、generic_table。
3. **合并 Excel** 时 Sheet 命名规则？（文件夹名 / 文件名 / 自动截断 31 字符）
4. **多 Excel 下载** 第一期是否 ZIP？还是只下载合并的一个？
5. **识别预览** 是否必须两阶段？MVP 可一步导出，P2 再加 preview。

---

## 16. 附录：与 v2.4 文件映射

| v2.4 文件 | v3.0 命运 |
|-----------|-----------|
| `invoice_extractor.py` | 保留；发票模式底层；逐步抽 OCR → `recognizers/` |
| `app.py` | 扩展路由；任务编排迁到 `core/task_runner.py` |
| `templates/index.html` | 重写为多步向导 |
| `使用说明.md` | 按模式分章节更新 |
| `requirements.txt` | 阶段 2 可能增加 paddle table 相关依赖 |

---

## 17. 总结

- **技术路线：** Python 演进，不迁 Java  
- **核心架构：** 模式策略 + 统一结果模型 + 可插拔导出  
- **交付节奏：** 约 2 周可用 v3.0 beta（阶段 0+1）→ 1 月表格能力 → 2～3 月通用模式 MVP  
- **铁律：** 发票模式功能与输出保持不变  

---

*文档结束 — 请评审「开放问题」后进入阶段 0 开发。*
