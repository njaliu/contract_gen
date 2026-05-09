# Contract Gen

基于大模型的合同生成工具，支持：

- 单份合同生成（`gen_single_contract`）
- 单份合同变换（`mutate_single_contract`）
- 从需求表（Excel）批量规划并执行生成任务（首份 `gen` + 关联合同 `mutate`）
- 输出 `md` / `docx` / `both`

---

## 1. 项目结构

```text
contract_gen/
├── requirements.txt
├── README.md
├── src/
│   ├── gen_contract.py         # 单份生成/变换主逻辑 + 单份CLI
│   ├── contract_task_plan.py   # 批量任务规划与执行CLI（xlsx）
│   └── output_writers.py       # md/docx输出与FormatSpec应用
├── resources/
│   ├── input/
│   │   ├── contract_skill.md
│   │   ├── contract_mutation_skill.md
│   │   ├── docx_format_skill.md    # 预留的DOCX格式生成规范，当前版本暂未接入执行链路
│   │   ├── contract_template.pdf
│   │   ├── info.md                 # 基本信息表，供单文件生成使用
│   │   └── request.xlsx            # 客户需求表格（来源于客户）
│   ├── formats/
│   │   └── docx_format_01~08.json  # 预置8套DOCX格式
│   └── output/
└── tests/
```

---

## 2. 环境准备

### 2.1 Python 依赖

```bash
python3 -m pip install -r requirements.txt
```

### 2.2 API Key

至少设置以下之一：

- `OPENROUTER_API_KEY`（优先）
- `LLM_API_KEY`
- `OPENAI_API_KEY`

示例：

```bash
export OPENROUTER_API_KEY="your_api_key"
```

可选：

```bash
export LLM_MODEL="deepseek/deepseek-v3.2"
export LLM_BASE_URL="https://openrouter.ai/api/v1"
```

---

## 3. 单份合同生成

脚本：`src/gen_contract.py`

```bash
python3 src/gen_contract.py \
  --skill-file resources/input/contract_skill.md \
  --template-file resources/input/contract_template.pdf \
  --basic-info-file resources/input/info.md \
  --output-file resources/output/compare_contract.md \
  --output-format both \
  --format-spec-dir resources/formats \
  --format-seed 42
```

### 常用参数

- `--output-format md|docx|both`（默认 `md`）
- `--format-spec-dir`：DOCX 格式配置目录
- `--format-id`：固定使用某套格式（如 `docx_format_02`）
- `--format-seed`：随机选格式时可复现
- `--model`、`--base-url`、`--temperature`、`--timeout`

---

## 4. 批量合同生成（需求表）

脚本：`src/contract_task_plan.py`

```bash
python3 src/contract_task_plan.py \
  --xlsx resources/input/request.xlsx \
  --template-file resources/input/contract_template.pdf \
  --output-dir resources/output \
  --output-format both \
  --format-scope group \
  --format-seed 42
```

### 批量模式逻辑

- 按 `所属编号` 分组
- 组内第一份：调用 `gen_single_contract`
- 组内其余份：调用 `mutate_single_contract`
- 当天输出目录：`{output-dir}/YYYY-MM-DD/`
- 链条组目录：`{output-dir}/YYYY-MM-DD/{所属编号}/`
- 独立合同：直接放在当天目录下

### 格式分配策略（`--format-scope`）

- `batch`：整批同一格式
- `group`：同组同一格式（推荐）
- `file`：每份独立随机格式

### Dry Run

仅看任务规划与格式分配，不调用大模型：

```bash
python3 src/contract_task_plan.py \
  --xlsx resources/input/request.xlsx \
  --template-file resources/input/contract_template.pdf \
  --dry-run \
  --output-format both \
  --format-scope group
```

---

## 5. DOCX 输出说明

### 5.1 FormatSpec

`resources/formats/docx_format_*.json` 定义页面、段落、表格参数。  
输出 docx 时会按以下优先级选取：

1. 若指定 `--format-id`，使用该格式
2. 否则从目录中随机选（可用 `--format-seed` 固定）

### 5.2 Markdown -> DOCX 层级映射（已做兜底）

- 支持 Markdown 标题映射：`# / ### / #### / #####`
- 并支持合同编号规则兜底：
  - `1~12`章标题 -> `h1`
  - `6.1/6.2` -> `h2`
  - `附件：...` -> `h2_emphasis`
  - `验收结论` -> `h3`
  - `1.1 / 2.1 / 6.1.1 ...` -> `body`

---

## 6. 关键代码入口

- `src/gen_contract.py`
  - `gen_single_contract(...)`
  - `mutate_single_contract(...)`
  - `build_messages(...)`
  - `build_mutation_messages(...)`
- `src/contract_task_plan.py`
  - `task_plan(...)`
  - `run_task_plan(...)`
  - `plan_and_run_contract_tasks(...)`
- `src/output_writers.py`
  - `write_contract_output(...)`

---

## 7. 输出与追踪

批量执行时，会在当天目录自动生成：

- 合同文件（md/docx）
- `_task_basic_info/` 中间信息文件
- `manifest.json`（任务、格式ID、输出路径追踪）

---

## 8. 常见问题

### Q1: 输出 docx 报依赖错误

确认已安装：

```bash
python3 -m pip install -r requirements.txt
```

### Q2: 模型不按预期输出

- 先用 `--dry-run` 验证任务规划
- 降低温度（`--temperature` / `--gen-temperature` / `--mutate-temperature`）
- 固定格式（`--format-id`）以减少干扰变量

### Q3: 为什么有些编号看起来像正文

这是合同模板与目标样式的一部分：并非所有编号都是标题，映射规则已在 writer 中按文档实际格式处理。