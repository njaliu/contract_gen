"""
从需求表格（xlsx）规划合同生成任务，并调用 gen_single_contract / mutate_single_contract 执行。

输出目录规则：在 ``--output-dir`` 下自动创建以**当天日期**（``YYYY-MM-DD``）命名的总文件夹；同一
``所属编号`` 的多行链条合同写入 ``日期/所属编号/`` 子目录；仅一行的独立合同直接写在日期文件夹根下。

运行示例（在项目根目录）::

    python3 src/contract_task_plan.py \\
      --xlsx /path/to/合同信息收集表.xlsx \\
      --output-dir resources/output \\
      --template-file /path/to/template.md
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import random
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Literal

_SRC = Path(__file__).resolve().parent
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from gen_contract import gen_single_contract, mutate_single_contract

logger = logging.getLogger(__name__)

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]+')
MULTI_UNDERSCORE = re.compile(r"_+")

HEADER_ALIASES: dict[str, str] = {
    "序号": "序号",
    "甲方名称": "甲方名称",
    "乙方名称": "乙方名称",
    "项目名称": "项目名称",
    "合同金额": "合同金额",
    "签订日期": "签订日期",
    "验收日期": "验收日期",
    "所属编号": "所属编号",
    "是否加急": "是否加急",
}


def _load_openpyxl():
    try:
        import openpyxl  # type: ignore
    except ImportError as e:
        raise RuntimeError(
            "读取 xlsx 需要 openpyxl。请执行: python3 -m pip install -r requirements.txt"
        ) from e
    return openpyxl


def _cell_str(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _format_amount_for_filename(v: Any) -> str:
    if v is None or (isinstance(v, str) and not v.strip()):
        return "金额待定"
    if isinstance(v, (int, float)):
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v)
    return _sanitize_segment(str(v), max_len=32)


def _sanitize_segment(s: str, *, max_len: int = 80) -> str:
    t = INVALID_FILENAME_CHARS.sub("_", s.strip())
    t = MULTI_UNDERSCORE.sub("_", t).strip("_")
    if not t:
        t = "未命名"
    if len(t) > max_len:
        t = t[:max_len].rstrip("_")
    return t or "未命名"


def _row_is_skippable_empty(row: dict[str, Any]) -> bool:
    def empty(v: Any) -> bool:
        return v is None or (isinstance(v, str) and not str(v).strip())

    if empty(row.get("甲方名称")) and empty(row.get("乙方名称")) and empty(row.get("项目名称")):
        return True
    return False


def _group_key_for_row(row: dict[str, Any]) -> str:
    gid = row.get("所属编号")
    if gid is None:
        return f"__solo_r{row['_excel_row']}"
    if isinstance(gid, str) and not gid.strip():
        return f"__solo_r{row['_excel_row']}"
    return str(gid).strip()


def _sort_key_seq(seq: Any) -> tuple[int, Any]:
    if seq is None:
        return (10**9, "")
    try:
        return (0, int(seq))
    except (TypeError, ValueError):
        return (1, str(seq))


def read_requirement_rows(xlsx_path: str) -> list[dict[str, Any]]:
    openpyxl = _load_openpyxl()
    path = Path(xlsx_path)
    if not path.exists():
        raise FileNotFoundError(f"xlsx 不存在: {xlsx_path}")

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header_row = next(rows_iter, None)
    if not header_row:
        wb.close()
        return []

    headers: list[str] = []
    for h in header_row:
        if h is None:
            headers.append("")
        else:
            headers.append(str(h).strip())

    out: list[dict[str, Any]] = []
    excel_row = 1
    for data in rows_iter:
        excel_row += 1
        row: dict[str, Any] = {"_excel_row": excel_row}
        for col_name, val in zip(headers, data):
            if not col_name:
                continue
            canonical = HEADER_ALIASES.get(col_name, col_name)
            row[canonical] = val
        if _row_is_skippable_empty(row):
            logger.debug("跳过空行: excel_row=%s", excel_row)
            continue
        out.append(row)
    wb.close()
    return out


def build_basic_info_markdown(row: dict[str, Any]) -> str:
    lines = [
        "# 合同基本信息（来自需求表格）",
        "",
        "生成或改写合同时，下列字段须与本表一致；研发内容、研发目标、产品目标可在不改动技术指标与量化要求的前提下调整表述。",
        "",
        "| 字段 | 值 |",
        "| --- | --- |",
    ]
    fields = [
        "序号",
        "所属编号",
        "甲方名称",
        "乙方名称",
        "项目名称",
        "合同金额",
        "签订日期",
        "验收日期",
        "是否加急",
    ]
    for f in fields:
        v = row.get(f)
        if v is None:
            cell = ""
        elif isinstance(v, float) and v == int(v):
            cell = str(int(v))
        else:
            cell = str(v)
        cell = cell.replace("|", "\\|")
        lines.append(f"| {f} | {cell} |")
    lines.append("")
    return "\n".join(lines)


def _chain_group_subdir_name(g_rows: list[dict[str, Any]]) -> str:
    """链条组子目录名：取组内第一行的所属编号（多行组内应一致）。"""
    v = g_rows[0].get("所属编号")
    if isinstance(v, float) and abs(v - int(v)) < 1e-9:
        v = int(v)
    return _sanitize_segment(str(v), max_len=64)


def contract_output_filename(row: dict[str, Any], *, disambig_suffix: str = "") -> str:
    jia = _cell_str(row.get("甲方名称")) or "未知甲方"
    yi = _cell_str(row.get("乙方名称")) or "未知乙方"
    proj = _cell_str(row.get("项目名称")) or "未知项目"
    amt = _format_amount_for_filename(row.get("合同金额"))
    stem = "-".join(
        [
            _sanitize_segment(jia, max_len=40),
            _sanitize_segment(yi, max_len=40),
            _sanitize_segment(amt, max_len=32),
            _sanitize_segment(proj, max_len=60),
        ]
    )
    if disambig_suffix:
        stem = f"{stem}_{disambig_suffix}"
    return stem + ".md"


@dataclass(frozen=True)
class PlannedTask:
    kind: Literal["gen", "mutate"]
    group_key: str
    excel_row: int
    row: dict[str, Any]
    output_path: Path
    basic_info_path: Path
    base_contract_path: Path | None = None


def _discover_format_ids(format_spec_dir: str) -> list[str]:
    spec_dir = Path(format_spec_dir).expanduser()
    if not spec_dir.exists():
        raise FileNotFoundError(f"FormatSpec 目录不存在: {spec_dir}")
    ids: list[str] = []
    for p in sorted(spec_dir.glob("docx_format_*.json")):
        ids.append(p.stem)
    if not ids:
        raise RuntimeError(f"FormatSpec 目录下未找到 docx_format_*.json: {spec_dir}")
    return ids


def _assign_task_format_ids(
    tasks: list[PlannedTask],
    *,
    output_format: str,
    format_scope: Literal["batch", "group", "file"],
    format_spec_dir: str,
    format_id: str | None,
    format_seed: int | None,
) -> dict[int, str | None]:
    fmt = output_format.lower()
    if fmt == "md":
        return {id(t): None for t in tasks}

    available_ids = _discover_format_ids(format_spec_dir)
    if format_id:
        if format_id not in available_ids:
            raise ValueError(f"指定的 format_id 不存在: {format_id}")
        return {id(t): format_id for t in tasks}

    rng = random.Random(format_seed)
    assigned: dict[int, str | None] = {}
    if format_scope == "batch":
        selected = rng.choice(available_ids)
        for t in tasks:
            assigned[id(t)] = selected
        return assigned

    if format_scope == "group":
        groups = sorted({t.group_key for t in tasks})
        group_to_format: dict[str, str] = {}
        for g in groups:
            group_to_format[g] = rng.choice(available_ids)
        for t in tasks:
            assigned[id(t)] = group_to_format[t.group_key]
        return assigned

    # file
    for t in tasks:
        assigned[id(t)] = rng.choice(available_ids)
    return assigned


def task_plan(
    xlsx_path: str,
    output_dir: str,
    *,
    contract_skill_path: str,
    template_path: str,
    mutation_skill_path: str,
) -> list[PlannedTask]:
    rows = read_requirement_rows(xlsx_path)
    if not rows:
        return []

    groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        groups[_group_key_for_row(row)].append(row)

    for gk in groups:
        groups[gk].sort(key=lambda r: _sort_key_seq(r.get("序号")))

    group_order: list[tuple[int, str, list[dict[str, Any]]]] = []
    for gk, g_rows in groups.items():
        min_excel = min(r["_excel_row"] for r in g_rows)
        group_order.append((min_excel, gk, g_rows))
    group_order.sort(key=lambda x: x[0])

    day_name = date.today().strftime("%Y-%m-%d")
    root_run_dir = Path(output_dir).expanduser() / day_name
    basic_dir = root_run_dir / "_task_basic_info"

    used_names_by_parent: dict[str, set[str]] = {}
    tasks: list[PlannedTask] = []

    for _min_excel, group_key, g_rows in group_order:
        if len(g_rows) > 1:
            contract_parent = root_run_dir / _chain_group_subdir_name(g_rows)
        else:
            contract_parent = root_run_dir

        parent_key = str(contract_parent.resolve())
        if parent_key not in used_names_by_parent:
            used_names_by_parent[parent_key] = set()

        base_output: Path | None = None
        for idx, row in enumerate(g_rows):
            excel_row = int(row["_excel_row"])
            seq = row.get("序号")
            disambig = f"g{group_key}_序{seq}"
            fname = contract_output_filename(row, disambig_suffix="")
            if fname in used_names_by_parent[parent_key]:
                fname = contract_output_filename(row, disambig_suffix=disambig)
            used_names_by_parent[parent_key].add(fname)

            out_path = (contract_parent / fname).resolve()
            bip = (basic_dir / f"basic_info_r{excel_row}.md").resolve()

            if len(g_rows) == 1:
                tasks.append(
                    PlannedTask(
                        kind="gen",
                        group_key=group_key,
                        excel_row=excel_row,
                        row=row,
                        output_path=out_path,
                        basic_info_path=bip,
                        base_contract_path=None,
                    )
                )
            else:
                if idx == 0:
                    tasks.append(
                        PlannedTask(
                            kind="gen",
                            group_key=group_key,
                            excel_row=excel_row,
                            row=row,
                            output_path=out_path,
                            basic_info_path=bip,
                            base_contract_path=None,
                        )
                    )
                    base_output = out_path
                else:
                    if base_output is None:
                        raise RuntimeError(f"组 {group_key} 缺少 base 输出路径")
                    tasks.append(
                        PlannedTask(
                            kind="mutate",
                            group_key=group_key,
                            excel_row=excel_row,
                            row=row,
                            output_path=out_path,
                            basic_info_path=bip,
                            base_contract_path=base_output,
                        )
                    )
    return tasks


def run_task_plan(
    tasks: list[PlannedTask],
    *,
    contract_skill_path: str,
    template_path: str,
    mutation_skill_path: str,
    api_key: str | None = None,
    base_url: str = "https://openrouter.ai/api/v1",
    model: str = "deepseek/deepseek-v3.2",
    gen_temperature: float = 0.2,
    mutate_temperature: float = 0.4,
    timeout: int = 300,
    output_format: str = "md",
    format_scope: Literal["batch", "group", "file"] = "group",
    format_spec_dir: str = "resources/formats",
    format_id: str | None = None,
    format_seed: int | None = None,
    manifest_path: str | None = None,
) -> list[tuple[PlannedTask, str]]:
    assigned_format_ids = _assign_task_format_ids(
        tasks,
        output_format=output_format,
        format_scope=format_scope,
        format_spec_dir=format_spec_dir,
        format_id=format_id,
        format_seed=format_seed,
    )
    results: list[tuple[PlannedTask, str]] = []
    manifest_records: list[dict[str, Any]] = []
    for t in tasks:
        selected_format_id = assigned_format_ids.get(id(t))
        t.basic_info_path.parent.mkdir(parents=True, exist_ok=True)
        t.output_path.parent.mkdir(parents=True, exist_ok=True)
        t.basic_info_path.write_text(build_basic_info_markdown(t.row), encoding="utf-8")
        logger.info(
            "执行任务 kind=%s group=%s excel_row=%s format_id=%s -> %s",
            t.kind,
            t.group_key,
            t.excel_row,
            selected_format_id,
            t.output_path,
        )
        if t.kind == "gen":
            text = gen_single_contract(
                skill_file=contract_skill_path,
                template_file=template_path,
                basic_info_file=str(t.basic_info_path),
                output_file=str(t.output_path),
                api_key=api_key,
                base_url=base_url,
                model=model,
                temperature=gen_temperature,
                timeout=timeout,
                output_format=output_format,
                format_spec_dir=format_spec_dir,
                format_id=selected_format_id,
                format_seed=None,
            )
        else:
            if not t.base_contract_path:
                raise RuntimeError("mutate 任务缺少 base_contract_path")
            text = mutate_single_contract(
                base_contract_file=str(t.base_contract_path),
                skill_file=mutation_skill_path,
                basic_info_file=str(t.basic_info_path),
                output_file=str(t.output_path),
                api_key=api_key,
                base_url=base_url,
                model=model,
                temperature=mutate_temperature,
                timeout=timeout,
                output_format=output_format,
                format_spec_dir=format_spec_dir,
                format_id=selected_format_id,
                format_seed=None,
            )
        results.append((t, text))
        manifest_records.append(
            {
                "group_key": t.group_key,
                "excel_row": t.excel_row,
                "task_kind": t.kind,
                "output_base": str(t.output_path),
                "output_format": output_format,
                "selected_format_id": selected_format_id,
                "format_scope": format_scope,
                "row_summary": {
                    "序号": _cell_str(t.row.get("序号")),
                    "甲方名称": _cell_str(t.row.get("甲方名称")),
                    "乙方名称": _cell_str(t.row.get("乙方名称")),
                    "项目名称": _cell_str(t.row.get("项目名称")),
                    "合同金额": _cell_str(t.row.get("合同金额")),
                    "所属编号": _cell_str(t.row.get("所属编号")),
                },
            }
        )

    if manifest_path:
        mp = Path(manifest_path).expanduser()
        mp.parent.mkdir(parents=True, exist_ok=True)
        mp.write_text(
            json.dumps(
                {
                    "output_format": output_format,
                    "format_scope": format_scope,
                    "format_spec_dir": str(Path(format_spec_dir).expanduser()),
                    "format_id_override": format_id,
                    "format_seed": format_seed,
                    "records": manifest_records,
                },
                ensure_ascii=False,
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )
        logger.info("已写入 manifest: %s", mp.resolve())
    return results


def plan_and_run_contract_tasks(
    xlsx_path: str,
    output_dir: str,
    *,
    contract_skill_path: str,
    template_path: str,
    mutation_skill_path: str,
    api_key: str | None = None,
    base_url: str = "https://openrouter.ai/api/v1",
    model: str = "deepseek/deepseek-v3.2",
    gen_temperature: float = 0.2,
    mutate_temperature: float = 0.4,
    timeout: int = 300,
    output_format: str = "md",
    format_scope: Literal["batch", "group", "file"] = "group",
    format_spec_dir: str = "resources/formats",
    format_id: str | None = None,
    format_seed: int | None = None,
) -> list[tuple[PlannedTask, str]]:
    tasks = task_plan(
        xlsx_path,
        output_dir,
        contract_skill_path=contract_skill_path,
        template_path=template_path,
        mutation_skill_path=mutation_skill_path,
    )
    return run_task_plan(
        tasks,
        contract_skill_path=contract_skill_path,
        template_path=template_path,
        mutation_skill_path=mutation_skill_path,
        api_key=api_key,
        base_url=base_url,
        model=model,
        gen_temperature=gen_temperature,
        mutate_temperature=mutate_temperature,
        timeout=timeout,
        output_format=output_format,
        format_scope=format_scope,
        format_spec_dir=format_spec_dir,
        format_id=format_id,
        format_seed=format_seed,
        manifest_path=str(
            (Path(output_dir).expanduser() / date.today().strftime("%Y-%m-%d") / "manifest.json")
        ),
    )


def _parse_cli() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="从需求 xlsx 批量规划并生成合同")
    p.add_argument("--xlsx", required=True, help="合同信息收集表 xlsx 路径")
    p.add_argument(
        "--output-dir",
        default="resources/output",
        help="输出根目录；其下会创建当天日期文件夹 YYYY-MM-DD（默认 resources/output）",
    )
    p.add_argument(
        "--contract-skill",
        default="resources/input/contract_skill.md",
        help="首份合同生成用 skill 文件",
    )
    p.add_argument(
        "--mutation-skill",
        default="resources/input/contract_mutation_skill.md",
        help="关联合同 mutation 用 skill 文件",
    )
    p.add_argument(
        "--template-file",
        required=True,
        help="合同模板文件路径（.md/.docx/.pdf）",
    )
    p.add_argument("--model", default=None, help="模型名（默认 deepseek/deepseek-v3.2 或环境变量 LLM_MODEL）")
    p.add_argument("--base-url", default=None, help="OpenAI 兼容 base URL（默认 LLM_BASE_URL 或 OpenRouter）")
    p.add_argument("--gen-temperature", type=float, default=0.2)
    p.add_argument("--mutate-temperature", type=float, default=0.4)
    p.add_argument("--timeout", type=int, default=300)
    p.add_argument(
        "--output-format",
        default="md",
        choices=["md", "docx", "both"],
        help="输出文件格式（默认: md）",
    )
    p.add_argument(
        "--format-scope",
        default="group",
        choices=["batch", "group", "file"],
        help="格式分配范围：batch/group/file（默认 group）",
    )
    p.add_argument(
        "--format-spec-dir",
        default="resources/formats",
        help="DOCX FormatSpec 目录（默认: resources/formats）",
    )
    p.add_argument(
        "--format-id",
        default=None,
        help="指定固定 FormatSpec ID（如 docx_format_03）",
    )
    p.add_argument(
        "--format-seed",
        type=int,
        default=None,
        help="随机选择 FormatSpec 的种子（未指定 format-id 时生效）",
    )
    p.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="仅规划任务并打印，不调用大模型",
    )
    return p.parse_args()


def main() -> None:
    args = _parse_cli()
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        stream=sys.stderr,
    )
    model = args.model or os.getenv("LLM_MODEL", "deepseek/deepseek-v3.2")
    base_url = args.base_url or os.getenv("LLM_BASE_URL", "https://openrouter.ai/api/v1")

    tasks = task_plan(
        args.xlsx,
        args.output_dir,
        contract_skill_path=args.contract_skill,
        template_path=args.template_file,
        mutation_skill_path=args.mutation_skill,
    )
    logger.info("共规划 %d 个任务", len(tasks))
    for t in tasks:
        logger.info(
            "  [%s] group=%s row=%s -> %s",
            t.kind,
            t.group_key,
            t.excel_row,
            t.output_path,
        )

    day_folder = Path(args.output_dir).expanduser() / date.today().strftime("%Y-%m-%d")
    assigned_format_ids = _assign_task_format_ids(
        tasks,
        output_format=args.output_format,
        format_scope=args.format_scope,
        format_spec_dir=args.format_spec_dir,
        format_id=args.format_id,
        format_seed=args.format_seed,
    )
    if args.output_format != "md":
        logger.info("格式分配预览（task -> format_id）:")
        for t in tasks:
            logger.info(
                "  group=%s row=%s kind=%s format_id=%s",
                t.group_key,
                t.excel_row,
                t.kind,
                assigned_format_ids.get(id(t)),
            )
    if args.dry_run:
        print(
            f"dry-run 完成，已规划 {len(tasks)} 个任务；当日输出将位于: {day_folder.resolve()}",
            file=sys.stderr,
        )
        return

    plan_and_run_contract_tasks(
        args.xlsx,
        args.output_dir,
        contract_skill_path=args.contract_skill,
        template_path=args.template_file,
        mutation_skill_path=args.mutation_skill,
        model=model,
        base_url=base_url,
        gen_temperature=args.gen_temperature,
        mutate_temperature=args.mutate_temperature,
        timeout=args.timeout,
        output_format=args.output_format,
        format_scope=args.format_scope,
        format_spec_dir=args.format_spec_dir,
        format_id=args.format_id,
        format_seed=args.format_seed,
    )
    print(f"已完成 {len(tasks)} 个任务，输出目录: {day_folder.resolve()}")


if __name__ == "__main__":
    main()
