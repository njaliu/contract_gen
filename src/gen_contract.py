import argparse
import json
import logging
import os
import re
import time
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib import error, request

from output_writers import write_contract_output

logger = logging.getLogger(__name__)


def read_text_file(file_path: str) -> str:
    path = Path(file_path)
    logger.info("读取文本文件: %s", path.resolve())
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    text = path.read_text(encoding="utf-8")
    logger.info("文本读取完成，字符数: %d", len(text))
    return text


def read_docx_text(file_path: str) -> str:
    path = Path(file_path)
    logger.info("解析 DOCX 模板: %s", path.resolve())
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    with zipfile.ZipFile(path, "r") as zf:
        xml_data = zf.read("word/document.xml")

    root = ET.fromstring(xml_data)
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs: list[str] = []
    for p in root.findall(".//w:p", namespace):
        text_nodes = p.findall(".//w:t", namespace)
        paragraph_text = "".join(node.text or "" for node in text_nodes)
        paragraphs.append(paragraph_text)
    out = "\n".join(paragraphs)
    logger.info("DOCX 解析完成，段落数: %d，合并字符数: %d", len(paragraphs), len(out))
    return out


def read_pdf_text(file_path: str) -> str:
    path = Path(file_path)
    logger.info("解析 PDF 模板: %s", path.resolve())
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    # 尽量复用常见 PDF 库，避免把二进制 PDF 当成 UTF-8 文本读取。
    readers = []
    try:
        from pypdf import PdfReader as PypdfReader  # type: ignore

        readers.append(("pypdf", PypdfReader))
    except Exception:
        pass

    try:
        from PyPDF2 import PdfReader as PyPdf2Reader  # type: ignore

        readers.append(("PyPDF2", PyPdf2Reader))
    except Exception:
        pass

    if not readers:
        raise RuntimeError(
            "读取 PDF 失败：未安装 PDF 解析依赖。"
            "请先执行 `python3 -m pip install pypdf`，或改用 .docx/.md 模板。"
        )

    logger.debug("可用 PDF 解析库: %s", [name for name, _ in readers])
    last_error: Exception | None = None
    for name, reader_cls in readers:
        try:
            logger.info("尝试使用 %s 读取 PDF …", name)
            reader = reader_cls(str(path))
            pages = []
            for page in reader.pages:
                pages.append((page.extract_text() or "").strip())
            text = "\n\n".join(part for part in pages if part)
            if text.strip():
                logger.info(
                    "PDF 读取成功（%s），非空页片段数: %d，合并字符数: %d",
                    name,
                    sum(1 for p in pages if p),
                    len(text),
                )
                return text
        except Exception as e:
            last_error = e
            logger.warning("使用 %s 读取 PDF 失败: %s", name, e)
            continue

    raise RuntimeError(f"读取 PDF 失败，请检查文件是否损坏或加密。详细信息: {last_error}")


def read_template(file_path: str) -> str:
    suffix = Path(file_path).suffix.lower()
    logger.info("加载合同模板，类型后缀: %s", suffix or "(无)")
    if suffix == ".docx":
        return read_docx_text(file_path)
    if suffix == ".pdf":
        return read_pdf_text(file_path)
    return read_text_file(file_path)


def build_messages(
    skill_text: str,
    template_text: str,
    basic_info_text: str,
) -> list[dict[str, str]]:
    system_prompt = (
        "你是一名专业法务合同写作助手。"
        "你必须严格按照模板生成合同。"
        "禁止输出解释、分析过程、免责声明或额外说明。"
    )
    user_prompt = (
        "请基于以下三份输入生成完整合同，且满足“固定内容不变、仅替换[]变量内容”规则：\n\n"
        "【输入1：Prompt要求（contract_skill.md）】\n"
        f"{skill_text}\n\n"
        "【输入2：合同模板】\n"
        f"{template_text}\n\n"
        "【输入3：合同基本信息】\n"
        f"{basic_info_text}\n\n"
        "强制要求：\n"
        "1) 模板中方括号 [] 内是可变内容，必须优先依据合同基本信息生成并替换。\n"
        "2) 固定内容按“语义与顺序”保持不变，不得新增、删除或改变条款含义。\n"
        "3) Markdown 层级必须遵循以下结构：文档主标题用 #；章标题（如“1 项目名称与研发范围”到“12 合同份数”）用 ###。\n"
        "4) 仅“6.1 甲方权利与义务”和“6.2 乙方权利与义务”使用 ####；“验收结论”使用 #####；“附件：项目验收单”为附件强调标题。\n"
        "5) 编号条款正文（如 1.1、2.1、6.1.1 等）必须保持为普通正文段落，不得使用 #### 或 ##### 标题层级。\n"
        "6) 若某变量信息不足，用“待补充”填入对应 []。\n"
        "7) 只输出最终合同正文，不输出任何解释。"
    )
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]
    total_chars = sum(len(m["content"]) for m in messages)
    logger.info(
        "已构建对话消息: %d 条，user 提示估算字符数: %d，合计约 %d 字符",
        len(messages),
        len(user_prompt),
        total_chars,
    )
    return messages


def build_mutation_messages(
    mutation_skill_text: str,
    base_contract_text: str,
    basic_info_text: str,
) -> list[dict[str, str]]:
    system_prompt = (
        "你是一名专业法务合同文本变换助手。"
        "你的任务是在基准合同基础上做受约束改写。"
        "禁止输出解释、分析过程、免责声明或额外说明。"
    )
    user_prompt = (
        "请基于以下三份输入，对基准合同做一次 mutation 变换，并输出完整合同正文：\n\n"
        "【输入1：Mutation要求（contract_mutation_skill.md）】\n"
        f"{mutation_skill_text}\n\n"
        "【输入2：基准合同（base contract）】\n"
        f"{base_contract_text}\n\n"
        "【输入3：合同基本信息】\n"
        f"{basic_info_text}\n\n"
        "强制要求：\n"
        "1) 必须以基准合同为基础改写，不得重写全新合同。\n"
        "2) 技术指标、量化技术要求、评价标准必须严格保持不变。\n"
        "3) 仅允许调整“研发内容、研发目标、产品目标”的文字描述。\n"
        "4) 可以适当扩展或调整表述方式，但不得新增冲突要求。\n"
        "5) 只输出最终合同正文，不输出任何解释。"
    )
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]
    total_chars = sum(len(m["content"]) for m in messages)
    logger.info(
        "已构建 mutation 对话消息: %d 条，user 提示估算字符数: %d，合计约 %d 字符",
        len(messages),
        len(user_prompt),
        total_chars,
    )
    return messages


def call_llm(
    api_key: str,
    base_url: str,
    model: str,
    messages: list[dict[str, str]],
    temperature: float = 0.1,
    timeout: int = 300,
) -> str:
    url = f"{base_url.rstrip('/')}/chat/completions"
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
    }
    data = json.dumps(payload).encode("utf-8")
    req = request.Request(
        url=url,
        data=data,
        method="POST",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": os.getenv("OPENROUTER_SITE_URL", "https://localhost"),
            "X-Title": os.getenv("OPENROUTER_APP_NAME", "contract-generator"),
        },
    )
    logger.info(
        "调用大模型 POST %s | model=%s | timeout=%ds | temperature=%s | payload_bytes=%d",
        url,
        model,
        timeout,
        temperature,
        len(data),
    )
    # 对网络超时类故障进行最多 3 次指数退避重试（1s, 2s, 4s）。
    max_retries = 3
    retryable_err: Exception | None = None
    for attempt in range(max_retries):
        try:
            logger.info("HTTP 请求第 %d/%d 次 …", attempt + 1, max_retries)
            with request.urlopen(req, timeout=timeout) as resp:
                status = getattr(resp, "status", None) or resp.getcode()
                logger.info("HTTP 状态: %s", status)
                body = resp.read().decode("utf-8")
                logger.info("收到响应体，长度: %d 字符", len(body))
                logger.debug("响应体原文:\n%s", body)
                result = json.loads(body)
            logger.info("JSON 解析成功，choices 数量: %d", len(result.get("choices", [])))
            break
        except error.HTTPError as e:
            detail = e.read().decode("utf-8", errors="ignore")
            logger.error("HTTP 错误 %s，详情: %s", e.code, detail[:2000])
            raise RuntimeError(f"调用大模型失败，HTTP {e.code}: {detail}") from e
        except TimeoutError as e:
            retryable_err = e
            logger.warning("请求超时(将重试): %s", e)
        except error.URLError as e:
            retryable_err = e
            logger.warning("网络 URL 错误(将重试): %s", e)

        if attempt < max_retries - 1:
            delay = 2**attempt
            logger.info("等待 %d 秒后重试 …", delay)
            time.sleep(delay)
    else:
        logger.error("已达最大重试次数，放弃请求")
        raise RuntimeError(f"网络错误，无法连接到大模型接口: {retryable_err}") from retryable_err

    try:
        content = result["choices"][0]["message"]["content"].strip()
        logger.info("提取模型回复成功，合同正文长度: %d 字符", len(content))
        return content
    except (KeyError, IndexError, TypeError) as e:
        logger.error("解析 choices 失败，result 键: %s", list(result.keys()) if isinstance(result, dict) else type(result))
        raise RuntimeError(f"大模型返回格式异常: {result}") from e


def validate_fixed_content(template_text: str, generated_text: str) -> tuple[bool, str]:
    logger.info(
        "校验固定文本：模板字符数=%d，生成文本字符数=%d",
        len(template_text),
        len(generated_text),
    )
    fixed_segments = re.split(r"\[[^\]]*\]", template_text)
    current = 0
    for idx, segment in enumerate(fixed_segments):
        if not segment:
            continue
        pos = generated_text.find(segment, current)
        if pos == -1:
            preview = segment[:80].replace("\n", "\\n")
            logger.warning("固定片段校验失败，索引=%d，片段预览: %r …", idx, preview)
            return False, f"固定片段未按顺序保留，片段索引: {idx}"
        current = pos + len(segment)
    logger.info("固定文本校验通过")
    return True, ""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="基于合同要求、模板和基本信息文件，调用大模型生成完整合同。"
    )
    parser.add_argument(
        "--skill-file",
        required=True,
        help="输入1：contract_skill.md 文件路径",
    )
    parser.add_argument(
        "--template-file",
        required=True,
        help="输入2：合同模板文件路径",
    )
    parser.add_argument(
        "--basic-info-file",
        required=True,
        help="输入3：合同基本信息文件路径",
    )
    parser.add_argument(
        "--output-file",
        default="generated_contract.md",
        help="输出合同文件路径（默认: generated_contract.md）",
    )
    parser.add_argument(
        "--model",
        default=os.getenv("LLM_MODEL", "deepseek/deepseek-v3.2"),
        help="模型名称（默认读取 LLM_MODEL 或 deepseek/deepseek-v3.2）",
    )
    parser.add_argument(
        "--base-url",
        default=os.getenv("LLM_BASE_URL", "https://openrouter.ai/api/v1"),
        help="OpenAI 兼容接口地址（默认读取 LLM_BASE_URL 或 https://openrouter.ai/api/v1）",
    )
    parser.add_argument(
        "--temperature",
        type=float,
        default=0.2,
        help="采样温度（默认: 0.2）",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=300,
        help="请求超时时间（秒，默认: 300）",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="日志级别（默认: INFO；DEBUG 会打印完整 HTTP 响应体）",
    )
    parser.add_argument(
        "--output-format",
        default="md",
        choices=["md", "docx", "both"],
        help="输出文件格式（默认: md；可选 docx 或 both）",
    )
    parser.add_argument(
        "--format-spec-dir",
        default="resources/formats",
        help="DOCX FormatSpec 目录（默认: resources/formats）",
    )
    parser.add_argument(
        "--format-id",
        default=None,
        help="指定 DOCX FormatSpec ID（如 docx_format_03，不指定则随机）",
    )
    parser.add_argument(
        "--format-seed",
        type=int,
        default=None,
        help="DOCX FormatSpec 随机种子（未指定 format-id 时生效）",
    )
    return parser.parse_args()


def gen_single_contract(
    skill_file: str,
    template_file: str,
    basic_info_file: str,
    output_file: str | None = None,
    *,
    api_key: str | None = None,
    base_url: str = "https://openrouter.ai/api/v1",
    model: str = "deepseek/deepseek-v3.2",
    temperature: float = 0.2,
    timeout: int = 300,
    output_format: str = "md",
    format_spec_dir: str = "resources/formats",
    format_id: str | None = None,
    format_seed: int | None = None,
) -> str:
    resolved_api_key = (
        api_key
        or os.getenv("OPENROUTER_API_KEY")
        or os.getenv("LLM_API_KEY")
        or os.getenv("OPENAI_API_KEY")
    )
    if not resolved_api_key:
        raise RuntimeError(
            "未设置 API Key。请设置 OPENROUTER_API_KEY（或 LLM_API_KEY / OPENAI_API_KEY）"
        )

    logger.info("步骤 1/5：读取合同要求文件 …")
    skill_text = read_text_file(skill_file)
    logger.info("步骤 2/5：读取合同模板 …")
    template_text = read_template(template_file)
    logger.info("步骤 3/5：读取合同基本信息 …")
    basic_info_text = read_text_file(basic_info_file)
    logger.info("步骤 4/5：构建消息并调用大模型 …")
    messages = build_messages(skill_text, template_text, basic_info_text)

    contract_text = call_llm(
        api_key=resolved_api_key,
        base_url=base_url,
        model=model,
        messages=messages,
        temperature=temperature,
        timeout=timeout,
    )
    logger.info("模型生成完成，合同正文长度: %d 字符", len(contract_text))
    logger.debug("模型生成的合同正文:\n%s", contract_text)

    # 固定片段校验逻辑临时关闭：
    # is_valid, reason = validate_fixed_content(template_text, contract_text)
    # if not is_valid:
    #     logger.warning("首次生成未通过固定文本校验，原因: %s，将发起第二次生成 …", reason)
    #     retry_messages = messages + [
    #         {
    #             "role": "user",
    #             "content": (
    #                 "你上一版输出未满足“固定文本完全不变”规则。"
    #                 f"问题: {reason}。请严格仅替换 [] 内变量，重新输出完整合同。"
    #             ),
    #         }
    #     ]
    #     contract_text = call_llm(
    #         api_key=resolved_api_key,
    #         base_url=base_url,
    #         model=model,
    #         messages=retry_messages,
    #         temperature=0.0,
    #         timeout=timeout,
    #     )
    #     is_valid, reason = validate_fixed_content(template_text, contract_text)
    #     if not is_valid:
    #         logger.error("二次生成后仍未通过校验: %s", reason)
    #         raise RuntimeError(
    #             "生成结果未通过固定内容校验，请更换模型或补充更明确变量信息。"
    #             f"原因: {reason}"
    #         )
    #     logger.info("二次生成后固定文本校验通过")
    logger.info("已跳过固定片段校验逻辑（临时注释）")

    if output_file:
        logger.info("步骤 5/5：写入输出文件 …")
        written_paths = write_contract_output(
            contract_text=contract_text,
            output_file=output_file,
            output_format=output_format,
            format_spec_dir=format_spec_dir,
            format_id=format_id,
            format_seed=format_seed,
        )
        logger.info("合同已写入: %s（字符数: %d）", written_paths, len(contract_text))

    return contract_text


def mutate_single_contract(
    base_contract_file: str,
    skill_file: str,
    basic_info_file: str,
    output_file: str | None = None,
    *,
    api_key: str | None = None,
    base_url: str = "https://openrouter.ai/api/v1",
    model: str = "deepseek/deepseek-v3.2",
    temperature: float = 0.4,
    timeout: int = 300,
    output_format: str = "md",
    format_spec_dir: str = "resources/formats",
    format_id: str | None = None,
    format_seed: int | None = None,
) -> str:
    resolved_api_key = (
        api_key
        or os.getenv("OPENROUTER_API_KEY")
        or os.getenv("LLM_API_KEY")
        or os.getenv("OPENAI_API_KEY")
    )
    if not resolved_api_key:
        raise RuntimeError(
            "未设置 API Key。请设置 OPENROUTER_API_KEY（或 LLM_API_KEY / OPENAI_API_KEY）"
        )

    logger.info("开始执行单份合同 mutation")
    logger.info("步骤 1/4：读取 mutation 要求文件 …")
    mutation_skill_text = read_text_file(skill_file)
    logger.info("步骤 2/4：读取基准合同 …")
    base_contract_text = read_template(base_contract_file)
    logger.info("步骤 3/4：读取合同基本信息 …")
    basic_info_text = read_text_file(basic_info_file)
    logger.info("步骤 4/4：构建 mutation 消息并调用大模型 …")
    messages = build_mutation_messages(
        mutation_skill_text=mutation_skill_text,
        base_contract_text=base_contract_text,
        basic_info_text=basic_info_text,
    )

    mutated_contract_text = call_llm(
        api_key=resolved_api_key,
        base_url=base_url,
        model=model,
        messages=messages,
        temperature=temperature,
        timeout=timeout,
    )
    logger.info("mutation 生成完成，合同正文长度: %d 字符", len(mutated_contract_text))
    logger.debug("mutation 生成的合同正文:\n%s", mutated_contract_text)

    if output_file:
        written_paths = write_contract_output(
            contract_text=mutated_contract_text,
            output_file=output_file,
            output_format=output_format,
            format_spec_dir=format_spec_dir,
            format_id=format_id,
            format_seed=format_seed,
        )
        logger.info(
            "mutation 合同已写入: %s（字符数: %d）",
            written_paths,
            len(mutated_contract_text),
        )
    return mutated_contract_text


def main() -> None:
    args = parse_args()
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        stream=sys.stderr,
    )
    logger.info("程序启动: gen_contract")
    logger.info(
        "参数: skill=%s | template=%s | basic_info=%s | output=%s | output_format=%s | model=%s | base_url=%s | temperature=%s | timeout=%s | log_level=%s",
        args.skill_file,
        args.template_file,
        args.basic_info_file,
        args.output_file,
        args.output_format,
        args.model,
        args.base_url,
        args.temperature,
        args.timeout,
        args.log_level,
    )

    api_key = (
        os.getenv("OPENROUTER_API_KEY")
        or os.getenv("LLM_API_KEY")
        or os.getenv("OPENAI_API_KEY")
    )
    if not api_key:
        print(
            "错误: 未设置环境变量 OPENROUTER_API_KEY（或 LLM_API_KEY / OPENAI_API_KEY）",
            file=sys.stderr,
        )
        sys.exit(1)
    key_source = (
        "OPENROUTER_API_KEY"
        if os.getenv("OPENROUTER_API_KEY")
        else ("LLM_API_KEY" if os.getenv("LLM_API_KEY") else "OPENAI_API_KEY")
    )
    tail = api_key[-4:] if len(api_key) >= 4 else "****"
    logger.info("已检测到 API Key（来源环境变量: %s，仅显示末尾 4 位: …%s）", key_source, tail)

    contract_text = gen_single_contract(
        skill_file=args.skill_file,
        template_file=args.template_file,
        basic_info_file=args.basic_info_file,
        output_file=args.output_file,
        api_key=api_key,
        model=args.model,
        base_url=args.base_url,
        temperature=args.temperature,
        timeout=args.timeout,
        output_format=args.output_format,
        format_spec_dir=args.format_spec_dir,
        format_id=args.format_id,
        format_seed=args.format_seed,
    )
    print(f"合同已生成: {args.output_file}")


if __name__ == "__main__":
    main()
