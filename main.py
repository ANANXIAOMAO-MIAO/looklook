"""
LOOKLOOK — FastAPI 后端：文献检索与智能速递流水线
"""

from __future__ import annotations

import base64
import json
import os
import re
import threading
import time
import uuid
from datetime import date, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any, Callable

import requests
import uvicorn
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

FREE_MODE = True
DEFAULT_DEEPSEEK_KEY = os.getenv("DEFAULT_DEEPSEEK_KEY", "")

DEEPSEEK_API_URL = "https://api.deepseek.com/chat/completions"
DEEPSEEK_MODEL = "deepseek-chat"
OPENALEX_SEARCH_URL = "https://api.openalex.org/works"
OPENALEX_PER_PAGE = 200
OPENALEX_SLEEP = 0.1
MAX_FETCH_PAPERS = 100
MAX_SELECTED_PAPERS = 10
MAX_RETRIES = 3
RETRY_INTERVAL = 2

TIME_RANGE_OPTIONS = ["一月内", "半年内", "一年内", "三年内", "五年内"]

TIME_RANGE_DAYS = {
    "一月内": 30,
    "半年内": 182,
    "一年内": 365,
    "三年内": 1095,
    "五年内": 1825,
}

# 异步任务进度存储 {task_id: {step, message, done, result, error}}
task_status: dict[str, dict[str, Any]] = {}
task_lock = threading.Lock()

# ---------------------------------------------------------------------------
# 流水线核心函数（自 app.py 提取，无 Streamlit 依赖）
# ---------------------------------------------------------------------------


def get_date_range(time_label: str) -> tuple[date, date]:
    """根据时间范围标签计算 start_date 和 end_date。"""
    if time_label not in TIME_RANGE_DAYS:
        raise ValueError(f"无效的时间范围: {time_label}")
    today = date.today()
    days = TIME_RANGE_DAYS[time_label]
    start = today - timedelta(days=days)
    return start, today


def call_deepseek(
    api_key: str,
    messages: list[dict[str, str]],
    *,
    json_mode: bool = True,
    temperature: float = 0.3,
) -> str:
    """调用 DeepSeek Chat API，带重试与错误处理。"""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    body: dict[str, Any] = {
        "model": DEEPSEEK_MODEL,
        "messages": messages,
        "temperature": temperature,
    }
    if json_mode:
        body["response_format"] = {"type": "json_object"}

    last_error: Exception | None = None
    for attempt in range(MAX_RETRIES):
        try:
            resp = requests.post(DEEPSEEK_API_URL, headers=headers, json=body, timeout=120)
            if resp.status_code == 401:
                raise ValueError("DeepSeek API Key 无效或已过期，请检查后重试。")
            if resp.status_code == 429:
                raise ValueError("DeepSeek API 请求过于频繁，请稍后再试。")
            if resp.status_code >= 400:
                raise RuntimeError(
                    f"DeepSeek API 错误 (HTTP {resp.status_code}): {resp.text[:300]}"
                )
            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            return content.strip()
        except ValueError:
            raise
        except Exception as e:
            last_error = e
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_INTERVAL)
    raise RuntimeError(f"DeepSeek API 调用失败（已重试 {MAX_RETRIES} 次）: {last_error}")


def parse_json_response(text: str) -> Any:
    """解析 DeepSeek 返回的 JSON，支持外层包裹对象。"""
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    parsed = json.loads(text)
    if isinstance(parsed, dict):
        for key in ("results", "papers", "items", "data", "relevant_papers", "batch"):
            if key in parsed and isinstance(parsed[key], list):
                return parsed[key]
        for v in parsed.values():
            if isinstance(v, list):
                return v
    return parsed


def expand_keywords(
    api_key: str, keywords: str, interest: str, journal_input: str
) -> tuple[list[str], str]:
    """关键词拓展、期刊翻译与 OpenAlex 广撒网搜索词生成。"""
    keywords = keywords.strip()
    journal_hint = journal_input.strip() if journal_input.strip() else "（用户未填写期刊）"
    keywords_note = (
        "（用户未填写关键词，请完全依据个人兴趣描述生成 simple_terms）"
        if not keywords
        else keywords
    )

    system_msg = (
        "你是学术文献检索专家，专门为 OpenAlex 设计「广撒网」式检索词。"
        "目标是尽可能多地召回可能相关的论文，因此搜索词要宽泛，包含同义词和近义词，"
        "不要使用双引号做短语精确匹配。"
        "必须返回合法 JSON，且只包含 simple_terms 和 journal_names 两个字段。"
        "无论用户是否提供关键词，都必须返回 simple_terms 字段。"
    )
    user_msg = f"""请根据以下信息返回 JSON 对象（仅 JSON，无其他文字）：

字段要求：
1. "simple_terms": 用于 OpenAlex search 参数的英文检索串。策略：广撒网、高召回。
   - 使用 3–5 个最核心的英文关键词，以空格分隔
   - 不要使用双引号包裹短语
   - 同义词、近义词用大写 OR 连接，例如：bias OR prejudice interpersonal OR social relationships
   - 若兴趣涉及多个维度，可用 OR 连接不同维度
   - 词组不要加引号；OR 必须大写
   - **若用户关键词为空**，则完全依据「个人兴趣描述」自动生成 simple_terms，
     范围要广泛，涵盖同义词、近义词，使用 OR 连接不同概念，确保 OpenAlex 能尽可能多地召回相关论文
2. "journal_names": 字符串数组。若用户填写了期刊（中英文均可），翻译为标准英文期刊全称；
   若未填写则返回空数组 []

用户关键词：{keywords_note}
个人兴趣描述：{interest}
用户期刊输入：{journal_hint}

返回格式示例：
{{"simple_terms": "bias OR prejudice interpersonal relationships", "journal_names": ["Nature"]}}"""

    content = call_deepseek(
        api_key,
        [{"role": "system", "content": system_msg}, {"role": "user", "content": user_msg}],
        json_mode=True,
    )
    data = json.loads(content)
    journal_names = data.get("journal_names", [])
    simple_terms = data.get("simple_terms", "")
    if not isinstance(journal_names, list):
        journal_names = []
    if not simple_terms or not str(simple_terms).strip():
        simple_terms = keywords if keywords else interest[:80]
    return [str(j) for j in journal_names], str(simple_terms).strip()


def fetch_papers_openalex(
    search_terms: str,
    start_date: str,
    end_date: str,
    journal_names: list[str],
    progress_callback: Callable[[int], None] | None = None,
    max_papers: int = MAX_FETCH_PAPERS,
) -> list[dict]:
    """使用 OpenAlex API 抓取论文。"""
    papers: list[dict] = []
    page = 1

    while True:
        params = {
            "search": search_terms,
            "filter": f"from_publication_date:{start_date},to_publication_date:{end_date}",
            "per_page": OPENALEX_PER_PAGE,
            "page": page,
        }

        if journal_names and page == 1:
            source_ids = []
            for jname in journal_names:
                jname = jname.strip()
                if len(jname) <= 1:
                    continue
                try:
                    source_url = "https://api.openalex.org/sources"
                    source_params = {"search": jname, "per_page": 3}
                    source_resp = None
                    for sa in range(3):
                        try:
                            source_resp = requests.get(
                                source_url, params=source_params, timeout=15
                            )
                            break
                        except requests.exceptions.ConnectionError:
                            if sa < 2:
                                time.sleep(2)
                            else:
                                source_resp = None
                    if source_resp is None or source_resp.status_code != 200:
                        continue
                    source_data = source_resp.json()
                    for src in source_data.get("results", []):
                        sid = src.get("id", "")
                        if sid:
                            short_id = sid.split("/")[-1] if "/" in sid else sid
                            source_ids.append(short_id)
                            break
                except Exception:
                    pass

            if source_ids:
                if len(source_ids) == 1:
                    source_filter = source_ids[0]
                else:
                    source_filter = "|".join(source_ids)
                existing_filter = params.get("filter", "")
                if existing_filter:
                    params["filter"] = (
                        existing_filter + f",primary_location.source.id:{source_filter}"
                    )
                else:
                    params["filter"] = f"primary_location.source.id:{source_filter}"

        resp = None
        data = None
        max_retries = 3
        retry_delay = 3
        for attempt in range(max_retries):
            try:
                resp = requests.get(OPENALEX_SEARCH_URL, params=params, timeout=30)
                resp.raise_for_status()
                data = resp.json()
                break
            except requests.exceptions.ConnectionError as ce:
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    continue
                raise RuntimeError(f"OpenAlex 连接失败（已重试{max_retries}次）: {ce}")
            except Exception as e:
                if resp is not None:
                    raise RuntimeError(
                        f"OpenAlex 请求失败 (HTTP {resp.status_code}): {resp.text[:300]}"
                    ) from e
                raise RuntimeError(f"OpenAlex 请求失败: {e}") from e

        for work in data.get("results", []):
            if len(papers) >= max_papers:
                break
            abstract = ""
            inverted_index = work.get("abstract_inverted_index")
            if inverted_index:
                try:
                    max_pos = max(max(positions) for positions in inverted_index.values())
                    words = [None] * (max_pos + 1)
                    for word, positions in inverted_index.items():
                        for pos in positions:
                            words[pos] = word
                    abstract = " ".join(filter(None, words))
                except Exception:
                    pass

            authors = ", ".join(
                [
                    a["author"]["display_name"]
                    for a in work.get("authorships", [])
                    if a.get("author")
                ]
            )
            try:
                journal = (
                    work.get("primary_location", {})
                    .get("source", {})
                    .get("display_name", "")
                )
            except AttributeError:
                journal = ""
            year = work.get("publication_year", "")
            citation_count = work.get("cited_by_count", 0)

            papers.append(
                {
                    "title": work.get("title", ""),
                    "authors": authors,
                    "abstract": abstract,
                    "journal": journal,
                    "citationCount": citation_count,
                    "year": int(year) if year else 0,
                    "paperId": work.get("id", ""),
                }
            )

        if progress_callback:
            progress_callback(len(papers))

        total_count = data.get("meta", {}).get("count", 0)
        if (
            len(papers) >= max_papers
            or len(papers) >= total_count
            or len(data.get("results", [])) == 0
        ):
            break
        page += 1
        time.sleep(OPENALEX_SLEEP)

    print(f" 实际抓取到的论文数: {len(papers)}")
    return papers


def filter_papers_batch(
    api_key: str, batch: list[dict], interest: str, batch_num: int
) -> list[dict]:
    """AI 智能筛选一批论文。"""
    lines = []
    for i, p in enumerate(batch):
        abstract = p.get("abstract") or "（无摘要）"
        lines.append(f"[{i}] 标题: {p['title']}\n摘要: {abstract[:800]}")
    papers_text = "\n\n".join(lines)

    system_msg = (
        "你是学术论文筛选助手。根据用户兴趣判断每篇论文是否高度相关。"
        "必须返回 JSON 对象，包含 results 数组。"
    )
    user_msg = f"""用户个人兴趣描述：
{interest}

以下论文（批次内序号从 0 开始）：
{papers_text}

请返回 JSON：
{{"results": [{{"index": 0, "relevant": true}}, {{"index": 1, "relevant": false}}, ...]}}

对每篇论文判断 relevant（true/false）。仅返回 JSON。"""

    content = call_deepseek(
        api_key,
        [{"role": "system", "content": system_msg}, {"role": "user", "content": user_msg}],
        json_mode=True,
    )
    parsed = json.loads(content)
    results = parsed.get("results", parsed) if isinstance(parsed, dict) else parsed
    if not isinstance(results, list):
        results = parse_json_response(content)
        if isinstance(results, dict):
            results = list(results.values())[0] if results else []

    relevant_indices = set()
    for item in results:
        if isinstance(item, dict) and item.get("relevant") is True:
            relevant_indices.add(item.get("index", -1))

    return [batch[i] for i in range(len(batch)) if i in relevant_indices]


def enrich_papers_batch(api_key: str, batch: list[dict], interest: str) -> list[dict]:
    """为一批论文生成中文信息与评分。"""
    lines = []
    for i, p in enumerate(batch):
        abstract = p.get("abstract") or "（无摘要）"
        lines.append(f"[{i}] 标题: {p['title']}\n摘要: {abstract[:600]}")

    system_msg = "你是学术论文分析助手。为每篇论文生成中文翻译、总结和评分。必须返回 JSON。"
    user_msg = f"""用户兴趣：{interest}

论文列表：
{chr(10).join(lines)}

返回 JSON：
{{"results": [
  {{
    "index": 0,
    "relevance_score": 0.85,
    "title_cn": "中文标题",
    "one_liner": "一句话总结（中文）",
    "recommendation_reason": "推荐理由（中文，约30字）"
  }}
]}}

relevance_score 为 0-1 浮点数。仅返回 JSON。"""

    content = call_deepseek(
        api_key,
        [{"role": "system", "content": system_msg}, {"role": "user", "content": user_msg}],
        json_mode=True,
    )
    parsed = json.loads(content)
    results = parsed.get("results", [])
    if not isinstance(results, list):
        results = parse_json_response(content)
        if not isinstance(results, list):
            results = []

    enrich_map: dict[int, dict] = {}
    for item in results:
        if isinstance(item, dict) and "index" in item:
            enrich_map[item["index"]] = item

    enriched = []
    for i, p in enumerate(batch):
        info = enrich_map.get(i, {})
        paper = dict(p)
        paper["relevance_score"] = float(info.get("relevance_score", 0.5))
        paper["title_cn"] = info.get("title_cn") or p["title"]
        paper["one_liner"] = info.get("one_liner", "")
        paper["recommendation_reason"] = info.get("recommendation_reason", "")
        enriched.append(paper)
    return enriched


def compute_scores_and_select(
    candidate_papers: list[dict],
) -> tuple[list[dict], list[dict]]:
    """计算推荐指数、排序并分为精选与其他。"""
    counts = [p.get("citationCount", 0) or 0 for p in candidate_papers]
    min_c, max_c = min(counts), max(counts)
    if max_c == min_c:
        norm_citations = [1.0] * len(counts)
    else:
        norm_citations = [(c - min_c) / (max_c - min_c) for c in counts]

    for p, norm_c in zip(candidate_papers, norm_citations):
        rel = float(p.get("relevance_score", 0.5))
        p["norm_citation"] = norm_c
        p["score"] = round((rel * 0.5 + norm_c * 0.5) * 10, 1)

    sorted_papers = sorted(candidate_papers, key=lambda x: x["score"], reverse=True)
    selected = sorted_papers[:MAX_SELECTED_PAPERS]
    other = sorted_papers[MAX_SELECTED_PAPERS:]
    return selected, other


def generate_review(api_key: str, selected_papers: list[dict]) -> str:
    """生成中文综述。"""
    lines = []
    for p in selected_papers:
        lines.append(f"- {p.get('title_cn', p['title'])}：{p.get('one_liner', '')}")
    summary_input = "\n".join(lines)

    system_msg = "你是学术综述撰写专家，用流畅中文撰写文献速递综述。"
    user_msg = f"""根据以下精选论文的中文标题与一句话总结，撰写中文文献速递综述。
直接输出综述正文，不要额外标题。综述必须包含以下三部分：

1. 以「本期研究围绕……」开头，概括本期研究方向
2. 以「覆盖……方向」描述主要主题分布
3. 以「值得关注的结论有：」引出 5 条结论，每条以数字序号开头（1. 2. 3. 4. 5.）

总字数 200-400 字，语言流畅自然。

论文信息：
{summary_input}"""

    return call_deepseek(
        api_key,
        [{"role": "system", "content": system_msg}, {"role": "user", "content": user_msg}],
        json_mode=False,
        temperature=0.5,
    )


def build_excel(review: str, selected: list[dict], other: list[dict]) -> bytes:
    """生成三 Sheet 的 Excel 文件。"""
    wb = Workbook()
    bold = Font(bold=True)
    wrap = Alignment(wrap_text=True, vertical="top")

    ws1 = wb.active
    ws1.title = "文献速递"
    ws1.merge_cells("A1:J1")
    cell = ws1["A1"]
    cell.value = review
    cell.alignment = wrap
    ws1.row_dimensions[1].height = 120
    ws1["A3"] = f"以上综述基于精选的 {len(selected)} 篇论文，详情见 Sheet2。"

    ws2 = wb.create_sheet("精选论文数据")
    headers2 = [
        "推荐指数",
        "英文标题",
        "中文标题",
        "作者",
        "期刊",
        "发表年份",
        "引用次数",
        "英文摘要",
        "一句话总结（中文）",
        "推荐理由（中文）",
    ]
    col_widths2 = [10, 30, 30, 20, 20, 8, 8, 50, 40, 40]
    for col, (h, w) in enumerate(zip(headers2, col_widths2), 1):
        c = ws2.cell(row=1, column=col, value=h)
        c.font = bold
        ws2.column_dimensions[get_column_letter(col)].width = w

    for row_idx, p in enumerate(selected, 2):
        ws2.cell(row=row_idx, column=1, value=p.get("score", 0))
        ws2.cell(row=row_idx, column=2, value=p.get("title", ""))
        ws2.cell(row=row_idx, column=3, value=p.get("title_cn", ""))
        ws2.cell(row=row_idx, column=4, value=p.get("authors", ""))
        ws2.cell(row=row_idx, column=5, value=p.get("journal", ""))
        ws2.cell(row=row_idx, column=6, value=p.get("year", ""))
        ws2.cell(row=row_idx, column=7, value=p.get("citationCount", 0))
        c8 = ws2.cell(row=row_idx, column=8, value=p.get("abstract", ""))
        c8.alignment = wrap
        c9 = ws2.cell(row=row_idx, column=9, value=p.get("one_liner", ""))
        c9.alignment = wrap
        c10 = ws2.cell(row=row_idx, column=10, value=p.get("recommendation_reason", ""))
        c10.alignment = wrap

    ws3 = wb.create_sheet("其他相关论文")
    for col, h in enumerate(["英文标题", "中文标题"], 1):
        c = ws3.cell(row=1, column=col, value=h)
        c.font = bold
        ws3.column_dimensions[get_column_letter(col)].width = 40

    if other:
        for row_idx, p in enumerate(other, 2):
            ws3.cell(row=row_idx, column=1, value=p.get("title", ""))
            ws3.cell(row=row_idx, column=2, value=p.get("title_cn", ""))
    else:
        ws3["A2"] = "无"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def format_paper_for_api(p: dict) -> dict:
    """将内部论文字典格式化为 API 响应字段。"""
    return {
        "title": p.get("title", ""),
        "title_cn": p.get("title_cn", ""),
        "authors": p.get("authors", ""),
        "journal": p.get("journal", ""),
        "year": p.get("year", 0),
        "citation_count": p.get("citationCount", 0),
        "abstract": p.get("abstract", ""),
        "one_liner": p.get("one_liner", ""),
        "recommendation_reason": p.get("recommendation_reason", ""),
        "score": p.get("score", 0),
    }


def run_pipeline(
    api_key: str,
    keywords: str,
    journal_input: str,
    time_label: str,
    interest: str,
    status_callback: Callable[[int, str], None] | None = None,
) -> dict[str, Any]:
    """
    主处理流程，返回完整结果字典。
    status_callback: 接收 (step 1-5, message) 用于任务进度更新。
    """

    def report(step: int, msg: str) -> None:
        print(msg)
        if status_callback:
            status_callback(step, msg)

    keywords = keywords or ""
    journal_input = journal_input or ""

    start_date, end_date = get_date_range(time_label)
    start_date_str = start_date.strftime("%Y-%m-%d")
    end_date_str = end_date.strftime("%Y-%m-%d")

    report(1, " 正在拆解兴趣，拓展搜索词...")
    journal_names, simple_terms = expand_keywords(
        api_key, keywords, interest, journal_input
    )

    journal_names_for_fetch = journal_names if journal_input.strip() else []

    def on_fetch_progress(count: int) -> None:
        report(2, f" 正在从 OpenAlex 抓取论文...（已获取 {count} 篇）")

    report(2, " 正在从 OpenAlex 抓取论文...")
    raw_papers = fetch_papers_openalex(
        simple_terms,
        start_date_str,
        end_date_str,
        journal_names_for_fetch,
        progress_callback=on_fetch_progress,
    )

    if not raw_papers:
        raise ValueError("未找到符合日期或期刊条件的论文，请调整关键词或时间范围后重试。")

    report(3, " AI 正在筛选相关论文...")
    candidate_papers: list[dict] = []
    batch_size_filter = 10
    total_batches = (len(raw_papers) + batch_size_filter - 1) // batch_size_filter
    for b in range(total_batches):
        batch = raw_papers[b * batch_size_filter : (b + 1) * batch_size_filter]
        filtered = filter_papers_batch(api_key, batch, interest, b + 1)
        candidate_papers.extend(filtered)
        report(
            3,
            f" AI 正在筛选相关论文...（第 {b + 1}/{total_batches} 批，已保留 {len(candidate_papers)} 篇）",
        )
        time.sleep(0.5)

    if not candidate_papers:
        raise ValueError("未找到高度相关论文，请调整兴趣描述后重试。")

    report(4, " 正在生成总结与评分...")
    enriched_all: list[dict] = []
    batch_size_enrich = 5
    total_enrich = (len(candidate_papers) + batch_size_enrich - 1) // batch_size_enrich
    for b in range(total_enrich):
        batch = candidate_papers[b * batch_size_enrich : (b + 1) * batch_size_enrich]
        enriched = enrich_papers_batch(api_key, batch, interest)
        enriched_all.extend(enriched)
        report(4, f" 正在生成总结与评分...（第 {b + 1}/{total_enrich} 批）")
        time.sleep(0.5)

    selected_papers, other_papers = compute_scores_and_select(enriched_all)

    report(5, " 正在生成综述...")
    review = generate_review(api_key, selected_papers)

    report(5, " 正在生成 Excel...")
    excel_bytes = build_excel(review, selected_papers, other_papers)

    journals_used: list[str] = []
    if journal_input.strip():
        journals_used = [j.strip() for j in journal_input.split(",") if j.strip()]
    elif journal_names:
        journals_used = journal_names

    return {
        "summary": review,
        "selected_papers": [format_paper_for_api(p) for p in selected_papers],
        "other_papers_count": len(other_papers),
        "simple_terms": simple_terms,
        "journals_used": journals_used,
        "time_range": time_label,
        "interest": interest,
        "excel_bytes": excel_bytes,
    }


# ---------------------------------------------------------------------------
# FastAPI 应用
# ---------------------------------------------------------------------------

app = FastAPI(title="LOOKLOOK API", description="文献检索与智能速递后端")

TEMPLATES_DIR = Path(__file__).parent / "templates"


class SearchRequest(BaseModel):
    api_key: str = Field(default="", description="DeepSeek API Key（免 Key 模式可省略）")
    interest: str = Field(..., description="个人兴趣描述")
    keywords: str = Field(default="", description="可选关键词")
    journal_input: str = Field(default="", description="可选期刊名")
    time_range: str = Field(default="三年内", description="时间范围")


class TaskStartResponse(BaseModel):
    task_id: str


class TaskStatusResponse(BaseModel):
    step: int
    message: str
    done: bool
    result: dict[str, Any] | None = None
    error: str | None = None


def _update_task(
    task_id: str,
    *,
    step: int | None = None,
    message: str | None = None,
    done: bool | None = None,
    result: dict[str, Any] | None = None,
    error: str | None = None,
) -> None:
    with task_lock:
        if task_id not in task_status:
            return
        if step is not None:
            task_status[task_id]["step"] = step
        if message is not None:
            task_status[task_id]["message"] = message
        if done is not None:
            task_status[task_id]["done"] = done
        if result is not None:
            task_status[task_id]["result"] = result
        if error is not None:
            task_status[task_id]["error"] = error


def _run_search_task(task_id: str, req: SearchRequest) -> None:
    """后台线程执行流水线，更新 task_status。"""

    def on_status(step: int, message: str) -> None:
        _update_task(task_id, step=step, message=message)

    try:
        _update_task(task_id, step=0, message="正在准备...")
        result = run_pipeline(
            api_key=req.api_key.strip(),
            keywords=req.keywords.strip() if req.keywords else "",
            journal_input=req.journal_input.strip() if req.journal_input else "",
            time_label=req.time_range,
            interest=req.interest.strip(),
            status_callback=on_status,
        )
        excel_bytes = result.pop("excel_bytes")
        result["excel_base64"] = base64.b64encode(excel_bytes).decode("utf-8")
        _update_task(
            task_id,
            step=5,
            message=" 分析完成！",
            done=True,
            result=result,
        )
    except ValueError as e:
        _update_task(task_id, done=True, error=str(e), message=str(e))
    except Exception as e:
        _update_task(
            task_id,
            done=True,
            error=f"流水线执行失败: {e}",
            message=f"流水线执行失败: {e}",
        )


@app.get("/")
def index_page():
    """返回前端单页应用。"""
    html_path = TEMPLATES_DIR / "index.html"
    if not html_path.is_file():
        raise HTTPException(status_code=404, detail="前端页面未找到")
    return FileResponse(html_path, media_type="text/html; charset=utf-8")


@app.get("/health")
def health():
    return {"status": "healthy"}


@app.get("/api/mode")
def api_mode():
    """返回当前是否为免 Key 模式。"""
    return {"free_mode": FREE_MODE}


def _resolve_api_key(client_api_key: str) -> str:
    """免 Key 模式使用服务端默认 Key，否则要求客户端提供。"""
    if FREE_MODE:
        key = DEFAULT_DEEPSEEK_KEY.strip()
        if not key:
            raise HTTPException(
                status_code=500,
                detail="服务器未配置默认 DeepSeek API Key，请联系管理员",
            )
        return key
    if not client_api_key or not client_api_key.strip():
        raise HTTPException(status_code=400, detail="请提供有效的 DeepSeek API Key")
    return client_api_key.strip()


@app.post("/api/search", response_model=TaskStartResponse)
def api_search(req: SearchRequest):
    """创建异步检索任务，立即返回 task_id。"""
    effective_api_key = _resolve_api_key(req.api_key)
    if not req.interest or not req.interest.strip():
        raise HTTPException(status_code=400, detail="请提供个人兴趣描述")
    if req.time_range not in TIME_RANGE_OPTIONS:
        raise HTTPException(
            status_code=400,
            detail=f"无效的时间范围，可选: {', '.join(TIME_RANGE_OPTIONS)}",
        )

    task_id = str(uuid.uuid4())
    with task_lock:
        task_status[task_id] = {
            "step": 0,
            "message": "任务已创建，等待执行...",
            "done": False,
            "result": None,
            "error": None,
        }

    req.api_key = effective_api_key
    thread = threading.Thread(target=_run_search_task, args=(task_id, req), daemon=True)
    thread.start()
    return TaskStartResponse(task_id=task_id)


@app.get("/api/search/status", response_model=TaskStatusResponse)
def api_search_status(task_id: str):
    """查询异步任务进度与结果。"""
    with task_lock:
        if task_id not in task_status:
            raise HTTPException(status_code=404, detail="任务不存在或已过期")
        task = task_status[task_id].copy()

    return TaskStatusResponse(
        step=task.get("step", 0),
        message=task.get("message", ""),
        done=task.get("done", False),
        result=task.get("result"),
        error=task.get("error"),
    )


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
