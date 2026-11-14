#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
引文啄木鸟·参考文献格式自检官（reference-woodpecker）
GB/T 7714-2015 顺序编码制参考文献格式确定性检查引擎。

- 纯 Python 标准库实现，无第三方依赖。
- 只做可确定性判定的格式检查；无法确定性判定的项不强行报错。
- 结论分级：ERROR（错误）/ WARN（警告）/ INFO（存疑）/ PASS（通过）。
- 支持输入：.txt / .md / .docx 文件路径，或 "-" 读 stdin；--refs-only 表示输入本身就是文献列表。

用法：
    python check_refs.py <input> [--refs-only] [--report OUT.md] [--json OUT.json] [--lang auto]

退出码：0 = 检查正常完成（无论是否发现格式问题）；2 = 输入/读取错误。
"""

import argparse
import html
import json
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------- 常量

LEVEL_ERROR = "ERROR"
LEVEL_WARN = "WARN"
LEVEL_INFO = "INFO"
LEVEL_PASS = "PASS"

LEVEL_ORDER = {LEVEL_INFO: 0, LEVEL_PASS: 1, LEVEL_WARN: 2, LEVEL_ERROR: 3}

# GB/T 7714-2015 附录 B 文献类型标识
BASE_TYPES = {"M", "C", "G", "N", "J", "D", "R", "S", "P", "A", "CM", "DS", "DB", "CP", "EB"}
CARRIER = {"OL", "MT", "DK", "CD"}

# 条目起始：半角/全角方括号 + 数字（支持全角数字）
FW_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")
ENTRY_RE = re.compile(r"^\s*([\[［])\s*([0-9０-９]+)\s*([\]］])\s*(.*)$")
HEADING_RE = re.compile(r"^\s*(?:[#*\-–—>\s]*)?(参\s*考\s*文\s*献|文\s*献|references|bibliography|参考文献[:：]?)\s*[:：]?\s*$", re.I)
TYPE_RE = re.compile(r"\[([A-Za-z]{1,2})(?:/([A-Za-z]{1,2}))?\]")
URL_RE = re.compile(r"(https?://|www\.)\S+", re.I)
CITED_DATE_RE = re.compile(r"\[\d{4}-\d{1,2}-\d{1,2}\]")
STD_NO_RE = re.compile(r"\b[A-Z]{1,4}(?:/T)?\s*\d+(?:\.\d+)?(?:-\d{2,4})?\b")
PATENT_NO_RE = re.compile(r"\b[A-Z]{1,3}\s?\d{5,}[A-Z]?\b")
YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")
NEWSPAPER_DATE_RE = re.compile(r"\b(?:19|20)\d{2}-\d{1,2}-\d{1,2}\b")
PAGE_BAD_DASH_RE = re.compile(r"(\d+)\s*([~—–、，,]|／)\s*(\d+)\s*\.?\s*$")
PAGE_RE = re.compile(r"[:：]\s*\d+(\s*-\s*\d+)?\s*[.。]?\s*$")
CITE_RE = re.compile(r"\[(\d{1,3}(?:\s*[,，\-–]\s*\d{1,3})*)\]")
MAX_ENTRIES = 500

FULLWIDTH_PUNCT = "，、：；（）“”‘’"


# ---------------------------------------------------------------- 读取与提取

def read_text_file(path: Path) -> str:
    data = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def extract_docx_text(path: Path) -> str:
    """从 .docx 中按段落提取纯文本（仅 word/document.xml，w:t 节点拼接）。"""
    with zipfile.ZipFile(str(path)) as z:
        xml = z.read("word/document.xml").decode("utf-8", errors="replace")
    xml = xml.replace("<w:tab/>", " ").replace("<w:br/>", "\n")
    paras = re.split(r"</w:p>", xml)
    lines = []
    for p in paras:
        ts = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)
        line = html.unescape("".join(ts)).strip()
        if line:
            lines.append(line)
    return "\n".join(lines)


def load_input(path_str: str):
    """返回 (text, source_desc, error_message)。"""
    if path_str == "-":
        text = sys.stdin.read()
        return text, "stdin", None
    path = Path(path_str)
    if not path.exists():
        return None, str(path), "输入文件不存在: %s" % path
    if not path.is_file():
        return None, str(path), "输入路径不是文件: %s" % path
    suffix = path.suffix.lower()
    try:
        if suffix == ".docx":
            return extract_docx_text(path), str(path), None
        if suffix in (".doc",):
            return None, str(path), "暂不支持旧版 .doc 格式，请先另存为 .docx 或纯文本"
        return read_text_file(path), str(path), None
    except zipfile.BadZipFile:
        return None, str(path), ".docx 文件损坏或不是有效的 zip 容器"
    except Exception as e:  # noqa: BLE001 读取失败要如实报告
        return None, str(path), "读取失败: %s" % e


# ---------------------------------------------------------------- 文献节定位

def locate_references(text: str, refs_only: bool):
    """返回 (body_text, refs_text, locate_mode, extra_notes)。"""
    notes = []
    if refs_only:
        return "", text, "refs-only", notes

    lines = text.splitlines()
    heading_idx = None
    for i, ln in enumerate(lines):
        if ln.strip() and HEADING_RE.match(ln.strip()):
            heading_idx = i
            break

    if heading_idx is not None:
        body = "\n".join(lines[:heading_idx])
        refs = "\n".join(lines[heading_idx + 1:])
        return body, refs, "heading", notes

    # 宽容模式：寻找从 [1] 开始、条目数 >=2 的编号块
    starts = []
    for i, ln in enumerate(lines):
        m = ENTRY_RE.match(ln)
        if m and int(m.group(2).translate(FW_DIGITS)) == 1:
            starts.append(i)
    for s in starts:
        count = sum(1 for ln in lines[s:] if ENTRY_RE.match(ln))
        if count >= 2:
            body = "\n".join(lines[:s])
            refs = "\n".join(lines[s:])
            notes.append("未找到“参考文献”标题，已按首个编号条目块定位文献列表（宽容模式）")
            return body, refs, "tolerant", notes

    notes.append("未能定位参考文献列表（无标题、无连续编号条目块）")
    return text, "", "none", notes


# ---------------------------------------------------------------- 条目切分

def split_entries(refs_text: str):
    """把文献区文本切分为条目。返回 (entries, global_findings)。
    entry = {"no": int|None, "raw": str, "seq_note": str|None}"""
    entries = []
    g_findings = []
    stray = []
    cur = None

    for raw_line in refs_text.splitlines():
        line = raw_line.rstrip()
        m = ENTRY_RE.match(line)
        if m:
            if cur:
                entries.append(cur)
                cur = None
            no_raw = m.group(2).translate(FW_DIGITS)
            no = int(no_raw)
            bracket = m.group(1)
            note = None
            if bracket == "［":
                note = "序号使用了全角方括号［］，应使用半角 []"
            elif not no_raw.isascii():
                note = "序号使用了全角数字，应使用半角阿拉伯数字"
            cur = {"no": no, "raw": line, "seq_note": note, "lines": [line]}
        else:
            if line.strip():
                if cur:
                    cur["lines"].append(line.strip())
                    cur["raw"] = "\n".join(cur["lines"])
                else:
                    stray.append(line.strip())
    if cur:
        entries.append(cur)

    if stray:
        g_findings.append({
            "level": LEVEL_INFO, "code": "I_EXTRA_TEXT",
            "message": "文献列表中混入 %d 行不以编号开头的文字（可能为标题残留或说明文字）：%s" % (
                len(stray), stray[0][:60]),
        })
    for e in entries:
        e["raw"] = e["raw"].strip()
        del e["lines"]
    return entries, g_findings


# ---------------------------------------------------------------- 单条检查

def _entry_verdict(findings):
    lv = LEVEL_PASS
    for f in findings:
        if LEVEL_ORDER[f["level"]] > LEVEL_ORDER[lv]:
            lv = f["level"]
    return lv


def _add(findings, level, code, message, advice=""):
    findings.append({"level": level, "code": code, "message": message, "advice": advice})


def _strip_quoted(tail: str) -> str:
    """去掉《》与 “”引号段，返回剩余部分（用于结构标点检查，避免误伤照录题名）。"""
    t = re.sub(r"《[^》]*》", "", tail)
    t = re.sub(r"“[^”]*”", "", t)
    return t


def _find_type(raw: str):
    """返回 (type_code, carrier, match, is_lower, unknown)。找不到返回 (None, None, None, False, False)。"""
    m = TYPE_RE.search(raw)
    if m:
        base_raw, carrier_raw = m.group(1), m.group(2)
        base = base_raw.upper()
        carrier = carrier_raw.upper() if carrier_raw else None
        is_lower = base_raw != base or bool(carrier_raw and carrier_raw != carrier)
        unknown = base not in BASE_TYPES or bool(carrier and carrier not in CARRIER)
        return base, carrier, m, is_lower, unknown
    return None, None, None, False, False


def check_entry(entry):
    """对单条文献执行确定性检查，返回填充 findings 后的 entry。"""
    raw = entry["raw"]
    findings = []
    no = entry["no"]
    if entry.get("seq_note"):
        _add(findings, LEVEL_WARN, "W_SEQ_FULLWIDTH", entry["seq_note"],
             "将序号改为半角方括号和半角阿拉伯数字，如 [3]")
    if no is None:
        _add(findings, LEVEL_ERROR, "E_NO_SEQ", "条目缺少 [数字] 序号",
             "顺序编码制要求每条以 [1] [2] … 连续编号")

    base, carrier, tmatch, is_lower, unknown = _find_type(raw)

    # ---- 类型标识 ----
    if tmatch is None:
        _add(findings, LEVEL_ERROR, "E_NO_TYPE", "未找到文献类型标识（如 [J]、[M]、[D]、[C]、[S]、[N]、[R]、[P]、[EB/OL]）",
             "在题名之后补写合法类型标识，见 GB/T 7714-2015 附录 B")
        base, carrier = None, None
    elif is_lower:
        _add(findings, LEVEL_ERROR, "E_TYPE_LOWER",
             "文献类型标识使用了小写字母（%s）" % tmatch.group(0),
             "类型标识应为大写，如 [J]、[M]、[EB/OL]")
    elif unknown:
        _add(findings, LEVEL_WARN, "W_TYPE_UNKNOWN",
             "类型标识 %s 不在 GB/T 7714-2015 常用标识表中" % tmatch.group(0),
             "核对附录 B：M C G N J D R S P A CM DS DB CP EB，载体 OL/MT/DK/CD")
    tail = raw[tmatch.end():] if tmatch else raw

    # ---- 条目结尾与中文句号 ----
    stripped = raw.rstrip()
    if stripped.endswith("。"):
        _add(findings, LEVEL_ERROR, "E_ZH_PERIOD", "条目以中文句号“。”结尾，著录应用半角句点",
             "把结尾句号改为半角“.”")
    elif not stripped.endswith("."):
        _add(findings, LEVEL_ERROR, "E_NO_END_PERIOD", "条目未以半角句点“.”结尾",
             "每条文献应以“.”结尾")
    if "。" in stripped[:-1]:
        _add(findings, LEVEL_ERROR, "E_ZH_PERIOD_INSIDE", "条目内部出现中文句号“。”",
             "著录标点使用半角；如系题名照录含“。”，请人工确认后保留")

    # ---- 类型标识前空格 ----
    if tmatch and tmatch.start() > 0 and raw[tmatch.start() - 1] == " ":
        _add(findings, LEVEL_WARN, "W_SPACE_BEFORE_TYPE", "题名与类型标识之间存在空格（“%s…”前）" % tmatch.group(0),
             "规范写法为题名后紧跟类型标识，如：题名[J].")

    # ---- 作者区（第一个“. ”之前）----
    first_dot = re.search(r"\.\s", raw)
    if tmatch and first_dot is None:
        _add(findings, LEVEL_ERROR, "E_NO_AUTHOR_SEP", "未找到“主要责任者. 题名”之间的句点分隔",
             "作者与题名之间用“. ”分隔")
    elif first_dot:
        author_zone = raw[:first_dot.start()]
        if "、" in author_zone:
            _add(findings, LEVEL_ERROR, "E_DUNHAO_AUTHOR", "作者之间使用了顿号“、”",
                 "GB/T 7714 著录用“,”分隔作者，如：张三, 李四, 王五")
        if "，" in author_zone:
            _add(findings, LEVEL_ERROR, "E_FULLWIDTH_AUTHOR", "作者区使用了全角逗号“，”",
                 "改用半角逗号加空格“, ”分隔作者")

    # ---- 出版项区全角标点（去掉照录引号段后检查）----
    if tmatch:
        structural = _strip_quoted(tail)
        bad = [ch for ch in FULLWIDTH_PUNCT if ch in structural]
        if bad:
            _add(findings, LEVEL_ERROR, "E_FULLWIDTH_PUNCT",
                 "出版项区使用了全角标点：%s" % " ".join(sorted(set(bad))),
                 "著录标点用半角并后跟一个空格，如：北京: 出版社, 2019: 1-5.")

    # ---- 分类型检查 ----
    if base == "J":
        if not YEAR_RE.search(tail):
            _add(findings, LEVEL_ERROR, "E_MISSING_YEAR", "期刊条目缺出版年（4 位数字）",
                 "补写出版年，如：刊名, 2020, 43(5): 101-110.")
        if "年" in _strip_quoted(tail):
            _add(findings, LEVEL_WARN, "W_YEAR_FORMAT", "出版项区出现“年”字（如“2020年”）",
                 "著录年份只写 4 位数字，不带“年”字")
        if PAGE_BAD_DASH_RE.search(tail):
            m = PAGE_BAD_DASH_RE.search(tail)
            _add(findings, LEVEL_ERROR, "E_PAGE_RANGE_DASH",
                 "页码范围分隔符“%s”不规范" % m.group(2),
                 "起止页码用半字线“-”，如 101-110")
        elif not PAGE_RE.search(tail):
            _add(findings, LEVEL_WARN, "W_MISSING_PAGE", "未检测到“: 起止页码”",
                 "期刊条目建议著录页码，如：43(5): 101-110.（在线优先发表无页码可人工确认）")

    elif base == "M":
        if not YEAR_RE.search(tail):
            _add(findings, LEVEL_ERROR, "E_MISSING_YEAR", "专著条目缺出版年")
        if not re.search(r"[^,.:：]{1,30}[:：]\s*[^,:：]{1,40}[,，]\s*(?:19|20)\d{2}", _strip_quoted(tail)):
            _add(findings, LEVEL_WARN, "W_MISSING_PUBLISHER", "未检测到“出版地: 出版者, 出版年”结构",
                 "专著需著录出版地与出版者，如：北京: 清华大学出版社, 2019.")

    elif base == "D":
        if not YEAR_RE.search(tail):
            _add(findings, LEVEL_ERROR, "E_MISSING_YEAR", "学位论文条目缺年份")
        if not re.search(r"[^,.:：]{1,30}[:：]", _strip_quoted(tail)):
            _add(findings, LEVEL_WARN, "W_MISSING_UNIVERSITY", "未检测到“保存地: 保存单位”结构",
                 "学位论文需著录保存单位，如：北京: 北京航空航天大学, 2021.")

    elif base == "N":
        if not NEWSPAPER_DATE_RE.search(tail):
            _add(findings, LEVEL_ERROR, "E_MISSING_DATE", "报纸条目缺出版日期（YYYY-MM-DD）",
                 "如：中国教育报, 2021-06-15(8).")

    elif base == "R":
        if not YEAR_RE.search(tail):
            _add(findings, LEVEL_WARN, "W_MISSING_YEAR", "报告条目未检测到出版年")

    elif base == "P":
        if not PATENT_NO_RE.search(raw):
            _add(findings, LEVEL_WARN, "W_MISSING_PATENT_NO", "未检测到专利号（如 CN110123456A）",
                 "专利条目应著录专利号：专利题名: 专利号[P]. 公开日期.")

    elif base == "S":
        if not STD_NO_RE.search(raw):
            _add(findings, LEVEL_WARN, "W_NO_STD_NO", "未检测到标准编号（如 GB/T 7714-2015）",
                 "标准条目应含标准编号，如：标准名称: GB/T 7714-2015[S].")

    # ---- 电子资源（OL 载体或 EB 类型）----
    if (carrier == "OL") or (base == "EB"):
        if not URL_RE.search(tail):
            _add(findings, LEVEL_ERROR, "E_NO_URL", "电子资源条目缺获取和访问路径（URL）",
                 "补写 http:// 或 https:// 开头的访问路径")
        if not CITED_DATE_RE.search(tail):
            _add(findings, LEVEL_WARN, "W_NO_CITED_DATE", "未检测到引用日期 [YYYY-MM-DD]",
                 "联机文献建议著录引用日期，如 [2021-05-06]")

    entry["type"] = (base + ("/" + carrier if carrier else "")) if base else None
    entry["findings"] = findings
    entry["verdict"] = _entry_verdict(findings)
    return entry


# ---------------------------------------------------------------- 全局检查

def check_sequence(entries):
    findings = []
    nos = [e["no"] for e in entries if e["no"] is not None]
    if nos:
        if nos[0] != 1:
            findings.append({"level": LEVEL_ERROR, "code": "E_SEQ_NOT_START",
                             "message": "编号从 [%d] 开始，顺序编码制应从 [1] 开始" % nos[0],
                             "advice": "把首条编号改为 [1]，后续依次连续编号"})
        prev = None
        for n in nos:
            if prev is not None and n == prev:
                findings.append({"level": LEVEL_ERROR, "code": "E_SEQ_DUP",
                                 "message": "编号 [%d] 重复" % n,
                                 "advice": "删除重复条目或重新连续编号"})
            if prev is not None and n > prev + 1:
                findings.append({"level": LEVEL_ERROR, "code": "E_SEQ_SKIP",
                                 "message": "编号从 [%d] 跳到 [%d]，缺少 %s" % (
                                     prev, n, "、".join("[%d]" % k for k in range(prev + 1, n))),
                                 "advice": "补齐缺失条目或重新连续编号"})
            prev = n
    return findings


def expand_cites(body: str):
    """提取正文中的数字引用并展开。返回 (cited_order, suspects)。"""
    text = CITED_DATE_RE.sub(" ", body)
    order = []
    suspects = []
    seen = set()
    for m in CITE_RE.finditer(text):
        seg = m.group(1)
        nums = []
        parts = re.split(r"([,，\-–])", seg)
        cur = None
        rng_flag = False
        for p in parts:
            p = p.strip()
            if p in (",", "，", "-", "–"):
                rng_flag = p in ("-", "–")
                continue
            if not p.isdigit():
                continue
            v = int(p)
            if cur is not None and rng_flag:
                a, b = cur, v
                if 0 < b - a <= 50:
                    nums.extend(range(a, b + 1))
                else:
                    suspects.append(seg)
                cur = None
                rng_flag = False
            else:
                nums.append(v)
                cur = v
                rng_flag = False
        for n in nums:
            if n not in seen:
                seen.add(n)
                order.append(n)
    return order, suspects


def check_citation(body: str, entries):
    findings = []
    if not body.strip():
        findings.append({"level": LEVEL_INFO, "code": "I_NO_BODY",
                         "message": "未提供正文文本，跳过“正文引用与文献表对应”检查",
                         "advice": "如需检查引用对应关系，请提供含正文的完整文档"})
        return findings, []
    entry_nos = {e["no"] for e in entries if e["no"] is not None}
    order, suspects = expand_cites(body)
    cited = set(order)
    for n in sorted(cited - entry_nos):
        findings.append({"level": LEVEL_ERROR, "code": "E_CITE_NOT_FOUND",
                         "message": "正文引用了 [%d]，但文献表中没有该条目" % n,
                         "advice": "在文献表中补写该文献，或修正正文引用编号"})
    not_cited = sorted(n for n in entry_nos if n not in cited)
    if not_cited:
        findings.append({"level": LEVEL_WARN, "code": "W_NOT_CITED",
                         "message": "文献表条目 %s 未被正文引用" % "、".join("[%d]" % n for n in not_cited),
                         "advice": "顺序编码制只列正文引用过的文献；删除未引用条目或补齐正文引用"})
    if order:
        first_bad = None
        for i in range(1, len(order)):
            if order[i] < order[i - 1]:
                first_bad = i
                break
        if first_bad:
            findings.append({"level": LEVEL_WARN, "code": "W_CITE_ORDER",
                             "message": "正文首次引用顺序未按递增排列（[%d] 先于 [%d] 出现）" % (
                                 order[first_bad], order[first_bad - 1]),
                             "advice": "顺序编码制按正文首次引用先后编号，请按引用顺序重排文献表"})
    for seg in sorted(set(suspects)):
        findings.append({"level": LEVEL_INFO, "code": "I_CITE_SUSPECT",
                         "message": "正文数值引用 [%s] 展开范围异常，请人工确认是否为引用" % seg,
                         "advice": "确认该处是引用还是年份/数值"})
    return findings, order


def check_duplicates(entries):
    findings = []
    seen = {}
    for e in entries:
        key = re.sub(r"^\s*\[\d+\]\s*", "", e["raw"]).strip()
        key = re.sub(r"\s+", "", key)
        if key in seen:
            findings.append({"level": LEVEL_WARN, "code": "W_DUP_ENTRY",
                             "message": "条目 [%d] 与 [%d] 内容重复" % (seen[key], e["no"]),
                             "advice": "删除重复条目并重新编号"})
        else:
            seen[key] = e["no"]
    return findings


# ---------------------------------------------------------------- 汇总与报告


def _md_escape(text: str) -> str:
    """对进入 Markdown 报告的动态文本做最小转义，防 HTML 标签/链接/代码/表格注入。
    渲染后仍显示原字符，纯文本查看时会看到转义序列。"""
    if not text:
        return text
    out = (text.replace("&", "&amp;")
               .replace("<", "&lt;")
               .replace(">", "&gt;"))
    out = out.replace("\\", "\\\\")
    for ch in ("[", "]", "`", "|", "*", "_", "(", ")"):
        out = out.replace(ch, "\\" + ch)
    return out


def summarize(entries, g_findings):
    total = len(entries)
    n_pass = sum(1 for e in entries if e["verdict"] == LEVEL_PASS)
    n_warn = sum(1 for e in entries if e["verdict"] == LEVEL_WARN)
    n_error = sum(1 for e in entries if e["verdict"] == LEVEL_ERROR)
    err_count = sum(1 for f in g_findings if f["level"] == LEVEL_ERROR)
    warn_count = sum(1 for f in g_findings if f["level"] == LEVEL_WARN)
    n_error += err_count
    n_warn += warn_count
    if err_count or any(e["verdict"] == LEVEL_ERROR for e in entries):
        verdict = "需修改"
    elif n_warn:
        verdict = "建议修改"
    else:
        verdict = "通过"
    return {
        "total": total, "pass": n_pass, "warn_entries": n_warn, "error_entries": n_error,
        "verdict": verdict,
    }


def markdown_report(meta, summary, entries, g_findings):
    L = []
    L.append("# 参考文献格式自检报告")
    L.append("")
    L.append("> - 生成时间：%s" % meta["time"])
    L.append("> - 输入来源：%s" % _md_escape(meta["source"]))
    L.append("> - 检查依据：GB/T 7714-2015《信息与文献 参考文献著录规则》（顺序编码制）")
    L.append("> - 工具：引文啄木鸟 reference-woodpecker（确定性规则引擎，纯本地检查）")
    L.append("> - 声明：本报告只覆盖可确定性判定的格式项，未覆盖项见报告末尾；不代改正文，修改决定权在作者。")
    L.append("")
    L.append("## 一、总体结论")
    L.append("")
    L.append("| 指标 | 数值 |")
    L.append("|---|---|")
    L.append("| 文献条目总数 | %d |" % summary["total"])
    L.append("| ✅ 通过 | %d |" % summary["pass"])
    L.append("| ⚠️ 含警告条目 | %d |" % summary["warn_entries"])
    L.append("| ❌ 含错误条目 | %d |" % summary["error_entries"])
    L.append("")
    L.append("**整体判定：%s**" % summary["verdict"])
    L.append("")

    code_count = {}
    for e in entries:
        for f in e["findings"]:
            code_count.setdefault((f["code"], f["message"][:40]), set()).add(e["no"])
    for f in g_findings:
        if f["level"] in (LEVEL_ERROR, LEVEL_WARN):
            code_count.setdefault((f["code"], f["message"][:40]), set()).add(0)

    L.append("## 二、问题汇总")
    L.append("")
    if code_count:
        L.append("| 级别 | 问题 | 涉及条目 |")
        L.append("|---|---|---|")
        for (code, msg), nos in sorted(code_count.items()):
            level = code.split("_")[0]
            level_show = {"E": "❌", "W": "⚠️"}.get(level, "ℹ️")
            nos_show = "全局" if nos == {0} else "、".join("[%d]" % n for n in sorted(nos))
            L.append("| %s | %s | %s |" % (level_show, _md_escape(msg), nos_show))
    else:
        L.append("未发现可确定性判定的问题。")
    L.append("")

    L.append("## 三、逐条明细")
    L.append("")
    for e in entries:
        icon = {"PASS": "✅", "WARN": "⚠️", "ERROR": "❌", "INFO": "ℹ️"}.get(e["verdict"], "·")
        body_text = re.sub(r"^\s*[\[［][^\]］]*[\]］]\s*", "", e["raw"])
        excerpt = body_text.replace("\n", " ")
        if len(excerpt) > 120:
            excerpt = excerpt[:120] + "…"
        L.append("### %s 条目[%s] %s" % (icon, e["no"], _md_escape(excerpt)))
        L.append("")
        if e["findings"]:
            for f in e["findings"]:
                L.append("- %s **%s**：%s" % (
                    {"ERROR": "❌", "WARN": "⚠️", "INFO": "ℹ️", "PASS": "✅"}.get(f["level"], "·"),
                    _md_escape(f["message"]), ("｜建议：" + f["advice"]) if f.get("advice") else ""))
        else:
            L.append("- ✅ 通过本引擎全部可确定性检查项")
        L.append("")

    if g_findings:
        L.append("## 四、全局检查结果（编号序列、正文引用对应等）")
        L.append("")
        for f in g_findings:
            L.append("- %s **%s**：%s%s" % (
                {"ERROR": "❌", "WARN": "⚠️", "INFO": "ℹ️"}.get(f["level"], "·"),
                _md_escape(f["message"]), ("｜建议：" + f["advice"]) if f.get("advice") else "", ""))
        L.append("")
        idx = "五"
    else:
        idx = "四"

    L.append("## %s、本引擎未覆盖、需人工确认的项" % idx)
    L.append("")
    L.append("- 作者是否满 3 人应加“等/et al.”（需知道真实作者数）")
    L.append("- 文献是否真实存在（本工具不联网核验）")
    L.append("- 题名内部标点是否与原文一致（题名照录）")
    L.append("- 页边距、字体字号等版式项（不属于著录规则）")
    L.append("- 著者-出版年制（无 [序号] 列表）不在检查范围")
    L.append("")
    return "\n".join(L)


# ---------------------------------------------------------------- 主流程

def run_check(text: str, source: str, refs_only: bool):
    body, refs, mode, notes = locate_references(text, refs_only)
    entries, g_findings = split_entries(refs)
    for n in notes:
        g_findings.append({"level": LEVEL_INFO, "code": "I_LOCATE", "message": n, "advice": ""})

    if mode == "none" or not entries:
        g_findings.append({
            "level": LEVEL_ERROR, "code": "E_NO_REFS",
            "message": "未检出任何参考文献条目",
            "advice": "确认文档含“参考文献”标题及 [1] 开头的编号条目；或用 --refs-only 直接传入文献列表；若原文采用著者-出版年制（无 [序号] 编号），本工具不适用，请告知用户",
        })
        entries = []
    else:
        if len(entries) > MAX_ENTRIES:
            g_findings.append({"level": LEVEL_INFO, "code": "I_TOO_MANY",
                               "message": "条目数 %d 超过单次检查上限 %d，仅检查前 %d 条" % (
                                   len(entries), MAX_ENTRIES, MAX_ENTRIES),
                               "advice": "分批检查"})
            entries = entries[:MAX_ENTRIES]
        for e in entries:
            check_entry(e)
        g_findings.extend(check_sequence(entries))
        g_findings.extend(check_duplicates(entries))
        cite_findings, order = check_citation(body, entries)
        g_findings.extend(cite_findings)

    meta = {"time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "source": source, "locate_mode": mode}
    summary = summarize(entries, g_findings)
    return meta, summary, entries, g_findings


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="GB/T 7714-2015 顺序编码制参考文献格式自检（引文啄木鸟）")
    ap.add_argument("input", help="输入文件路径（.txt/.md/.docx）或 - （stdin）")
    ap.add_argument("--refs-only", action="store_true", help="输入本身即为参考文献列表（无正文）")
    ap.add_argument("--report", help="Markdown 报告输出路径")
    ap.add_argument("--json", dest="json_path", help="JSON 结果输出路径")
    ap.add_argument("--lang", default="auto", help="预留参数（当前自动识别）")
    args = ap.parse_args(argv)

    text, source, err = load_input(args.input)
    if err:
        print("[输入错误] %s" % err, file=sys.stderr)
        return 2
    if not text.strip():
        print("[输入错误] 输入为空，没有可检查内容", file=sys.stderr)
        return 2

    meta, summary, entries, g_findings = run_check(text, source, args.refs_only)
    report = markdown_report(meta, summary, entries, g_findings)

    if args.json_path:
        payload = {
            "meta": meta, "summary": summary,
            "entries": [{"no": e["no"], "type": e.get("type"), "raw": e["raw"],
                         "verdict": e["verdict"], "findings": e["findings"]} for e in entries],
            "global_findings": g_findings,
        }
        Path(args.json_path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.report:
        Path(args.report).write_text(report, encoding="utf-8")

    print("=== 引文啄木鸟 · 参考文献格式自检 ===")
    print("输入来源: %s（定位方式: %s）" % (meta["source"], meta["locate_mode"]))
    print("条目总数: %d | 通过 %d | 含警告 %d | 含错误 %d" % (
        summary["total"], summary["pass"], summary["warn_entries"], summary["error_entries"]))
    print("整体判定: %s" % summary["verdict"])
    if args.report:
        print("报告已写入: %s" % args.report)
    if args.json_path:
        print("JSON 已写入: %s" % args.json_path)
    if not args.report and not args.json_path:
        print()
        print(report)
    return 0


if __name__ == "__main__":
    sys.exit(main())
