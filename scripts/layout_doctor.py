#!/usr/bin/env python3
"""docx-layout-doctor — DOCX 跨引擎版面漂移体检与修正（纯标准库，零依赖）。

事实基准取自用户现有生成器 _quiz2docx.py:31，勿凭空改动数值。
核心原理：docx 存的是「排版指令 + 缓存结果」，不是最终画面。
auto 行距与多余分页符会在不同引擎（鸿蒙 WPS / Word / 本地渲染器）被重新解释，
导致页数漂移；本工具把漂移源在生成后立刻检出并固化为 exact 值。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

# ── 事实基准（_quiz2docx.py:31，用户实测）──────────────────────
# half-point 字号 → exact 绝对行高 twips
EXACT_LINE_MAP = {20: 336, 21: 336, 22: 360, 24: 396, 30: 456}
DEFAULT_LINE = 336

#  IPA / 音标字符（触发 Noto Sans 要求）
IPA_RE = re.compile(r"[\u0250-\u02ffæðθŋɜː]")

# ── XML 片段正则（保留原始字节，不做序列化）────────────────────
RE_PARA = re.compile(r"<w:p[ >].*?</w:p>", re.S)
RE_SPACING = re.compile(r"<w:spacing\b[^>]*/>")
RE_LINE = re.compile(r'w:line="(\d+)"')
RE_SZ = re.compile(r'<w:sz w:val="(\d+)"')
RE_T = re.compile(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.S)
RE_PBB = re.compile(r"<w:pageBreakBefore\b([^>]*)/>")
RE_HARD_BR = re.compile(r'<w:br\b[^>]*w:type="page"[^>]*/>')
RE_FIELD = re.compile(r"<w:instrText[^>]*>\s*(PAGE|NUMPAGES)\s*</w:instrText>")
RE_DRAWING = re.compile(r"<w:drawing\b")
RE_RFONTS = re.compile(r"<w:rFonts\b[^>]*>")


def _is_fake_break(attrs: str) -> bool:
    """w:val=0/false/off 是 WPS 保存写的『不分页』标记，非真分页。"""
    v = re.search(r'w:val="([^"]+)"', attrs)
    return bool(v and v.group(1).lower() in ("0", "false", "off", "no"))


def read_docx(path: str) -> dict:
    """读取 docx 内 word/ 下所有 XML 文本。"""
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if n.startswith("word/") and n.endswith(".xml")]
        return {n: z.read(n).decode("utf-8", errors="replace") for n in names}


def para_text(p: str, limit: int = 80) -> str:
    return "".join(RE_T.findall(p))[:limit]


def inspect(path: str) -> dict:
    parts = read_docx(path)
    doc = parts.get("word/document.xml", "")
    issues = []

    # ① 行距：auto 是跨引擎漂移的头号来源
    auto_spans = [m.group(0) for m in RE_SPACING.finditer(doc) if 'w:lineRule="auto"' in m.group(0)]
    exact_count = len([m for m in RE_SPACING.finditer(doc) if 'w:lineRule="exact"' in m.group(0)])
    if auto_spans:
        dist = Counter(
            (RE_LINE.search(s).group(1) if RE_LINE.search(s) else "?") for s in auto_spans
        )
        sample = []
        for m in RE_PARA.finditer(doc):
            if 'w:lineRule="auto"' in m.group(0):
                sample.append(para_text(m.group(0)))
            if len(sample) >= 5:
                break
        issues.append({
            "id": "auto_line", "level": "WARN",
            "msg": (f"auto 行距 {len(auto_spans)} 处（值分布 {dict(dist)}），exact {exact_count} 处 | "
                    f"auto line spacing: {len(auto_spans)} (values {dict(dist)}), exact {exact_count}"),
            "why": ("auto 倍数由各引擎按自身字体度量换算，WPS/腾讯/本地渲染器结果不同 → 同一 docx 一边 4 页一边 5 页 | "
                    "auto multiples are converted per-engine, so the same docx may render 4 or 5 pages"),
            "fix": "改 exact 绝对 twips / convert to exact twips",
            "sample": sample,
        })

    # ② 分页符：区分真分页与 WPS 假分页
    pbb = list(RE_PBB.finditer(doc))
    real_pbb = [m for m in pbb if not _is_fake_break(m.group(1))]
    fake_pbb = [m for m in pbb if _is_fake_break(m.group(1))]
    hard = RE_HARD_BR.findall(doc)
    real_total = len(real_pbb) + len(hard)
    if real_total > 1:
        issues.append({
            "id": "multi_break", "level": "WARN",
            "msg": (f"真分页符 {real_total} 个（pageBreakBefore {len(real_pbb)} + w:br {len(hard)}）| "
                    f"{real_total} real page breaks"),
            "why": ("中间硬分页会把内容切成固定的几块，跨引擎行距一变就出现半页空白或溢出 | "
                    "mid-document hard breaks split content into fixed blocks; any line-height shift shows blank/overflow"),
            "fix": "只保留答案页前 1 个 / keep only the one before the answer page",
        })
    elif real_total == 1:
        issues.append({
            "id": "break_ok", "level": "INFO",
            "msg": f"真分页 1 个（正常）| 1 real break (normal); 假分页 {len(fake_pbb)} 个",
        })
    if fake_pbb:
        issues.append({
            "id": "fake_break", "level": "INFO",
            "msg": (f'pageBreakBefore w:val="0" 共 {len(fake_pbb)} 个 = WPS 保存副作用、非真分页 | '
                    f"{len(fake_pbb)} fake breaks (w:val=0) from WPS save, not real breaks"),
        })

    # ③ 页脚页码：标准字段 + 写死缓存值
    for fn, xml in parts.items():
        if "footer" not in fn:
            continue
        fields = RE_FIELD.findall(xml)
        vals = RE_T.findall(xml)
        if fields:
            issues.append({
                "id": "footer_field", "level": "INFO",
                "msg": f"{fn}: 字段 {fields}，缓存文本 {vals} | fields {fields}, cached text {vals}",
            })
            if "NUMPAGES" in fields:
                issues.append({
                    "id": "page_cache", "level": "WARN",
                    "msg": ("NUMPAGES 缓存值写死 → 引擎不刷新字段时显示生成时页数 | "
                            "cached NUMPAGES value is hard-coded; engines that skip field refresh show the stale count"),
                    "fix": ("页数以真实 WPS/Word 导出 PDF 为准，本地渲染器只查结构 | "
                            "trust the real WPS/Word PDF export for page count; local renderer is structure-only"),
                })

    # ④ 页眉图片 + 文字残留（WPS 保存副作用）
    for fn, xml in parts.items():
        if "header" not in fn:
            continue
        texts = [t.strip() for t in RE_T.findall(xml) if t.strip()]
        if RE_DRAWING.search(xml) and texts:
            issues.append({
                "id": "header_residue", "level": "WARN",
                "msg": (f"{fn}: 图片页眉 + 文字残留（WPS 保存副作用）| "
                        f"image header with stray text after WPS save"),
                "sample": texts[:3],
            })

    # ⑤ IPA 字符未设 Noto Sans
    ipa_bad = []
    for m in RE_PARA.finditer(doc):
        p = m.group(0)
        if IPA_RE.search(p) and "Noto" not in p:
            ipa_bad.append(para_text(p))
    if ipa_bad:
        issues.append({
            "id": "ipa_font", "level": "WARN",
            "msg": (f"含 IPA 字符但未设 Noto Sans 的段落 {len(ipa_bad)} 处 | "
                    f"{len(ipa_bad)} paragraphs contain IPA without Noto Sans"),
            "sample": ipa_bad[:5],
        })

    return {"file": str(path), "issues": issues}


def _fix_para_exact(p: str, answer_line: int | None = None) -> str:
    """把单个 <w:p> 里的 auto 行距换成 exact。

    默认按字号分级映射；answer_line 指定时，原 auto 值等于该值的段落
    直接固化为该 exact 值——用于保留用户手工收紧的答案页行距，
    避免 fix 把 312 重映射成 336 而放宽答案区、挤溢出。
    """
    szs = [int(x) for x in RE_SZ.findall(p)]
    exact = EXACT_LINE_MAP.get(max(szs) if szs else 21, DEFAULT_LINE)

    def sub_spacing(m):
        s = m.group(0)
        if 'w:lineRule="auto"' not in s:
            return s
        lv = RE_LINE.search(s)
        target = answer_line if (answer_line is not None and lv and int(lv.group(1)) == answer_line) else exact
        s = re.sub(r'w:line="\d+"', 'w:line="%d"' % target, s)
        return s.replace('w:lineRule="auto"', 'w:lineRule="exact"')

    return RE_SPACING.sub(sub_spacing, p)


def fix(path: str, out: str, line_mode: str = "exact", break_mode: str = "keep-last",
        answer_line: int | None = None) -> list:
    """生成修正版 docx（不覆盖原文件），返回改动清单。"""
    src, dst = Path(path).resolve(), Path(out).resolve()
    if src == dst:
        raise SystemExit("拒绝原地覆盖 / refuse in-place overwrite: output must differ from input")
    if not src.exists():
        raise SystemExit("输入文件不存在 / input not found: %s" % path)
    dst.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(src) as z:
        infos = z.infolist()
        data = {i.filename: z.read(i.filename) for i in infos}

    doc = data["word/document.xml"].decode("utf-8", errors="replace")
    changes = []

    if line_mode == "exact":
        before = doc.count('w:lineRule="auto"')
        doc = RE_PARA.sub(lambda m: _fix_para_exact(m.group(0), answer_line), doc)
        after = doc.count('w:lineRule="auto"')
        changes.append("auto→exact 行距: %d→%d | auto→exact line spacing: %d→%d" % (before, after, before, after))
        if answer_line is not None:
            changes.append("答案页行距保留 / answer spacing preserved: auto %d → exact %d" % (answer_line, answer_line))

    if break_mode == "keep-last":
        real = [(m.start(), m.end()) for m in RE_PBB.finditer(doc) if not _is_fake_break(m.group(1))]
        real += [(m.start(), m.end()) for m in RE_HARD_BR.finditer(doc)]
        real.sort()
        if len(real) > 1:
            for s, e in reversed(real[:-1]):
                doc = doc[:s] + doc[e:]
            changes.append("真分页符 %d→1（保留最后 1 个）| real breaks %d→1 (kept last)" % (len(real), len(real)))
    elif break_mode == "none":
        n = len([m for m in RE_PBB.finditer(doc) if not _is_fake_break(m.group(1))]) + len(RE_HARD_BR.findall(doc))
        doc = RE_HARD_BR.sub("", doc)
        doc = RE_PBB.sub(lambda m: "" if not _is_fake_break(m.group(1)) else m.group(0), doc)
        changes.append("删除全部真分页符 %d 个 | removed all %d real breaks" % (n, n))

    data["word/document.xml"] = doc.encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i, data[i.filename])
    return changes


def print_report(report: dict) -> None:
    print("=" * 72)
    print("DOCX 版面体检报告 / Layout Doctor Report")
    print("文件 / File: %s" % report["file"])
    print("=" * 72)
    if not report["issues"]:
        print("未检出漂移风险 / No drift risks detected.")
        return
    for it in report["issues"]:
        print("[%s] %s" % (it["level"], it["msg"]))
        if it.get("why"):
            print("    原因 / Why: %s" % it["why"])
        if it.get("fix"):
            print("    建议 / Fix: %s" % it["fix"])
        if it.get("sample"):
            print("    样例 / Sample: %s" % " | ".join(it["sample"]))
    print("-" * 72)
    warns = [i for i in report["issues"] if i["level"] == "WARN"]
    print("合计 / TOTAL: %d 项，其中警告 %d 项 | %d issues, %d warnings" % (
        len(report["issues"]), len(warns), len(report["issues"]), len(warns)))


def main() -> None:
    ap = argparse.ArgumentParser(description="DOCX 跨引擎版面漂移体检与修正（中英双语输出）")
    sub = ap.add_subparsers(dest="cmd")

    p_lint = sub.add_parser("lint", help="体检 / inspect")
    p_lint.add_argument("inputs", nargs="+", help="一个或多个 docx 路径")
    p_lint.add_argument("--json", action="store_true", help="输出 JSON / machine-readable output")

    p_fix = sub.add_parser("fix", help="修正 / fix")
    p_fix.add_argument("input")
    p_fix.add_argument("output")
    p_fix.add_argument("--line-mode", choices=["exact", "off"], default="exact")
    p_fix.add_argument("--break-mode", choices=["keep-last", "keep-all", "none"], default="keep-last")
    p_fix.add_argument("--answer-line", type=int, default=None,
                       help="保留该 auto 值作为答案页 exact 行距，避免被字号映射放宽 / preserve this auto value as exact for the answer page")

    args = ap.parse_args()
    if not args.cmd:
        ap.print_help()
        return

    if args.cmd == "lint":
        reports = [inspect(p) for p in args.inputs]
        if args.json:
            print(json.dumps(reports, ensure_ascii=False, indent=2))
        else:
            for r in reports:
                print_report(r)
        warns = sum(1 for r in reports for i in r["issues"] if i["level"] == "WARN")
        sys.exit(1 if warns else 0)

    if args.cmd == "fix":
        changes = fix(args.input, args.output, args.line_mode, args.break_mode, args.answer_line)
        print("修正完成 / Fix complete: %s" % args.output)
        for c in changes:
            print("  - %s" % c)
        if not changes:
            print("  无需改动 / No changes applied.")


if __name__ == "__main__":
    main()
