---
name: docx-layout-doctor
description: 体检并修正 DOCX 跨引擎版面漂移——auto 行距、多余分页符、写死的 NUMPAGES 缓存、WPS 保存后的页眉文字残留、IPA 字体缺失。当用户说 docx 在鸿蒙 WPS / Word 里页数对不上、分页位置变了、需要手工调行距或删分页符、或要在生成 docx 后加一道版面门禁时使用。Inspect and fix cross-engine DOCX layout drift: auto line spacing, extra page breaks, hard-coded NUMPAGES cache, WPS-save header residue, missing IPA font. Use when a docx renders with wrong page counts or breaks in HarmonyOS WPS/Word, when manual spacing/break fixes repeat, or as a post-generation layout gate.
agent_created: true
---

# DOCX 版面医生 / DOCX Layout Doctor

## 目的 Purpose

在生成后立刻检出并固化 DOCX 的漂移源，使同一份文件在鸿蒙 WPS、Word、腾讯文档里得到一致的分页，消除"每次都要手工调行距和分页符"的返工。

Detect and pin down layout drift sources right after generation, so one docx paginates consistently across HarmonyOS WPS, Word, and Tencent Docs — eliminating repeated manual spacing/break fixes.

## 核心原理 Core principle

DOCX 存的是**排版指令 + 缓存结果**，不是最终画面。`auto` 行距的倍数由各引擎按自身中文字体度量换算；同一份 docx 可能在 A 引擎是 4 页、B 引擎是 5 页。根治办法是把相对行距固化为 `exact` 绝对 twips，并把中间硬分页符减到最少。

A docx stores layout *instructions plus cached results*, not a final image. `auto` line spacing is converted per engine using its own CJK font metrics, so the same file may render 4 pages in one engine and 5 in another. The cure is to pin relative spacing to `exact` twips and minimize mid-document hard breaks.

**这不是 WPS 的 bug**。页脚里的 `PAGE`/`NUMPAGES` 是标准字段，但字段结果（`w:t` 里的"共 4 页"）是生成时写死的缓存值；引擎是否刷新字段，决定了你看到的是新值还是旧值。

This is not a WPS bug. Footer `PAGE`/`NUMPAGES` are standard fields, but their results are hard-coded cached values; whether an engine refreshes fields decides which number you see.

## 何时使用 When to use

- 用户抱怨"页数对不上 / 分页位置变了 / 行距在 WPS 里变了"。
- 生成 docx 后要做自动化版面门禁。
- 接入现有生成器，在写盘环节自动体检或修正。

## 工作流 Workflow

### 1. 体检 Inspect

```bash
python3 scripts/layout_doctor.py lint <docx> [<docx> ...]
python3 scripts/layout_doctor.py lint <docx> --json    # 机器可读 / machine-readable
```

退出码：有 WARN 时为 `1`，干净时为 `0`，可直接当门禁使用。

报告六类检查项 / Six checks:

| id | 级别 | 含义 Meaning |
|---|---|---|
| `auto_line` | WARN | 相对行距，跨引擎漂移头号来源 / relative spacing, top drift source |
| `multi_break` | WARN | 真分页符 > 1，中间硬分页 / more than one real break |
| `fake_break` | INFO | `w:val="0"` 是 WPS 保存副作用、非真分页 / WPS-save artifact, not a real break |
| `page_cache` | WARN | NUMPAGES 缓存值写死 / hard-coded cached total-page count |
| `header_residue` | WARN | 图片页眉 + 文字残留 / image header with stray text |
| `ipa_font` | WARN | IPA 字符未设 Noto Sans / IPA chars without Noto Sans |

### 2. 修正 Fix

```bash
python3 scripts/layout_doctor.py fix <in.docx> <out.docx>
```

默认行为：`auto → exact` 行距 + 真分页符只保留最后一个（通常是答案页前）。**不覆盖原文件**。

Defaults: convert `auto` to `exact`, keep only the last real page break. Never overwrites the input.

参数 / Options:
- `--line-mode exact|off` — 是否转 exact 行距（默认 exact）
- `--break-mode keep-last|keep-all|none` — 分页符策略（默认 keep-last）
- `--answer-line <twips>` — 保留该 auto 值作为答案页 exact 行距。**已有手工收紧答案页的卷子必须传**，否则 312 会被字号映射放宽到 336，导致答案区溢出。

示例 / Example: `fix paper.docx out.docx --answer-line 312`

修正器只改 `word/document.xml` 的 spacing / break 标签，不动文字内容。**输出后必须复跑 lint 确认内容零丢失**。

The fixer only touches spacing/break tags in `word/document.xml`; text is untouched. Re-run lint on the output to confirm zero content loss.

### 3. 验收 Verify（不可跳过 / mandatory）

页数**只以真实 Word / WPS 导出的 PDF 为准**。自研渲染器（如 `_docx2pdf.js`）行高算法不同、不读 sectPr 边距、不渲染 `w:drawing`，它的页数不可用于改版判断。

Trust only the page count from a real Word/WPS PDF export. A home-grown renderer uses different line-height math, ignores sectPr margins, and skips `w:drawing`; never use its page count to drive layout changes.

### 4. 接入生成器 Integrate

根治优于事后修正——优先在生成阶段就复用 `references/drift_sources.md` 里的 `EXACT_LINE_MAP`，直接产出 exact 行距。

Prevention beats repair: reuse `EXACT_LINE_MAP` from `references/drift_sources.md` at generation time and emit exact spacing directly.

需要事后门禁时，在生成器写盘后追加 / For a post-generation gate:

```python
import subprocess, sys
subprocess.run([sys.executable, "<skill>/scripts/layout_doctor.py",
                "lint", docx_path], check=False)  # WARN 时退出码 1
```

## 事实基准 Ground truth

所有数值来自用户现有生成器 `_quiz2docx.py:31`（`EXACT_LINE_MAP`）与 2026-11-25 真实产物验证。**勿自行发明映射值**。

All constants come from the user's existing generator `_quiz2docx.py:31` and were verified against a real docx on 2026-11-25. Do not invent mapping values.

详见 / See `references/drift_sources.md`。

## 资源 Resources

- `scripts/layout_doctor.py` — 零依赖体检 + 修正器，中英双语输出。
- `references/drift_sources.md` — 漂移源原理、事实数值、WPS 副作用、开源脱敏清单。
