# docx-layout-doctor / DOCX 版面医生

体检并修正 DOCX 的跨引擎版面漂移，让同一份文件在鸿蒙 WPS、Word、腾讯文档里得到一致的分页。

Inspect and fix cross-engine DOCX layout drift so one file paginates consistently across HarmonyOS WPS, Word, and Tencent Docs.

## 为什么做 / Why

DOCX 存的是"排版指令 + 缓存结果"，不是最终画面。`auto` 行距的倍数由各引擎按自身中文字体度量换算，同一份文件可能在 A 引擎 4 页、B 引擎 5 页。这常被误判为某个 WPS 的 bug，其实是 DOCX 的设计使然。

根治办法：把相对行距固化为 `exact` 绝对 twips，把中间硬分页符减到最少，页数以真实引擎导出的 PDF 为准。

## 用法 / Usage

零依赖，只需要 Python 3。

```bash
# 体检（有 WARN 时退出码 1，可作门禁）
python3 scripts/layout_doctor.py lint paper.docx
python3 scripts/layout_doctor.py lint paper.docx --json

# 修正（不覆盖原文件）
python3 scripts/layout_doctor.py fix paper.docx paper.fixed.docx

# 分页符策略 / break policy
python3 scripts/layout_doctor.py fix paper.docx out.docx --break-mode none
```

修正后请复跑 `lint`，并用真实 Word/WPS 导出 PDF 确认页数——脚本不承诺页数。

## 检查项 / Checks

- `auto_line` — 相对行距（漂移主因）
- `multi_break` — 真分页符 > 1
- `fake_break` — `w:val="0"` 假分页（WPS 保存副作用）
- `page_cache` — NUMPAGES 缓存值写死
- `header_residue` — 图片页眉 + 文字残留
- `ipa_font` — IPA 字符未设 Noto Sans

## 隐私 / Privacy

本仓库已脱敏：不含任何真实学生姓名、机构名、电话号码、云文档 ID 或答卷照片。事实数值来自公开可复现的生成器常量。

This repository is sanitized: no real student names, institution names, phone numbers, document IDs, or answer-sheet photos.

## 协议 / License

MIT — 可自由引用与二次使用，唯一要求标注来源。
