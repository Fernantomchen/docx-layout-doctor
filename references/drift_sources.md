# 漂移源与事实基准 / Drift Sources and Ground Truth

> 本文件的数值全部来自用户现有生成器与真实产物验证，用于给出确定性的修正参数。
> All constants below come from the user's existing generator and verified real docx output.

## 1. 行距：头号漂移源 / Line spacing — primary drift source

### 事实基准（来自 `_quiz2docx.py:31`）

| half-point 字号 `w:sz` | exact 绝对行高 twips |
|---|---|
| 20 | 336 |
| 21 | 336 |
| 22 | 360 |
| 24 | 396 |
| 30 | 456 |
| 无显式 sz（继承默认） | 336 |

换算：`240 = 单倍`、`276 = 1.15 倍`、`300 = 1.25 倍`、`360 = 1.5 倍`（这些是 `auto` 的相对值）。

### 为什么 auto 会漂移 / Why `auto` drifts

用户代码原注释（ `_quiz2docx.py:27-30`）：

> auto 的「1.5倍」由各引擎按自身中文字体度量换算绝对高度，WPS/腾讯与本地渲染器
> 结果不同 → 同一 docx 一边4页一边5页；exact 给固定 twips，任何引擎一致。

即：`auto` 把"1.5 倍"留给用户端引擎解释，不同引擎的中文字体度量不同；`exact` 直接给定绝对 twips，绕开解释差异。

### 真实产物验证（2026-11-25）

修复前：`lineRule="auto"` 102 处（360 × 83，312 × 19），`exact` 0 处。
修复后：`auto` 0 处，文字内容逐字符相同（4896 → 4896）。

## 2. 分页符：真分页 vs WPS 假分页 / Real vs fake breaks

| 形态 XML | 性质 | 处理 |
|---|---|---|
| `<w:pageBreakBefore/>` | 真分页 | 计入；保留最后一个（答案页前） |
| `<w:pageBreakBefore w:val="1"/>` | 真分页 | 同上 |
| `<w:pageBreakBefore w:val="0"/>` | **假分页** | WPS 保存自动写出的"不分页"标记，忽略 |
| `<w:br w:type="page"/>` | 真分页（硬分页） | 计入；通常应删除中间的 |

**坑**：只数 `pageBreakBefore` 会把 19 个 `w:val="0"` 误判成真分页。体检器必须先按 `w:val` 分类。

**Why it matters**：中间硬分页把内容切成固定几块，行距一变就出现半页空白或溢出；省纸原则是"按张算不按页算"，宁可自然流动也不要硬切。

### 答案页行距必须单独保留 / Preserve answer-page spacing

用户会在 WPS 里手工把答案区行距从 360 收紧到 312（"答案区不需手写空间、收紧防溢出"）。
fix 默认按字号映射会把 312 一并转成 336，反而**放宽答案区**、有溢出风险。

因此修正已有手工调整的卷子时，必须传 `--answer-line 312`：原 auto 值等于 312 的段落
直接固化为 exact 312，不被字号映射覆盖。

2026-11-25 实证：现行定稿卷 102 处 auto（360×83 + 312×19），
用 `--answer-line 312` 修正后 → 336×74 + 312×19 + 360×2 + 396×6 + 456×1，
答案区 19 段保持 312，正文逐字相同。

## 3. 页脚页码缓存 / Footer page-number cache

真实 docx 的 `word/footer1.xml` 结构：

```xml
<w:instrText> PAGE </w:instrText>
<w:instrText> NUMPAGES </w:instrText>
<w:t xml:space="preserve">第 <w:t>1</w:t> 页 / 共 <w:t>4</w:t> 页 · …</w:t>
```

- 字段指令（`PAGE`/`NUMPAGES`）是标准的。
- 字段结果显示为 `w:t`，其中页数是**生成时写死的缓存值**。
- 引擎刷新字段 → 显示真实页数；不刷新 → 显示缓存页数。这就是"页码不一致"的传导链。

**结论**：脚本不该猜页数。页数以真实 Word/WPS 导出 PDF 为准；脚本只负责让版面稳定，从而让"猜"和"真实"尽量一致。

## 4. WPS 保存副作用 / WPS save side effects

1. **页眉残留**：图片页眉被 WPS 保存后，可能退化出文字页眉残留（`header*.xml` 同时含 `w:drawing` 与非空 `w:t`）。图片仍在、内容完整，属误报，但应记录。
2. **假分页标记**：每个段落被写入 `pageBreakBefore w:val="0"`（见第 2 节）。

## 5. IPA 字体 / IPA font

含 IPA 字符（`U+0250–02FF`、`æðθŋɜː` 等）的 run 必须显式设 `Noto Sans`。生成器不自动加字体，否则音标印成空白。体检器会把未设 Noto 的段落列为 WARN。

## 6. 本地渲染器不可作判据 / Never trust a home-grown renderer

来自用户长期实证（写入 MEMORY.md）：

- `_docx2pdf.js` 不渲染 `<w:drawing>`、不读 sectPr 边距、行高算法不同。
- 曾据其页数误判 3 页并把卷子改坏，已回滚。
- 真实 WPS 导出的 PDF 才是页数依据；本地渲染显示的"第 3 页偏空"是假象。

**规则 / Rule**：体检器只查结构风险，不承诺页数；页数一律由真实引擎导出后目视确认。

## 7. 开源脱敏清单 / Open-source sanitization checklist

发布到 GitHub 前，逐项确认（Pre-release checklist）：

- [ ] 学生真实姓名（文件名、页眉页脚、报告正文、示例内容）→ 替换为 `StudentA` / 占位符
- [ ] 机构名称 → 泛化为"某培训机构"
- [ ] 电话号码 → 删除或替换为 `000-0000-0000`
- [ ] 资料库节点 ID / 云文档链接 → 删除
- [ ] 家长对话截图、微信截图 → 不上传
- [ ] 学生分数、学情报告 → 不上传（或全部伪造为示例数据）
- [ ] 数据集里的真实答卷照片 → 不上传
- [ ] 保留 MIT 协议与来源标注（作者要求：自由引用/二次使用，唯一要求标注来源）

本 skill 的脚本与文档已按此清单脱敏，不含任何真实学员信息。
