# 真实学校模板工作流（复杂论文 / 复杂模板）

本文件沉淀「真实论文 + 真实学校模板」场景的完整方法。当你拿到的论文**已含文字、表格、图片、代码**，或学校模板是**多分节、自定义样式、奇偶页眉、罗马/阿拉伯页码混用**时，必须用这里的路径，而不是 SKILL.md 里那条「克隆模板替换正文」的简单路径。

## 一、两条路径的选择

| 判断依据 | 路径 A：克隆法 | 路径 B：移植法 |
| --- | --- | --- |
| 论文形态 | 纯文本 / 简单初稿，无图表 | 已完成的长文，含表格/图片/代码/脚注 |
| 模板形态 | 单节、标准样式（Normal/Heading1…） | 多节、自定义样式（CMxxx 等）、奇偶页眉 |
| 核心脚本 | `build_document.py` | `transform_paper.py`（+ `rebuild_cover.py`） |
| 机制 | 从模板整包克隆，只换正文 | 以论文为基线，把模板格式移植进来 |

**红线**：真实论文含表格/图片时**绝不能走克隆法**——重建正文会丢图丢表。凡是有 `word/media/`、`w:tbl`、`w:drawing` 的论文，一律走移植法。

## 二、先诊断，再动手（必做）

在写任何转换代码前，先用只读脚本把模板和论文摸清楚。 Step 0 的诊断决定后面所有配置。

### 2.1 诊断模板

```python
# 用 skill 自带解析器出规格 + 人类可读摘要
python scripts/inspect_template.py <模板.docx> -o spec.json --summary summary.md
```

然后补充确认这几项（inspect 不会自动给全）：

- **分节结构**：`word/document.xml` 里所有 `w:sectPr`（含段落 `pPr` 内的），逐节看 `pgSz/pgMar/pgNumType/headerReference/footerReference/titlePg`。真实学校模板常有 8~12 节：封面 / 摘要 / Abstract / 目录 / 各章 / 参考文献 / 致谢。
- **奇偶页眉**：检查 `word/settings.xml` 是否有 `<w:evenAndOddHeaders/>`。有则奇数页眉=章节名、偶数页眉=论文题目（学校常见规范）。
- **样式体系**：列出 `word/styles.xml` 里的 styleId。若不是标准名而是 `CMheading1`、`CMcontent`、`CMreflist` 之类自定义名，记录它们对应的字体字号（正文/一级/二级/三级标题/参考文献/图表题注）。
- **页码规则**：摘要/目录用罗马数字（upperRoman，start=1），正文从第 1 章起阿拉伯数字（decimal，start=1）。
- **页眉页脚部件**：枚举 `word/header*.xml`、`word/footer*.xml`，看每节的 rId 映射到哪个文件，文件里的 `w:t` 文字和 PAGE/NUMPAGES 域。

### 2.2 诊断论文

- **段落结构**：正文 body 直属 `w:p` 的数量；各标题段落用的样式 ID（Heading1/2/3？还是中文数字「第一章」写在正文里？）。
- **标题编号**：章标题是「第一章」还是「第1章」？二级是「1.1」还是「一、」？必须与模板统一（见第四节）。
- **内容清单（指纹）**：按文档顺序拼接所有 `w:t` 做 SHA256；记录表格数、`w:drawing` 数、`word/media/` 各图片的 md5。这是后面「零丢失」验证的基线。
- **直接格式覆盖**：run 上是否带显式 `rFonts/sz/b`（会覆盖样式）。真实论文常有几百处，必须在移植时清掉。
- **分节与封面**：论文是否已有封面？是否单节？封面后是否另起页？

## 三、移植法五步（transform_paper.py）

配置驱动的脚本 `scripts/transform_paper.py` 已把下面五步固化为可复用流程。先 `python scripts/transform_paper.py --print-template` 看配置模板，再按诊断结果填写。

1. **样式移植**：以论文 styles.xml 为基线，用模板的 `docDefaults` 整体替换，再按需导入模板自定义样式（`styles.import_ids`）。这样正文继承模板默认字号，标题用模板样式。
2. **段落样式映射 + 清直接格式**：`role_rules` 是有序规则（先匹配先应用），每条支持 `when`（index 区间 / style / regex / equals / maxlen，多条件 AND）和 `style`。命中后 `set_style` 并 `strip_direct`。**标题类务必 `strip_bold: true`**——原论文的 run 常残留 `bold=True`，不清掉会盖过模板的「不加粗」。
3. **标题编号规范化**：`renumber.cn_to_arabic=true` 把「第一章」改成「第1章」。目录条目（`toc_index_range`）只改章号，页码域节点必须保留。
4. **分节重建**：`sections` 按顺序列出每节，用 `end`（本节最后段落索引，None=body 末尾）切分，配置 header/footer 引用与 pgNumType。脚本自动为每个 header/footer 部件分配不冲突的新 rId。
5. **页眉页脚移植 + 关系修复**：复制模板 header/footer 部件，`title.old→new` 替换示例题目，`header_text` 覆盖奇数页眉为章节名；同步追加 rels、Content_Types、`evenAndOddHeaders`。

脚本输出后**必须**接着跑第 5 节的验证三件套。

## 四、标题编号规范（最常见返工点）

真实学校模板对标题编号有明确写法，务必先看清再转换：

- 一级标题：模板可能写「第1章 绪论.」（阿拉伯数字 + 空格 + 末尾句点），论文常是「第一章 绪论」（中文数字、无点）。用 `renumber` 转换。
- 注意句点：若模板只有个别章带句点（多为模板笔误），不要强行统一加，用 `add_period: false` 保持规范一致即可；用户要求严格照抄时再开。
- 二级/三级：多为「1.1」「1.1.1」，若论文已是此格式则不要动。
- **目录必须同步**：目录里章号要和正文一致，且只改章号文字，页码与制表位引导符不能动。

## 五、验证三件套（缺一不可）

### 5.1 内容零丢失

```bash
python scripts/audit_content.py <原论文.docx> <排版后.docx> --normalize-headings
```

检查：文字 SHA256（归一化章号后）、表格数、`w:drawing` 数、图片字节 md5 列表、ZIP 完好、所有 XML 合法。全 PASS 才算「只改格式没动内容」。若文字 FAIL 但图片/表格 PASS，多半是标题编号规范化——用 `--normalize-headings` 复核。

### 5.2 格式达标

```bash
python scripts/inspect_template.py <模板.docx> -o spec.json
python scripts/verify_requirements.py <排版后.docx> --spec spec.json --report report.md
```

要求 FAIL=0。**注意**：验收器是通用实现，只认标准样式名、且把模板示例文字当硬性项。对自定义样式模板（CM 系列）和「示例题目→真实题目」的替换，会产生两类误报，需人工判读并在验收报告里说明：

- 示例题目/院系等文字差异：属于正确替换，不是缺陷。
- 自定义样式：验收器可能按标准样式检查而漏项，此时手工核对 styles.xml 里 CM 样式的字体字号（可写进交付报告）。

### 5.3 结构目视

确认：封面节无页眉页脚、页码从摘要/正文正确起算、奇偶页眉正确、目录域需在 Word/WPS 中按 F9 更新。

## 六、封面重建（rebuild_cover.py，路径B 默认必做）

**零配置命令（推荐，换 AI/换模板都稳）：**

```bash
python scripts/rebuild_cover.py <第3步无封面成品.docx> <学校模板.docx> <带封面成品.docx>
```

脚本自动完成，**不需要用户填任何配置、不需要提供个人信息**：

1. **自动识别封面范围**：从模板 body 第一个 `w:p` 到第一个带 `sectPr` 的段落（含），兜底用"完成日期/答辩日期"所在段。
2. **自动取论文题目**：src 第一个非空段落。
3. **只替换题目**：封面「题目：」栏的示例值换成论文题目；其他段落一律不动结构。
4. **个人信息一律清空**：姓名/学号/导师/专业/学院/日期等命中标签后只保留"标签："与下划线，值清空——**属隐私，不强制客户提供，由使用者在 Word 里自填**。
5. **自动搬迁校徽**：从模板 rels 找到封面首图，分配不冲突的新 rId 与独立 media 文件（`school_logo.png`）。
6. **自动读封面页边距**：从模板封面节 `sectPr` 的 `pgSz/pgMar/cols/docGrid` 读取。
7. 原文标题段删除；紧随标题的短声明段移到封面末尾保留（内容零丢失）。

**可选覆盖**（自动识别不适用时）：`python scripts/rebuild_cover.py <config.json>`，骨架用 `--print-template`。可覆盖 `new_title` / `cover_range` / `title.old` / `fields_to_blank` / `logo` / `drop_paragraphs` / `move_to_cover` / `cover_section`。

**验收判定**：第一页应有校徽 + "毕业论文" + 新题目 + 空的下划线字段。若仍显示模板示例题目或示例姓名（如"快论文""20135091612"），说明漏跑本步骤或用错了旧版脚本，必须重做。

重建后必须重跑 5.1 与 5.2。注意 `audit_content.py` 会因"封面新增校徽+字段文字"误报 FAIL，用精确复核（排除封面段逐段+逐字节比对）确认正文零丢失即可。

## 七、坑位清单（血泪经验）

1. **直接格式覆盖样式**：论文 run 上的 `sz/rFonts/b` 会盖过模板样式。标题的 `b` 尤易漏——`strip_direct` 要用 `strip_bold=true`。
2. **缺 `pPrDefault`**：`docDefaults` 可能没有 `pPrDefault`，解析要容错（见 docx_utils.py）。
3. **关系文件属性无命名空间**：`document.xml.rels` 里是裸 `Id/Type/Target`，读取时别套 `w:` 前缀。
4. **页眉文字分散在多个 `w:t`**：替换章节名时要把第一个 `w:t` 设为新文字、其余清空，别只改一段。
5. **页脚 PAGE/NUMPAGES 必须是域**：不能写成静态数字。复制模板页脚原样即可。
6. **分节符写在段落 `pPr` 内**：不在 body 直属。重建时要清掉段落里旧的 `sectPr` 再挂新的。
7. **封面 sectPr 冲突**：模板封面段自带 sectPr，搬过来必须清掉，只在封面最后一段挂一个；否则封面被拆成多页。
8. **ElementTree 序列化丢前缀**：处理复杂命名空间文档时，先 `register_namespace` 全部前缀（transform/rebuild 脚本已内置）。
9. **`%` 格式化占位符混用**：构造 sectPr 字符串时，`%(w)d` 与位置参数不能混，统一用字典格式化。
10. **验收 FAIL 不一定是真缺陷**：区分「格式缺陷」与「示例文字替换/自定义样式误报」，在报告里说清楚。

## 八、交付口径

交付时给出：成品 .docx、验收报告，并说明三件事——(1) 走的是移植法/克隆法；(2) 验收 FAIL 数及误报说明；(3) 需用户在 Word/WPS 手动做的事（更新目录域 F9、填写封面个人信息、删除原文自带的声明/占位段）。
