# Paper Requirement Matrix — AP-178《CRF Translation and aCRF Generation》

> 来源论文：PharmaSUG China 2023 - Paper AP-178，Fei Chen, Jie Wei, Chengwei Li, Kewei Yuan（Eli Lilly and Company）
> 源文件：`01Document/Pharmasug-China-2023-AP178.pdf`

## 1. 论文元信息与背景

| 项目 | 内容 |
|---|---|
| 论文标题 | CRF Translation and aCRF Generation |
| 作者 / 机构 | Fei Chen, Jie Wei, Chengwei Li, Kewei Yuan / Eli Lilly and Company |
| 会议 | PharmaSUG China 2023（Paper AP-178） |
| EDC 系统 | Veeva Vault EDC（CRF 为 2–3 列结构化表单、分区混合背景色） |
| 法规驱动 | CDE 2020《药物临床试验数据递交指导原则》：递交 NMPA 的外文 CRF 须译为中文；SDTM 批注已成为申报常规 |
| 参考论文 | PharmaSUG China 2022《CRF Chinese Translation – Mask Plan》（Chunpeng Zhao 等，基于 RAVE EDC）；本论文沿用其 mask plan 思路并扩展到 Vault 系统 |
| 技术栈 | Python + SAS（结论指出也可由单一语言完整实现） |

**核心方法论（论文摘要 4 步）：**

1. 首个研究在 SDS 文件中手工填充 SDTM 批注，后续研究基于已保存 SDS 迭代复用；
2. 精确计算空白 CRF（BLANKCRF）每个条目的位置；
3. 将 SRCCRF 与 SDS 按 external ID 合并，捕获 SDTM 批注；
4. 将 BLANKCRF 与 SRCCRF 按 form name、section head、item 合并，把 SDTM 批注落到空白 CRF 各条目下方。

翻译与批注最终均生成 `.xfdf` 文件，导入空白 CRF 完成主体流程。

---

## 2. 需求矩阵总览

优先级：P0 = 核心主线（不可缺失）；P1 = 迭代复用 / 流程兜底。

| REQ-ID | 模块 | 需求描述 | 输入 | 输出 | 技术与工具 | 优先级 | 论文出处 |
|---|---|---|---|---|---|---|---|
| PDF-01 | A. PDF 文本与坐标抽取 | 从 SRCCRF（Vault 导出 PDF）抽取全部字符（字母、标点、数字）及其位置坐标 | SRCCRF PDF | 含字符与坐标的 txt 文件 | SAS GROOVY 调用 Apache PDFBox | P0 | p3 Display 2/3 |
| PDF-02 | A. PDF 文本与坐标抽取 | 每个字符以左上角为起点记录坐标（矩形坐标 x_s, x_e, y_s, y_e） | SRCCRF PDF | 字符坐标表 | PDFBox / Python | P0 | p5 Display 8 |
| ASM-01 | B. 字符组装 | 词识别：以空格为分隔符；相邻字符间无空格则持续拼接，直至遇空格 | 字符级 txt | 词（word）序列 | Python | P0 | p5 |
| ASM-02 | B. 字符组装 | 行识别：相邻字符横向距离超过阈值 → 划分不同 range（分词边界）；字符高度差 → 划分不同行（列内） | 词序列 + 坐标 | 分行短语（phrase） | Python | P0 | p5 Display 7 |
| ASM-03 | B. 字符组装 | 段落结束识别：行间高度差显著大于阈值（如句号与下一句首字符的高度差）→ 段落结束 | 行级短语 | 句子/段落（sentence/paragraph） | Python | P0 | p5 Display 7 |
| ASM-04 | B. 字符组装 | 输出句子时保留原始矩形坐标 (x_s,x_e,y_s,y_e)，并生成调整坐标 (xs,xe,ys,ye)：将终点从原始位置拉伸至单元格右下角，防止英文原文露出 | 句子 + 原始坐标 | 句子 + 调整坐标 | Python | P0 | p5 Display 5 |
| ASM-05 | B. 字符组装 | 组装结果（句子、段落、坐标）保存至翻译元数据 Excel 供中文翻译使用 | 组装结果 | 翻译元数据 Excel（translated metadata） | Python / Excel | P0 | p3 Display 2 |
| TRN-01 | C. 中文翻译 | 采用 mask plan 方法：中文无缝替换原文（沿用 PharmaSUG 2022 方法，扩展到 Vault） | 翻译元数据 Excel | 中文翻译结果 | 人工翻译 + 程序遮罩 | P0 | p2 Overview |
| TRN-02 | C. 中文翻译 | 逐 entry item 完成中文翻译，形成含 external ID 的 CRF CN metadata Excel | 翻译元数据 Excel | CRF CN metadata Excel（条目 + 中文 + external ID） | 人工 + Python | P0 | p3 Display 2；p8 |
| BGC-01 | D. 背景色捕获 | 从 BLANKCRF 与 SRCCRF 抽取背景色元数据：实际页码（Variable Page）、分区矩形坐标（left x, left y, right x, right y）、RGB 值、列分割参考线 x_2（列1/列2 间）、x_3（列2/列3 间） | BLANKCRF / SRCCRF PDF | 背景色元数据文件 | Python | P0 | p6 Display 6 |
| BGC-02 | D. 背景色捕获 | BLANKCRF 背景色用于遮罩：新绘制背景色与原色完全一致（视觉无损） | BLANKCRF 背景色元数据 | 遮罩用背景色文件 | Python | P0 | p6 |
| BGC-03 | D. 背景色捕获 | SRCCRF 背景色用于建立列1–列3 一对一链接，实现 SDTM 批注位置与 SRCCRF 及外部 SDTM 元数据的对齐 | SRCCRF 背景色元数据 | 对齐用背景色文件 | Python | P0 | p6；验证见 Display 9 |
| SDS-01 | E. SDS 提取与 DED | 从 Veeva Vault EDC 提取 SDS 及全部 CRF 条目属性：form name/label、items group name/label、item name/label、external ID、data type、codelist 等 | Vault EDC 导出的 SDS | 结构化 SDS | Vault EDC 导出 | P0 | p7 |
| SDS-02 | E. SDS 提取与 DED | 因 item 可在同一 form 内重复出现，须以唯一键 external ID 建立 SDS ↔ CRF 映射 | SDS + CRF | SDS-CRF 映射 | SAS / Python | P0 | p7 |
| DED-01 | E. SDS 提取与 DED | DED（每 item 的 SDTM 变量批注）仅支持 2 组批注：SDTM Variables Name / SDTM SuppQual Desc、Second SDTM Variables Name / Second SDTM SuppQual Desc | SDS + 批注 | DED 批注数据 | 手工 / SAS | P0 | p7 |
| DED-02 | E. SDS 提取与 DED | 首个研究手工在 SDS 中填充 SDTM 批注；后续研究基于已保存 SDS 迭代复用 | 首研 SDS（含批注） | 后续研究 SDS 迭代版 | 手工 + 程序 | P1 | p1 摘要步骤1；p9 |
| POS-01 | F. 批注位置映射 | 批注放置规则：位于每 item 正下方、与 item 起点左对齐 | 中文 CRF + DED | 批注定位结果 | SAS / Python | P0 | p8 Display 11 |
| POS-02 | F. 批注位置映射 | CRF CN metadata（Excel）按 form name + external ID 与 SRCCRF 链接，捕获 DED | CRF CN metadata + SRCCRF | 带 DED 的链接数据 | SAS | P0 | p8 |
| POS-03 | F. 批注位置映射 | CRF CN metadata 中每个 entry item 及其 external ID 通过背景色坐标建立链接 | CRF CN metadata + 背景色元数据 | 条目-坐标链接 | SAS / Python | P0 | p8 |
| POS-04 | F. 批注位置映射 | SRCCRF 与 BLANKCRF 按 form name + sentence 合并；句子重复时派生 SEQ 变量按 form name 计数，保证一对一匹配 | SRCCRF + BLANKCRF | 一一对应的合并表 | SAS | P0 | p8；p1 摘要步骤4 |
| POS-05 | F. 批注位置映射 | 建立字符宽度表：字体（Arial、Courier New、Times New Roman）× 字号（8–12）组合下每个字符的实测宽度（抽取方法同元数据抽取） | 字符样本 PDF | 字符宽度表（Display 13） | 同 PDF 抽取方法 | P0 | p8 Display 13 |
| POS-06 | F. 批注位置映射 | 字符按 4 类统计并设为宏变量：A–Z、a–z、其他符号平均宽度、空格宽度 | 字符宽度表 | 4 组宽度宏变量 | SAS macro | P0 | p8 |
| POS-07 | F. 批注位置映射 | 批注框尺寸计算：宽度 = Σ(各类字符计数 × 对应宽度宏变量)；高度 = 批注内最大字符高度 | 批注文本 + 宏变量 | 批注框宽高 | SAS（%annomap，Display 14） | P0 | p8/p10-11 |
| POS-08 | F. 批注位置映射 | 单行批注：起点为 item 正下方起点，宽高算出即完成定位；次级批注（second annotation）用同样方法定位 | 批注框宽高 + 起点 | 单行批注矩形框坐标 | SAS | P0 | p8-9 |
| POS-09 | F. 批注位置映射 | 多行折行：以列1/列2 分割参考线为横向切点；行数 = 批注总宽 ÷ (x_2 − 批注起点 x)，据此推导终点坐标 | 单行批注 + x_2 参考线 | 多行批注矩形框坐标 | SAS | P0 | p9 Display 12 |
| POS-10 | F. 批注位置映射 | 自动定位不 100% 替代人工：仍需相对少量的人工复核与调整 | aCRF 初稿 | 人工修订后的 aCRF | 人工 | P1 | p9 |
| XFDF-01 | G. XFDF 生成与导入 | 翻译与 SDTM 批注两套文件均以 .xfdf 格式生成，并导入已翻译的 BLANKCRF，完成 aCRF 主体 | 翻译元数据 + 背景色 + DED | .xfdf 文件 → aCRF | SAS 程序整合生成 | P0 | p1 摘要；p3 Display 2；p9 |
| LIB-01 | H. 复用库 | 首个研究完成后生成两个 SAS 库数据集：transmeta（翻译）与 sdsmeta（批注），作为后续研究的初始化库 | 首研输出 | transmeta / sdsmeta SAS 数据集 | SAS | P1 | p9 |
| LIB-02 | H. 复用库 | 两个库数据集支持跨研究迭代累积，扩大覆盖面 | 各研究增量数据 | 更新后的库数据集 | SAS | P1 | p9 |

---

## 3. 关键数据结构与中间文件

| 文件 / 数据 | 说明 | 关键字段 |
|---|---|---|
| PDFBox txt | SRCCRF 全字符抽取结果 | 字符、坐标 (x_s, x_e, y_s, y_e)、页码 |
| 翻译元数据 Excel | 字符组装（词/短语/句/段）+ 坐标，供中文翻译 | 句子、原始坐标、调整坐标 (xs, xe, ys, ye)、译文 |
| CRF CN metadata Excel | 中文翻译结果，作为批注链接源 | form name、external ID、中文译文、条目坐标 |
| SDS（Vault 导出） | CRF 条目属性全集 | form name/label、items group name/label、item name/label、external ID、data type、codelist |
| DED | 每 item 的 SDTM 批注（仅 2 组） | SDTM Variables Name / SuppQual Desc；Second SDTM Variables Name / Second SuppQual Desc |
| 背景色元数据 | BLANKCRF 与 SRCCRF 分区背景信息 | Variable Page、left x/left y/right x/right y、RGB、参考线 x_2、x_3 |
| 字符宽度表 | 3 字体 × 5 字号（8,9,10,11,12）每字符实测宽度 | 字体、字号、字符、宽度 |
| transmeta / sdsmeta | 跨研究复用库（SAS 数据集） | 翻译元数据 / SDS 批注元数据 |
| XFDF | 最终导入格式（翻译与批注两份） | 翻译文本、批注框坐标与内容 |

**坐标模型**（Display 8）：所有字符/框以左上角为起点；原始矩形 (x_s, x_e, y_s, y_e) 用于提取，调整矩形 (xs, xe, ys, ye) 将终点拉伸至单元格右下角，用于遮罩防原文露出。

**SAS 宏参考**（Display 14）：`%annomap(indat=annot2, outdat=annot3, var=_charalp1, id=1)`，保留变量 `page x_2 x_s y_s y_e _annotats: annotats: fontsize key _charalp: l_char: u_l_charen: l_l_charen: spwid _lsp: _loth: wid_:`。

---

## 4. 假设与约束

| 编号 | 约束 / 假设 | 来源 |
|---|---|---|
| C-01 | 仅面向 Veeva Vault EDC 系统（2–3 列结构化表单、混合背景色）；RAVE 系统细节不在本文范围 | Introduction |
| C-02 | 翻译与批注文件均以 .xfdf 导入空白 CRF（非直接改 PDF） | Abstract |
| C-03 | DED 批注仅 2 组（主批注 + Second 批注） | p7 |
| C-04 | 自动批注定位仍需人工复核兜底，目标是最小化而非消除人工 | p9 |
| C-05 | 实现为 Python + SAS 组合；作者认为单语言亦可完整实现 | Conclusion |
| C-06 | 法规前提：NMPA 要求外文 CRF 中文翻译（中国单中心或含中国亚组的全球研究均适用） | p2 |

---

## 5. 追溯矩阵：论文方法论步骤 ↔ 需求条目

| 论文步骤 / 工作流 | 对应 REQ-ID |
|---|---|
| 摘要步骤 1：首研 SDS 手工批注 → 迭代复用 | DED-02、LIB-01、LIB-02 |
| 摘要步骤 2：精确计算 BLANKCRF 各条目位置 | ASM-01 ~ ASM-04、POS-05 ~ POS-09 |
| 摘要步骤 3：SRCCRF × SDS 按 external ID 合并 | SDS-02、POS-02、POS-03 |
| 摘要步骤 4：BLANKCRF × SRCCRF 按 form name/section head/item 合并 | POS-04 |
| 工作流左路：英文元数据抽取（PDFBox txt → Python 组装 → 翻译 Excel） | PDF-01、PDF-02、ASM-01 ~ ASM-05 |
| 工作流右路：a) 中文翻译 b) 背景色提取 c) SAS 整合导出 XFDF | TRN-01、TRN-02、BGC-01 ~ BGC-03、XFDF-01 |
| 批注位置映射（减少人工移动批注） | POS-01 ~ POS-10 |
