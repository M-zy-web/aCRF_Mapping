# Worklog — aCRF_Mapping

> 本文件是项目的工作日志。后续会话直接阅读本文件即可接续工作,无需重读代码或翻旧对话。
> 最新状态在顶部;每完成一个阶段就在对应日期下追加/更新进度。

---

## 2026-10-11:逆向生成 Annotations 程序首轮端到端跑通 + 全部根因已定位(修复未开始)

### 项目背景

依据 PharmaSUG China 2023 论文 AP-178《CRF Translation and aCRF Generation》的方法,从手工修改后的 [01Document/aCRF.pdf](01Document/aCRF.pdf)(61 页,789 个 PDF /FreeText 批注,文字在 /Contents、位置在 /Rect、全部 /Helv 12)逆向重建 SDTM 批注映射文件 Annotations.xlsx。

- **事实源**:`aCRF.pdf`(批注 + 页面文字坐标)
- **映射源**:`01Document/als.xlsx`(SDS:43 表单 / 360 字段 / LabTableTitle 9 列 / DataDictionaryEntry 字典)
- **模板**:`01Document/Annotations.xlsx` —— **只读第 1 行表头**(42 列),不读数据行;旧文件是旧版 PDF 的产物,几何 0/560 匹配,仅供差异对比
- **用户决策**:① 输出**新文件** `Annotations_reverse.xlsx` + QC 报告,不覆盖原文件;② **42 列全部填充**——能精确推导的精确填,内部布局列按几何规则重算(规则写入 QC 报告),真正无法确定的留空(旧文件也从未填过)
- **已批准计划**:`C:\Users\Mason\.claude\plans\delegated-sprouting-ember.md`(含已验证的关键事实、实现阶段 A–G、42 列填充语义表,实现时对照此表)

### 本次工作块时间线(跨午夜)

| 时间 | 内容 |
|---|---|
| 10-10 22:45 | 输入文件就位(aCRF.pdf / BlankCRF.pdf / Annotations.xlsx / als.xlsx / 论文 PDF) |
| 10-10 22:59 | 完成 [AP178_Paper_Requirement_Matrix.md](AP178_Paper_Requirement_Matrix.md)(论文需求矩阵,REQ-ID 编号体系) |
| 10-11 00:34 | 计划获用户批准 |
| 10-11 01:08 | [02Program/reverse_annotations.py](02Program/reverse_annotations.py) 单文件程序完成,首轮端到端运行,产出 [01Document/Annotations_reverse.xlsx](01Document/Annotations_reverse.xlsx) + [01Document/QC_report_reverse.txt](01Document/QC_report_reverse.txt) |
| 10-11 01:15 | 逐项诊断出全部失败根因,修复清单定稿(见下) |

### 当前结果(首轮,尚未修复)

- 从 PDF 提取 789 个 FreeText 批注(1 个离页跳过:p15 `SUSMOKN in SUPPSU`)
- **批注行 228 / 预期 788**;单位行 0(预期 ≈6)、表格行 0(预期 ≈5)、LabTableTitle 行 0 —— 全部几何相关匹配失效所致
- **unmatched 78 条**:p4:1、p7:1、p10:4、p14:18、p20:12、p22:9、p37:5、p40:3、p44:14、p51:4、p60:4、p61:3
- QC 报告中大量页面表单显示 `??`(detect_forms 未识别)
- 已能正确输出的部分(如 p7 IE 43 行、p22 部分 lab 行)证明文本分类、宽度计算、字典 Number(IETEST=5、IETEST1=33)等核心逻辑正确

### QC 报告阅读指南(01Document/QC_report_reverse.txt)

- `[no annot]` 列表:旧文件有批注行、但当前运行未找到对应批注的字段(修复后应基本消失)
- `unmatched` 列表:批注已提取但未能映射到字段(含矩形坐标,修复后应仅剩真正无法映射的)
- `per page`:逐页 旧行数 vs 新行数 对比
- `inference rules`:21 条推导规则与旧文件差异的文档化说明(第 7 条:旧工具宽度表略宽于标准 AFM,<2% 差异,不可复现)
- `verification errors` + `spot checks`:自动校验与抽查样本

### 待修复清单(按优先级,均已定位根因,**尚未动手修**)

1. **y 轴翻转 bug(最关键,先修)**:PDF /Rect 的 y1 是**底部**、y2 是**顶部**;程序内部把 y1 当顶部用 → 所有垂直重叠/几何匹配/选项带匹配全部失效。修复:内部统一 `top = max(rect[1], rect[3])`、`bottom = min(rect[1], rect[3])`。修完立即重跑再定后续。
2. **域批注正则缺空格**:当前正则不匹配 `LB (实验室检查)`(当前 PDF 有空格;旧 PDF 是 `LB(实验室检查)` 无空格)。修正为 `^([A-Z]{2})\s*[（(].+[）)]$`。
3. **detect_forms 表头行被拆成 3 行**:页面表头行因 x 间隙被聚类拆散 → 需按 x0 排序拼接成一行再匹配(这是大量表单显示 `??` 的原因)。
4. **值→标签包含匹配(新增规则)**:`FATEST = 问题文本` → 映射到 QSRESxx(值文本 ≥4 字且被某字段 label 包含即映射,如 "FATEST = 有胃溃疡/反流性食管炎史" → QSRES3)。
5. **SPECIAL 补充表**:
   - QS:FATESTCD、FAORRES、SUSTTPT
   - VS2:PULSE→VS2CLSG1、SYSBP→VS2CLSG2、DIABP→VS2CLSG3、TEMP→VS2TEMP(作 grid 锚点)
   - FU:SSSTAT、SSDTC、SSTEST
   - LBOTH:LBCLSIGC→LBOCLSIG、LBORRES→LBORES
6. **网格容差**:当前 ≤6pt 可能需放宽到 ≤8–10pt(修复 #1 翻转后先重跑,再根据 unmatched 决定是否放宽)。

### 验证目标(全部修复后的判定标准)

- 批注行 = 788;单位行 ≈6(p20 HRU 等);表格行 ≈5(VS2 p44 / VS1 / PC / AE / CM);LabTableTitle 行(lab 页)
- unmatched 仅剩真正无法映射的:FA 段、EGMETHOD、EGTESTCD = INTP、p61 boilerplate
- 抽查页:p4 / p7 / p20 / p22 / p44 / p51 / p60 / p61(逐项条件见已批准计划 §验证)

### 已知旧文件 quirk(**不复制**)

- 旧工具 FieldOID_Lag 在 AESEV/AEREL/AEACN/AEREL1/AEACN1 上有"滞后=自身"的 bug,新程序不复制
- 旧工具宽度表(近 AFM)与标准 Helvetica AFM 有 <2% 差异,新程序用标准 AFM,差异已在 QC 报告第 7 条注明

### 运行方式

```bash
python -I 02Program/reverse_annotations.py
```

- Windows GBK 控制台:程序输出全部写 UTF-8 文件,**勿依赖控制台回显**;测试期用 `-I` 隔离
- 依赖:pypdf、pdfplumber、openpyxl
- openpyxl 读 als.xlsx **不能用** `read_only=True`(该文件会误报空表)
- als.xlsx 的 `variableNo` 是 CRF 内部变量号,**不能**当 SDTM 变量用(SDTM 变量名只存在于批注文本)

### 下一步

1. 修 #1(y 轴翻转)→ 重跑 → 观察行数/unmatched 变化
2. 逐条修 #2–#6,每修一条重跑一次,对照验证目标
3. 全部达标后抽查验证页 + QC 报告人工复核 → 更新本 worklog(追加新日期条目)

### 参考

- 已批准计划:`C:\Users\Mason\.claude\plans\delegated-sprouting-ember.md`(42 列填充语义表在 §E)
- 论文需求矩阵:[AP178_Paper_Requirement_Matrix.md](AP178_Paper_Requirement_Matrix.md)
- Claude 记忆:projects/d--Program-Project-aCRF-Mapping/memory/reverse-annotations-project.md(与本文件互为备份,以本文件为准)
