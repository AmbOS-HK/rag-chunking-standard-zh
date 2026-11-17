import fs from "node:fs";
import path from "node:path";
import {
  AlignmentType,
  Document,
  Footer,
  FootnoteReferenceRun,
  Header,
  HeadingLevel,
  ImportedXmlComponent,
  LevelFormat,
  Packer,
  PageNumber,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  convertInchesToTwip,
} from "docx";

const outputPath = process.argv[2];
if (!outputPath) throw new Error("Usage: node create_review.js <输出文件.docx（绝对或相对路径）>");

const T = String.raw;
const title = T`《RAG 知识库切片质量标准与验收清单 V3.0》平台冲突核查与文献适用性评审意见（修订版 V2）`;

const palette = {
  dark: "263238",
  primary: "37474F",
  light: "78909C",
  fill: "EEF3F6",
  warn: "FDF3E7",
};

const font = {
  ascii: "Times New Roman",
  hAnsi: "Times New Roman",
  cs: "Times New Roman",
  eastAsia: "SimSun",
};

const run = (text, options = {}) => new TextRun({ text, font, size: 24, ...options });
const para = (children, options = {}) =>
  new Paragraph({
    spacing: { after: 160, line: 300 },
    ...options,
    children: Array.isArray(children) ? children : [children],
  });
const p = (text) => para(run(text), { indent: { firstLine: convertInchesToTwip(0.33) } });
const h1 = (text) =>
  para(run(text, { bold: true, size: 30, color: palette.dark }), {
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300, after: 140 },
  });
const h2 = (text) =>
  para(run(text, { bold: true, size: 26, color: palette.primary }), {
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
  });
const bullet = (text, ref) =>
  para(ref ? [run(text), ref] : [run(text)], {
    numbering: { reference: "bullets", level: 0 },
  });

const cell = (content, options = {}) =>
  new TableCell({
    children: Array.isArray(content) ? content : [para(run(content), { spacing: { after: 40, line: 260 } })],
    margins: { top: 100, bottom: 100, left: 100, right: 100 },
    ...options,
  });
const hcell = (text, width, fill = palette.fill) =>
  cell([para(run(text, { bold: true, size: 22 }), { spacing: { after: 40, line: 260 } })], {
    shading: { type: ShadingType.CLEAR, fill },
    width: { size: width, type: WidthType.DXA },
  });
const bcell = (text, width, fill) =>
  cell([para(run(text, { size: 22 }), { spacing: { after: 40, line: 260 } })], {
    width: { size: width, type: WidthType.DXA },
    ...(fill ? { shading: { type: ShadingType.CLEAR, fill } } : {}),
  });

const xmlEscape = (value) =>
  String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");

const toc = (entries) => {
  const cached = entries
    .map(({ title: t, level, page }) => {
      const indent = Math.max(0, level - 1) * 360;
      return `<w:p><w:pPr><w:pStyle w:val="TOC${level}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs><w:ind w:left="${indent}"/></w:pPr><w:r><w:t>${xmlEscape(t)}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${page}</w:t></w:r></w:p>`;
    })
    .join("");
  return ImportedXmlComponent.fromXmlString(`<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:sdtPr><w:alias w:val="目录"/></w:sdtPr>
    <w:sdtContent>
      <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/>
        <w:instrText xml:space="preserve"> TOC \\o &quot;1-2&quot; \\h \\z \\u </w:instrText>
        <w:fldChar w:fldCharType="separate"/></w:r></w:p>
      ${cached}
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    </w:sdtContent>
  </w:sdt>`).root[0];
};

// ---------- 平台对照总表 ----------
const platformWidths = [1300, 1900, 1500, 1300, 1300, 1700];
const platformTable = new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: platformWidths,
  rows: [
    new TableRow({
      children: [
        hcell("平台", platformWidths[0]),
        hcell("切片策略", platformWidths[1]),
        hcell("长度参数", platformWidths[2]),
        hcell("overlap 配置", platformWidths[3]),
        hcell("配置可改性", platformWidths[4]),
        hcell("与文件关系", platformWidths[5]),
      ],
    }),
    new TableRow({
      children: [
        bcell("阿里云百炼", platformWidths[0]),
        bcell("智能切分（默认，分句+语义选点）/按长度/按页/按标题/按正则/按符号", platformWidths[1]),
        bcell("10–6000 token，默认 600（token 单位）", platformWidths[2]),
        bcell("控制台未暴露 overlap 参数", platformWidths[3]),
        bcell("切片方式与 Metadata 创建后不可更改", platformWidths[4]),
        bcell("无硬冲突；表格行级切分为实质差异（见 4.1）", platformWidths[5]),
      ],
    }),
    new TableRow({
      children: [
        bcell("百度千帆", platformWidths[0]),
        bcell("默认切分/自定义切片（按页、正则、标识符）/整文件切片（单切片最高 15 万字）", platformWidths[1]),
        bcell("默认 1200 字（字符），上限 15 万字", platformWidths[2]),
        bcell("重叠最大字数占比可配置（可设 0；正则切分时不可用）", platformWidths[3]),
        bcell("自定义切片支持调整", platformWidths[4]),
        bcell("兼容；默认 1200 字超出文件经验区间上限（张力，C 级）", platformWidths[5]),
      ],
    }),
    new TableRow({
      children: [
        bcell("CodeArts 盘古助手", platformWidths[0]),
        bcell("平台内置切分器（当前只支持中文切分器）", platformWidths[1]),
        bcell("50–1000，默认 300（单位未明示）", platformWidths[2]),
        bcell("无 overlap 配置项", platformWidths[3]),
        bcell("仅入库配置，选项有限", platformWidths[4]),
        bcell("无冲突；存在能力缺口（无标题路径/元数据/overlap 配置）", platformWidths[5]),
      ],
    }),
    new TableRow({
      children: [
        bcell("盘古大模型 Agent 平台", platformWidths[0]),
        bcell("自定义：分句字符标识/分段最大长度/分段重叠最大字数占比；AgentArts 系另支持层级分段（标题 1–10 级）", platformWidths[1]),
        bcell("自定义可配；AgentArts 系 1–6000，默认 500", platformWidths[2]),
        bcell("重叠最大字数占比可配置（范围默认值公开文档未载明）", platformWidths[3]),
        bcell("创建后可修改并重新解析切分", platformWidths[4]),
        bcell("四者中兼容性最好，无冲突项", platformWidths[5]),
      ],
    }),
  ],
});

// ---------- 差异点清单表 ----------
const findingWidths = [600, 1100, 3400, 1500, 2400];
const findingsTable = new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: findingWidths,
  rows: [
    new TableRow({
      children: [
        hcell("编号", findingWidths[0]),
        hcell("平台", findingWidths[1]),
        hcell("差异点", findingWidths[2]),
        hcell("性质", findingWidths[3]),
        hcell("处理建议", findingWidths[4]),
      ],
    }),
    new TableRow({
      children: [
        bcell("F-1", findingWidths[0]),
        bcell("百炼", findingWidths[1]),
        bcell("表格文件按数据行切分（表头拼接到每行），与文件§五“表格整块成切分单元”不同", findingWidths[2]),
        bcell("实质差异（目的相同，路径不同）", findingWidths[3], palette.warn),
        bcell("在§五增加平台例外说明：百炼表格行级切分+表头拼接可视为等效合规路径，需以检索验收结果判效", findingWidths[4]),
      ],
    }),
    new TableRow({
      children: [
        bcell("F-2", findingWidths[0]),
        bcell("百炼", findingWidths[1]),
        bcell("最大分段长度单位为 token（默认 600），文件经验区间为 300–800 中文字；官方另建议长篇报告 2048 token，超出文件区间上限", findingWidths[2]),
        bcell("张力（C 级经验值 vs 平台默认）", findingWidths[3]),
        bcell("维持区间不变，注明其为中文场景基线而非上限；平台 token 单位按§六“先确认单位”执行", findingWidths[4]),
      ],
    }),
    new TableRow({
      children: [
        bcell("F-3", findingWidths[0]),
        bcell("百炼", findingWidths[1]),
        bcell("切片方式、Metadata 抽取创建后不可更改；文件§十“升一档重导验证”在该平台等于新建知识库并全量重新导入，且元数据不可补配", findingWidths[2]),
        bcell("注意点（处置成本）", findingWidths[3]),
        bcell("§十处置路径增加平台成本提示，避免低估回归代价", findingWidths[4]),
      ],
    }),
    new TableRow({
      children: [
        bcell("F-4", findingWidths[0]),
        bcell("千帆", findingWidths[1]),
        bcell("切片最大长度默认 1200 字（字符），超出文件 300–800 字区间上限；overlap 占比可配置且可设 0", findingWidths[2]),
        bcell("张力（C 级 vs 平台默认）；overlap 兼容", findingWidths[3]),
        bcell("无需改文件；评审记录中说明默认更粗属平台取向，以检索验收为最终判据", findingWidths[4]),
      ],
    }),
    new TableRow({
      children: [
        bcell("F-5", findingWidths[0]),
        bcell("盘古助手", findingWidths[1]),
        bcell("无标题路径、元数据、overlap 配置项；只支持中文切分器；“分段预计长度”单位未明示", findingWidths[2]),
        bcell("能力缺口（检查项无可执行对象）", findingWidths[3]),
        bcell("增加通用条款：平台不暴露配置项时登记为“平台约束”，不判为不达标；入库前确认长度单位", findingWidths[4]),
      ],
    }),
    new TableRow({
      children: [
        bcell("F-6", findingWidths[0]),
        bcell("盘古 Agent 平台", findingWidths[1]),
        bcell("“分段重叠最大字数占比”的取值范围与默认值在公开文档中未查到", findingWidths[2]),
        bcell("待补信息", findingWidths[3]),
        bcell("实测确认后补入评审记录", findingWidths[4]),
      ],
    }),
  ],
});

// ---------- 论文适用性分层表 ----------
const layerWidths = [2600, 2200, 4200];
const layerTable = new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: layerWidths,
  rows: [
    new TableRow({
      children: [hcell("论文结论", layerWidths[0]), hcell("可迁移性", layerWidths[1]), hcell("理由", layerWidths[2])],
    }),
    new TableRow({
      children: [
        bcell("索引体积按 1/(1−o) 膨胀（o 为 overlap 比例）", layerWidths[0]),
        bcell("可直接采信", layerWidths[1]),
        bcell("算术事实，与语言、检索器、模型均无关", layerWidths[2]),
      ],
    }),
    new TableRow({
      children: [
        bcell("句子切分性价比最高（≈语义切分至约 5k tokens）；上下文不是越多越好；按任务目标调上下文预算", layerWidths[0]),
        bcell("方向可借鉴，数值需复验", layerWidths[1]),
        bcell("方向性结论大概率跨语言成立；但实验基座为英文 NQ + SPLADE（稀疏）+ Ministral-8B，中文分词粒度与稠密检索器行为不同", layerWidths[2]),
      ],
    }),
    new TableRow({
      children: [
        bcell("10–20% overlap 无可测量收益（|ΔBERTScore| ≤ 0.004）", layerWidths[0]),
        bcell("不可直接平移", layerWidths[1], palette.warn),
        bcell("零收益的机制解释依赖稀疏检索器特性；中文稠密检索实验（CRUD-RAG）与金融 PDF 实验（arXiv:2604.12047）均测得 overlap 正向收益，见 5.3", layerWidths[2]),
      ],
    }),
    new TableRow({
      children: [
        bcell("上下文悬崖约 2.5k tokens；语义质量峰值约 500 tokens", layerWidths[0]),
        bcell("不可直接套用", layerWidths[1], palette.warn),
        bcell("拐点位置模型相关（Ministral-8B 实测）；中文模型（Qwen/ERNIE/盘古）拐点需按所用 LLM 复调，文件已正确标注此限定", layerWidths[2]),
      ],
    }),
  ],
});

const tocEntries = [
  { title: "1 评审对象与范围", level: 1, page: 2 },
  { title: "2 评审方法与依据", level: 1, page: 2 },
  { title: "3 术语说明：稠密检索与稀疏检索", level: 1, page: 3 },
  { title: "4 平台冲突核查", level: 1, page: 3 },
  { title: "4.1 阿里云百炼", level: 2, page: 3 },
  { title: "4.2 百度千帆", level: 2, page: 4 },
  { title: "4.3 华为 CodeArts 盘古助手", level: 2, page: 4 },
  { title: "4.4 华为盘古大模型 Agent 平台", level: 2, page: 5 },
  { title: "4.5 平台对照总表与差异点清单", level: 2, page: 5 },
  { title: "5 arXiv:2601.14123 对中文知识库的适用性评审", level: 1, page: 6 },
  { title: "5.1 论文核验", level: 2, page: 6 },
  { title: "5.2 结论分层：可迁移与不可迁移", level: 2, page: 7 },
  { title: "5.3 反例研究与中文场景证据", level: 2, page: 7 },
  { title: "5.4 对文件引用方式的意见", level: 2, page: 8 },
  { title: "6 修改建议清单（按优先级）", level: 1, page: 9 },
  { title: "7 评审局限性声明", level: 1, page: 9 },
  { title: "8 参考来源", level: 1, page: 10 },
];

const children = [
  // 封面标题区
  para(run(title, { bold: true, size: 34, color: palette.dark }), {
    heading: HeadingLevel.TITLE,
    alignment: AlignmentType.CENTER,
    spacing: { before: 480, after: 240 },
  }),
  para(run("评审日期：2026-11-17　　评审方式：AI 辅助核查（全文比对 + 官方文档核验 + 论文原文核验）", { color: palette.primary }), {
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
  }),
  para(run("评审对象版本：V3.0（2026-11-17，正式评审版）　　本版说明：修订版 V2，补入中文场景证据 CRUD-RAG、中文法律文书基准 Legal-DC 与 ACM 切片评估框架 HOPE", { color: palette.primary }), {
    alignment: AlignmentType.CENTER,
    spacing: { after: 480 },
  }),
  para(run("目录", { bold: true, size: 30, color: palette.dark }), { heading: HeadingLevel.HEADING_1 }),
  para(run("页码为估计值，可在 Word/WPS 中右键目录选择“更新域”刷新。", { italics: true, color: palette.light })),
  toc(tocEntries),

  // 1
  h1("1 评审对象与范围"),
  p(T`评审对象为《RAG 知识库切片质量标准与验收清单》V3.0（正式评审版，2026-11-17）。该文件为厂商中立的内部技术规范，依据 ITU-T F.748.52（2025）与 GB/Z《人工智能关键技术 检索增强生成系统测评规范》（征求意见稿）制定（注：该标准已于 2026-08-27 正式发布为 GB/Z 236—2026《人工智能 检索增强生成系统评价指标与方法》，名称与征求意见稿不同），全部数值合格线标注为 C 级经验值。`),
  para([
    run(T`本次评审回答两个委托问题：（1）该文件与阿里云百炼、百度千帆、华为盘古三个平台的知识库切片标准/默认行为是否存在冲突；（2）文件引用的单篇英文实验论文 arXiv:2601.14123 的结论能否适用于该文件所面向的中文知识库切片质量标准，并补充检索适用于中文场景的权威证据。盘古平台按其两个知识库入口分别核查：CodeArts 盘古助手（研发场景）与盘古大模型 Agent 开发平台`),
    new FootnoteReferenceRun(5),
    new FootnoteReferenceRun(6),
    run("。"),
  ]),

  // 2
  h1("2 评审方法与依据"),
  p(T`评审分四轮执行，全部关键事实均可追溯至一手来源：`),
  bullet(T`第一轮：通读评审对象全文（13 页），提取全部标准条款、数值口径与文献引用点。`),
  bullet(T`第二轮：逐平台核对官方文档——百炼“切片与向量化/创建知识库/RAG 效果优化”文档、千帆“知识库简介/创建知识库”文档（已抓取全文）、华为云盘古助手“约束与限制”与盘古大模型产品功能/上传文档文档。`),
  bullet(T`第三轮：核验 arXiv:2601.14123 摘要与关键段落原文；就该论文的适用范围补充检索独立评述与反例研究，找到反例论文 arXiv:2604.12047 原文并核对其数据表（Table 11）；补充检索中文场景证据，找到中文 RAG 基准 CRUD-RAG（arXiv:2401.17043，ACM 收录）并核对其 chunk size 与 overlap 实验数据（Table 3/4）。`),
  bullet(T`第四轮（应委托方要求追加）：补充检索中文文书场景的权威证据——找到中文法律文书 RAG 基准 Legal-DC（arXiv:2603.11772，预印本）并核验其摘要、数据集构建与实验结论（含 Table 3/7）；找到 ACM SIGIR 2025 正式收录的切片质量评估框架 HOPE（DOI 10.1145/3726302.3729882）并核验其收录页面与摘要结论。`),
  para([
    run(T`性质判定口径：标准条款与平台产品参数默认值属于不同层面，数值不一致本身不构成“冲突”；仅当平台行为使文件条款无法满足或语义相反时判为“实质差异”。文件中所有数值合格线为 C 级经验值（文件自定声明）`),
    new FootnoteReferenceRun(11),
    run("。"),
  ]),

  // 3
  h1("3 术语说明：稠密检索与稀疏检索"),
  p(T`本节为理解第 5 章结论的前提。RAG 检索分两大流派：稀疏检索（Sparse）基于词面匹配（BM25、TF-IDF、SPLADE），问题与切片中相同词项越多得分越高；稠密检索（Dense）由 embedding 模型将文本压缩为数百至数千维的稠密向量，按向量空间语义距离匹配，能识别“退费/退款”这类同义不同词的情况。`),
  p(T`与本评审直接相关的事实：arXiv:2601.14123 使用的 SPLADE 属于稀疏检索；而百炼（知识库默认官方向量 text-embedding-v2，1536 维稠密向量）、千帆（bge/qwen3-embedding 等四款）、盘古（m3e-base、bge-large-zh-v1.5 等）的默认检索底座全部为稠密向量模型。因此该论文中依赖稀疏检索器机制解释的结论，不能默认平移到三个平台。`),

  // 4
  h1("4 平台冲突核查"),
  h2("4.1 阿里云百炼"),
  para([
    run(T`已核对要点：切片方式六种（智能切分为默认，先分句再按语义相关性自适应选切点）；最大分段长度 10–6000，默认 600，单位为 token；切片方式与 Metadata 抽取在知识库创建后不可更改；表格文件按数据行切分并将表头拼接到每个数据行切片；官方选型建议为 FAQ 256 token、产品手册 600 token、长篇报告 2048 token、法律合同按标题 512 token`),
    new FootnoteReferenceRun(1),
    new FootnoteReferenceRun(2),
    run("。"),
  ]),
  p(T`与文件的关系：无硬性冲突。需关注三点——（1）单位差异：百炼用 token，文件经验区间用中文字，文件§六“调参前先确认单位”的纪律恰好覆盖，600 token 约合数百中文字（估算），与 300–800 字区间大体兼容；（2）实质差异：百炼表格按数据行切分是平台固定行为，与文件§五“表格整块成切分单元”路径不同，但两者目的相同（防语义混淆），且百炼通过表头拼接保留了列名语义；（3）处置成本：文件§十“升一档重导验证”在百炼上等于新建知识库并全量重新导入，且元数据配置不可补配。`),

  h2("4.2 百度千帆"),
  para([
    run(T`已核对要点（已抓取官方文档全文）：切片策略为默认切分/自定义切片/整文件切片三种；自定义切片支持按页、按自定义正则、按标识符切分；切片最大长度默认 1200 字（字符），上限 15 万字；切片重叠最大字数占比可配置（按字符数占比，选择正则切分时该功能不可用）；“关联信息”可将文件名、标题及子标题拼入切片；知识增强（问题生成默认开启、段落概要、三元组抽取）；整文件切片单切片最高 15 万字，面向长上下文全文总结场景`),
    new FootnoteReferenceRun(3),
    new FootnoteReferenceRun(4),
    run("。"),
  ]),
  p(T`与文件的关系：兼容。overlap 可显式配置且可设 0，与文件“默认不设 overlap、上限 20%”完全兼容；“关联信息”机制与文件§三“标题路径必填”方向一致（均为切片自包含服务），属于互相印证。两点张力：默认 1200 字超出文件经验区间上限（平台取向更粗，非冲突）；整文件切片与文件引用的“上下文悬崖”结论场景不同（全文总结 vs 检索问答），不构成矛盾。`),

  h2("4.3 华为 CodeArts 盘古助手"),
  para([
    run(T`已核对要点（华为云官方文档）：切分器当前只支持中文切分器；分段预计长度取值 50–1000、默认 300（单位未明示字符或 token）；嵌入模型为 m3e-base（768 维中英文）、bge-large-zh-v1.5（仅中文）、paraphrase-multilingual-MiniLM-L12-v2、华为自研代码模型 VESO-v1.2；仅支持 TXT/DOCX（单文件 ≤20MB）与代码 ZIP；未暴露 overlap、标题路径、元数据等配置项`),
    new FootnoteReferenceRun(5),
    run("。"),
  ]),
  p(T`与文件的关系：无冲突，但存在能力缺口——文件§三要求“标题路径必填”、§六要求检查 overlap 设置，该平台这些参数不暴露，检查项无可执行对象；默认 300 落在文件 300–800 字区间下限；VESO 代码模型与文件§五“代码按语法单元切”方向一致。建议按第 6 章 F-5 条款处理。`),

  h2("4.4 华为盘古大模型 Agent 平台"),
  para([
    run(T`已核对要点（华为云官方文档）：知识库为开箱即用的企业级 RAG 服务；切片策略支持文本切分自定义配置（分句字符标识、分段最大长度、分段重叠最大字数占比），参数体系与千帆几乎一致；文档解析与分段策略创建后支持修改并可重新解析切分——与百炼“创建后不可更改”相反，文件“升档重导”的处置成本在该平台很低；支持 FAQ 模板与 20 余种文档格式（单文件 ≤128MB）`),
    new FootnoteReferenceRun(6),
    new FootnoteReferenceRun(7),
    run("。"),
  ]),
  para([
    run(T`另据华为智能体平台（AgentArts/Versatile 系）用户指南：拆分设置支持自动分段、长度分段、层级分段三种；分段预计长度取值 1–6000、默认 500；层级分段支持标题层级深度 1–10 级与多级标题组合保存进切片——与文件“结构感知切分”“标题路径”要求高度对应`),
    new FootnoteReferenceRun(8),
    run("。"),
  ]),
  p(T`与文件的关系：四者中兼容性最好，无冲突项。一项待补信息：“分段重叠最大字数占比”的取值范围与默认值未在公开文档中查到，建议实测确认后补入评审记录。`),

  h2("4.5 平台对照总表与差异点清单"),
  platformTable,
  para(run("差异点清单（F-1 至 F-6）", { bold: true, color: palette.primary }), { spacing: { before: 240, after: 120 } }),
  findingsTable,
  p(T`小结：三个平台（四个入口）与文件均无硬性冲突。真正需要评审处理的只有 F-1（实质差异，建议增加平台例外说明）与 F-5（能力缺口，建议增加平台约束登记条款），其余为 C 级经验值与平台默认值之间的正常张力。`),

  // 5
  h1("5 arXiv:2601.14123 对中文知识库的适用性评审"),
  h2("5.1 论文核验"),
  para([
    run(T`论文真实存在且与文件引用一致：《A Systematic Analysis of Chunking Strategies for Reliable Question Answering》（École polytechnique，2026-01-20 提交，ECIR 2026 短篇，正文 3 页）。实验配置：Natural Questions 数据集（英文）+ SPLADE 稀疏检索 + Ministral-8B 生成模型。文件引用的四条结论均与摘要逐条吻合：F1 overlap 无可测量收益且索引按 1/(1−o) 膨胀；F2 句子切分≈语义切分（至约 5k tokens）优于定长与代码切分；F3 约 2.5k tokens 后出现“上下文悬崖”；F4 最优上下文预算依任务目标而定`),
    new FootnoteReferenceRun(9),
    run("。"),
  ]),
  para([
    run(T`同时核对了独立第三方评述：该论文为单数据集、单检索器、两档 overlap（0% 与 20%，非扫描）的短篇实验，未测试生产环境主流的稠密向量检索器，论文作者亦承认结果不具普适代表性；作者自己的建议原文是“除非有证据表明所用检索器受益于边界冗余，否则使用零 overlap”——与文件的限定语一致`),
    new FootnoteReferenceRun(10),
    run("。"),
  ]),

  h2("5.2 结论分层：可迁移与不可迁移"),
  p(T`对“单篇英文实验能否适用于中文知识库切片质量标准”的回答：定性结论可借鉴，数值结论不可直接套用。分层如下表。`),
  layerTable,

  h2("5.3 反例研究与中文场景证据"),
  para([
    run(T`反例研究（英文，金融 PDF）：El Bachyr 等（卢森堡大学 + BGL 法国巴黎银行），《Empirical Evaluation of PDF Parsing and Chunking for Financial Question Answering with RAG》，arXiv:2604.12047，ICSE-SEIP 2026 正式发表。其在金融 PDF 问答上固定 neural chunker 扫描 overlap：25% overlap 检索表现最优（MRR 由 0.529 升至 0.658（FinanceBench）、由 0.735 升至 0.833（TableQuest）），0% overlap 得分最差，50% 反而回落`),
    new FootnoteReferenceRun(12),
    run("。"),
  ]),
  para([
    run(T`中文场景直接证据：CRUD-RAG（arXiv:2401.17043，ACM TOIS 期刊正式收录（43 卷 2 期，2025，DOI 10.1145/3701228））——基于近 30 万篇中文新闻构建的中文 RAG 基准，使用稠密向量检索（bge-base 等四款中文 embedding），系统扫描 chunk size（64/128/256/512）与 chunk overlap（0%–70%）。其结论：（1）最优 chunk size 因任务而异——单文档问答小切片更好（128 优于 512），文本续写与多文档问答大切片更好，印证文件“按文档结构选型、不一刀切”的立场；（2）overlap 对问答任务关键信息精确度与召回有正向作用：单文档 QA 的 RAGQuestEval 精确度由 52.67（0%）升至 54.06（70%）、召回由 70.82 升至 73.94，与 2601.14123 的 F1 方向相反`),
    new FootnoteReferenceRun(13),
    run("。"),
  ]),
  para([
    run(T`中文文书直接证据：Legal-DC（arXiv:2603.11772，2026-03，预印本，未经同行评审）——目前最贴合中文文书场景的实验证据，专门针对中国法律文书的 RAG 基准：480 份法律、行政法规与部门规章（覆盖市场监管、合同管理等）＋2,475 条条款级标注问答（两轮人工标注，Cohen's Kappa ≥ 0.83）。其 LegRAG 框架的实验结论：检索粒度是法律 RAG 性能的关键因素；条文级（article-level，以“第×条”为独立单元）结构切分与语义 chunk 切分并行、多粒度检索优于任一单一粒度（消融实验 Table 7）；整体检索 Recall 78.02%、MRR@5 50.23%，优于 LightRAG（75.86% / 47.56%）等基线`),
    new FootnoteReferenceRun(14),
    run("。"),
  ]),
  para([
    run(T`ACM 通用切片论证：HOPE（A New HOPE: Domain-agnostic Automatic Evaluation of Text Chunking，ACM SIGIR 2025 正式收录，DOI 10.1145/3726302.3729882）——领域无关的切片质量自动评估框架，跨七个领域的实证发现“切片间语义独立性”是对系统表现影响最大的切片特征（事实正确性提升最高 56.2%、答案正确性提升最高 21.1%）；而“切片内概念单一”这一传统假设影响甚微。其为 ACM 正式收录（证据等级高于预印本），但实验为通用领域、非中文专属`),
    new FootnoteReferenceRun(15),
    run("。"),
  ]),
  p(T`两篇反方向证据的层次澄清：2601.14123 的 F1 测的是端到端生成质量（BERTScore/EM），El Bachyr 与 CRUD-RAG 的 overlap 收益测的是检索层或关键信息层指标（MRR、RAGQuestEval 精确度/召回）。三者并非平面矛盾——更准确的综合表述是：“overlap 对检索召回与关键信息覆盖可能有收益，对端到端答案质量未必有收益”。CRUD-RAG 的局限（2024 年初、新闻语料、当时代模型）同样意味着其数值不可直接搬用；但其语言（中文）与检索器类型（稠密）与本文件目标场景一致，证据权重高于英文稀疏检索实验。`),
  p(T`另需提醒：网络上仍在流传“信通院《检索增强生成技术要求与评估方法》25 能力项 + 量化阈值（元数据完整性≥90%、MRR≥0.8、幻觉率<15%）”的版本——即文件 V2.1 已勘误的对象。其源头不可核验，文件“不采纳其外部标准出处主张、数值保留为自定经验起点”的处理正确，应坚持。`),

  h2("5.4 对文件引用方式的意见"),
  p(T`文件对 arXiv:2601.14123 的处理方式本身已属审慎：全部标注 C 级、注明“模型相关应按所用 LLM 复调、不应照搬数值”、注明“英文场景单位不同不应直接换算套用”。这些限定语方向正确，引用可以保留。`),
  p(T`建议再收紧半档（非强制）：（1）§六“默认不设 overlap（o=0）”的机制解释源自稀疏检索器实验，而三个目标平台默认均为稠密向量检索，且中文稠密检索已有反方向证据（CRUD-RAG），建议改为“默认 0，但须在目标平台的稠密检索链路上做 A/B 复验后方可固化为默认值”；（2）参考依据 C 级清单补入 CRUD-RAG 作为中文场景主锚点，并补注 arXiv:2604.12047 作为检索层反例，使 overlap 条款的证据基础表述为“结论冲突，须自建基线实测”，防止评审时被质询“单篇 3 页英文短文能否支撑中文生产默认值”；（3）§五“结构感知切分”条目与参考依据补入 Legal-DC（中文法律文书基准，预印本须注明状态）与 HOPE（ACM SIGIR 2025 收录，通用切片质量评估框架），使“按文档结构选型”的立场获得中文法律文书场景的直接实验支撑与国际会议级背书。`),

  // 6
  h1("6 修改建议清单（按优先级）"),
  bullet(T`【中】§六 overlap 条款：增加“结论源自稀疏检索器实验，稠密/中文链路需 A/B 复验”限定；条款证据基础改写为三方对照——2601.14123（英文/稀疏/端到端：无收益）vs CRUD-RAG（中文/稠密：QA 关键信息指标有收益）vs arXiv:2604.12047（金融 PDF：25% overlap 检索层最优），结论为“证据冲突，须自建基线实测”。`),
  bullet(T`【中】参考依据：C 级清单补入 CRUD-RAG（arXiv:2401.17043，中文场景主锚点，注明 2024 年初/新闻语料/当时代模型的局限）、arXiv:2604.12047（检索层反例）与 Legal-DC（arXiv:2603.11772，中文法律文书基准，预印本须注明状态）；并补入 HOPE（ACM SIGIR 2025 收录，通用切片质量评估框架）；§五“结构感知切分”条目同步补注 Legal-DC 的中文法律文书实验支撑。`),
  bullet(T`【中】§五 表格处置：增加平台例外说明——百炼类平台“按数据行切分 + 表头拼接”为平台固定行为，可登记为等效合规路径，最终以检索验收结果判效。`),
  bullet(T`【低】§三/§六：增加通用条款——平台不暴露的配置项（标题路径、元数据、overlap 等）登记为“平台约束”，不判为不达标。`),
  bullet(T`【低】§十 处置路径：注明百炼类平台“升档重导”等于新建知识库并全量重新导入、元数据不可补配，提示回归成本。`),
  bullet(T`【低】§六 中文经验区间：注明 300–800 字为中文场景基线而非上限，并列出各平台默认值（千帆 1200 字、盘古助手 300、AgentArts 500、百炼 600 token）供对照。`),
  bullet(T`【低】评审记录待办：实测盘古 Agent 平台“分段重叠最大字数占比”的取值范围与默认值（公开文档未载明）。`),

  // 7
  h1("7 评审局限性声明"),
  p(T`为保证本意见本身的客观性，如实交代以下局限，供评审组权衡：`),
  bullet(T`百炼相关结论依据阿里云官方文档的检索摘要（权威度高但未逐页抓取全文）；千帆文档已抓取全文核实。`),
  bullet(T`盘古助手“分段预计长度”的单位（字符/token）官方文档未明示；“600 token 约合数百中文字”为评审方估算，无来源支撑。`),
  bullet(T`“无硬性冲突”为判断性结论，其依据是“文件厂商中立 + 数值全标 C 级”的文件性质，属评审意见而非可核验事实。`),
  bullet(T`论文核验范围：arXiv:2601.14123 核验了摘要、关键结论段落与独立第三方评述（3 页短篇未逐页精读）；arXiv:2604.12047 核验了摘要、方法与 Table 11；CRUD-RAG 核验了实验设置与 Table 3/4 的 chunk size 与 overlap 数据；Legal-DC 核验了摘要、数据集构建方法与实验结论（含 Table 3/7），其预印本状态未经同行评审；HOPE 核验了 ACM 收录页面与摘要结论。`),
  bullet(T`第 6 章修改建议为工程判断，存在合理分歧空间：文件已有的“除非有证据表明所用检索器受益”免责条款不改亦能自洽，建议收紧的原因是三平台均为稠密检索且已有中文反方向证据，免责条款位置不够显眼。`),
  bullet(T`本版（V2）相对修订版初稿的更正，如实记录：（1）§3 百炼默认向量模型由“qwen3.7-text-embedding”更正为知识库默认官方向量 text-embedding-v2（1536 维稠密）——前者为百炼向量模型列表中的在售模型，但并非知识库默认底座；（2）CRUD-RAG 收录信息补正为 ACM TOIS 43(2)（2025），DOI 10.1145/3701228 经多篇第三方文献引用复核可解；（3）四篇文献证据等级与 V3.1 分级体系统一为 C 级，“ACM 收录”等改作载体说明。`),

  // 8
  h1("8 参考来源"),
  bullet(T`[1] 阿里云百炼《切片与向量化》，help.aliyun.com/zh/model-studio/rag/chunking`),
  bullet(T`[2] 阿里云百炼《创建知识库》《容量与限制》，help.aliyun.com/zh/model-studio/rag-knowledge-base`),
  bullet(T`[3] 百度千帆《知识库简介》，cloud.baidu.com/doc/qianfan/s/Imh4stpo0`),
  bullet(T`[4] 百度千帆《创建知识库》，cloud.baidu.com/doc/qianfan/s/smh4stre1`),
  bullet(T`[5] 华为云《CodeArts 盘古助手 - RAG 知识库资产管理：约束与限制》`),
  bullet(T`[6] 华为云《盘古大模型 产品功能 - Agent 开发》，support.huaweicloud.com/productdesc-pangulm/pangulm_01_0008.html`),
  bullet(T`[7] 华为云《盘古大模型 - 创建知识库及上传文档》，support.huaweicloud.com/usermanual-pangulm/pangulm_04_0477.html`),
  bullet(T`[8] 华为云《智能体平台 用户指南》（拆分设置/层级分段），support.huaweicloud.com/usermanual-versatile/`),
  bullet(T`[9] Bennani & Moslonka, A Systematic Analysis of Chunking Strategies for Reliable Question Answering, arXiv:2601.14123, 2026-01`),
  bullet(T`[10] fatherofai.in《Debug RAG Retrieval: Why Your RAG Returns Wrong Answers》（对 2601.14123 适用范围的独立评述）`),
  bullet(T`[11] 评审对象《RAG 知识库切片质量标准与验收清单》V3.0（2026-11-17）`),
  bullet(T`[12] El Bachyr et al., Empirical Evaluation of PDF Parsing and Chunking for Financial Question Answering with RAG, arXiv:2604.12047, ICSE-SEIP 2026`),
  bullet(T`[13] Lyu et al., CRUD-RAG: A Comprehensive Chinese Benchmark for Retrieval-Augmented Generation of Large Language Models, arXiv:2401.17043；ACM Transactions on Information Systems 43(2)，2025（DOI 10.1145/3701228）`),
  bullet(T`[14] Legal-DC: Benchmarking Retrieval-Augmented Generation for Legal Documents, arXiv:2603.11772, 2026-03（预印本，未经同行评审；代码与数据：github.com/legal-dc/Legal-DC）`),
  bullet(T`[15] Brådland et al., A New HOPE: Domain-agnostic Automatic Evaluation of Text Chunking, ACM SIGIR 2025，DOI 10.1145/3726302.3729882`),
];

const doc = new Document({
  features: { updateFields: true },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "●",
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: {
                indent: { left: convertInchesToTwip(0.4), hanging: convertInchesToTwip(0.2) },
              },
            },
          },
        ],
      },
    ],
  },
  styles: {
    paragraphStyles: [
      { id: "TOC1", name: "TOC 1", basedOn: "Normal", run: { font, size: 24 } },
      { id: "TOC2", name: "TOC 2", basedOn: "Normal", run: { font, size: 24 } },
    ],
  },
  footnotes: {
    1: { children: [para(run(T`阿里云百炼《切片与向量化》《容量与限制》：最大分段长度默认 600（token），范围 10–6000；切片方式创建后不可更改。help.aliyun.com/zh/model-studio/rag/chunking`, { size: 18 }))] },
    2: { children: [para(run(T`阿里云百炼《创建知识库》《RAG 效果优化》：表格文件按数据行切分、表头拼接至每个数据行切片；Metadata 创建后不可补配。help.aliyun.com/zh/model-studio/rag-knowledge-base`, { size: 18 }))] },
    3: { children: [para(run(T`百度千帆《知识库简介》（2026-11-06 更新）：默认切分/自定义切片/整文件切片；知识增强三种方式；Small-to-Big 扩展上下文。cloud.baidu.com/doc/qianfan/s/Imh4stpo0`, { size: 18 }))] },
    4: { children: [para(run(T`百度千帆《创建知识库》：切片最大长度默认 1200 字、上限 15 万字；切片重叠最大字数占比定义与示例；关联信息（文件名、标题及子标题）。cloud.baidu.com/doc/qianfan/s/smh4stre1`, { size: 18 }))] },
    5: { children: [para(run(T`华为云《CodeArts 盘古助手 - RAG 知识库资产管理：约束与限制》（2025-03-25）：当前只支持中文切分器；分段预计长度 50–1000、默认 300；四款嵌入模型清单。`, { size: 18 }))] },
    6: { children: [para(run(T`华为云《盘古大模型 产品功能 - Agent 开发》：知识库提供开箱即用的企业级 RAG 服务；切片策略支持分句字符标识、分段最大长度、分段重叠最大字数占比自定义。support.huaweicloud.com/productdesc-pangulm/pangulm_01_0008.html`, { size: 18 }))] },
    7: { children: [para(run(T`华为云《盘古大模型 - 创建知识库及上传文档》：配置完成的文档解析、分段策略支持修改，修改后可对文档重新解析切分。support.huaweicloud.com/usermanual-pangulm/pangulm_04_0477.html`, { size: 18 }))] },
    8: { children: [para(run(T`华为云《智能体平台 用户指南》：拆分设置含自动分段/长度分段/层级分段；分段预计长度 1–6000、默认 500；标题层级深度 1–10；多级标题组合保存。support.huaweicloud.com/usermanual-versatile/`, { size: 18 }))] },
    9: { children: [para(run(T`Bennani & Moslonka, “A Systematic Analysis of Chunking Strategies for Reliable Question Answering”, arXiv:2601.14123, 2026-01-20；摘要与 F1–F4 结论原文已核验。`, { size: 18 }))] },
    10: { children: [para(run(T`fatherofai.in, “Debug RAG Retrieval: Why Your RAG Returns Wrong Answers”：指出该论文为 3 页短篇、单数据集（NQ）、单检索器（SPLADE 稀疏）、两档 overlap，稠密检索器未测。`, { size: 18 }))] },
    11: { children: [para(run(T`评审对象 V3.0 第一部分§一：“本标准中的所有数值合格线一律标为 C 级经验值”；范围声明：“不设强制性数值阈值”。`, { size: 18 }))] },
    12: { children: [para(run(T`El Bachyr et al., “Empirical Evaluation of PDF Parsing and Chunking for Financial Question Answering with RAG”, arXiv:2604.12047, ICSE-SEIP 2026；RQ3b/Table 11：25% overlap 最优（MRR 0.529→0.658 / 0.735→0.833），0% 最差；测量层次为检索层 P@1/R@3/MRR。`, { size: 18 }))] },
    13: { children: [para(run(T`Lyu et al., “CRUD-RAG: A Comprehensive Chinese Benchmark for Retrieval-Augmented Generation of Large Language Models”, arXiv:2401.17043；ACM TOIS 43(2) 收录（DOI 10.1145/3701228）；近 30 万篇中文新闻语料，稠密检索（bge-base 等）；chunk size 64–512、overlap 0%–70% 系统扫描（Table 3/4）；单文档 QA 关键信息精确度 52.67→54.06、召回 70.82→73.94（0%→70% overlap）。`, { size: 18 }))] },
    14: { children: [para(run(T`Legal-DC: Benchmarking Retrieval-Augmented Generation for Legal Documents, arXiv:2603.11772, 2026-03（预印本）；480 份中国法律法规/部门规章＋2,475 条条款级标注 QA（Cohen's Kappa ≥ 0.83）；LegRAG：chunk 级语义切分与条文级（article-level）结构切分并行多粒度检索，Recall 78.02%／MRR@5 50.23%（vs LightRAG 75.86%／47.56%）；消融实验（Table 7）：双粒度组合优于任一单一粒度。`, { size: 18 }))] },
    15: { children: [para(run(T`Brådland et al., “A New HOPE: Domain-agnostic Automatic Evaluation of Text Chunking”, ACM SIGIR 2025（DOI 10.1145/3726302.3729882）；跨七领域实证：切片间语义独立性对事实正确性提升最高 56.2%、答案正确性最高 21.1%；通用领域实验，非中文专属。`, { size: 18 }))] },
  },
  sections: [
    {
      properties: {
        page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
      },
      headers: {
        default: new Header({
          children: [
            para(run(T`RAG 切片质量标准 V3.0 平台冲突核查与文献适用性评审意见（修订版 V2）`, { bold: true, color: palette.primary, size: 18 }), {
              alignment: AlignmentType.CENTER,
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            para(new TextRun({ children: [PageNumber.CURRENT] }), { alignment: AlignmentType.CENTER }),
          ],
        }),
      },
      children,
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync(outputPath, buffer);
