import fs from "node:fs";
import {
  AlignmentType,
  Document,
  Footer,
  Header,
  HeadingLevel,
  PageNumber,
  PageOrientation,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

const outputPath = process.argv[2];
if (!outputPath) throw new Error("Usage: node create_evidence.js <输出文件.docx（绝对或相对路径）>");

const T = String.raw;
const font = { ascii: "Times New Roman", hAnsi: "Times New Roman", cs: "Times New Roman", eastAsia: "SimSun" };
const palette = { dark: "263238", primary: "37474F", light: "78909C", fill: "EEF3F6", warn: "FDF3E7" };

const run = (text, options = {}) => new TextRun({ text, font, size: 16, ...options });
const para = (children, options = {}) =>
  new Paragraph({ spacing: { after: 60, line: 240 }, ...options, children: Array.isArray(children) ? children : [children] });

const cell = (content, options = {}) =>
  new TableCell({
    children: Array.isArray(content) ? content : [para(run(content))],
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    ...options,
  });
const hcell = (text, width) =>
  cell([para(run(text, { bold: true, size: 16, color: "FFFFFF" }), { alignment: AlignmentType.CENTER })], {
    shading: { type: ShadingType.CLEAR, fill: palette.primary },
    width: { size: width, type: WidthType.DXA },
  });
const bcell = (text, width, fill) =>
  cell([para(run(text))], {
    width: { size: width, type: WidthType.DXA },
    ...(fill ? { shading: { type: ShadingType.CLEAR, fill } } : {}),
  });

// A4 横排可用宽度约 15400 twips（页宽 16838 - 边距 2×720）
const W = [2300, 2500, 2000, 1900, 3290, 1900, 1500];

const table = new Table({
  width: { size: W.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  columnWidths: W,
  rows: [
    new TableRow({
      children: ["文献（出处）", "语料与语言", "检索器 / 模型", "测量层次", "与切片相关的核心实验结论", "证据等级与局限", "可支撑的本标准条款"].map((t, i) => hcell(t, W[i])),
    }),
    new TableRow({
      children: [
        bcell("CRUD-RAG（arXiv:2401.17043；ACM TOIS 43(2) 收录，DOI 10.1145/3701228；2025）", W[0]),
        bcell("中文；近 30 万篇中文新闻", W[1]),
        bcell("稠密检索（bge-base 等四款中文 embedding）", W[2]),
        bcell("关键信息层（RAGQuestEval 精确度/召回）＋端到端", W[3]),
        bcell("chunk size 64–512、overlap 0%–70% 系统扫描；最优 chunk size 因任务而异（单文档 QA 小切片优，续写/多文档 QA 大切片优）；overlap 对单文档 QA 关键信息有正向收益：精确度 52.67→54.06、召回 70.82→73.94（0%→70%）", W[4]),
        bcell("C 级（标准分级）；载体：ACM 正式收录（DOI 可溯）。局限：2024 年初、新闻语料、当时代模型，数值不可直接搬用", W[5]),
        bcell("§六 overlap“证据冲突须自建基线”；§五按结构选型不一刀切；中文经验区间的场景锚点", W[6]),
      ],
    }),
    new TableRow({
      children: [
        bcell("El Bachyr 等（arXiv:2604.12047；ICSE-SEIP 2026）", W[0]),
        bcell("英文；金融 PDF（FinanceBench、TableQuest）", W[1]),
        bcell("neural chunker＋混合检索", W[2]),
        bcell("检索层（P@1 / R@3 / MRR）", W[3]),
        bcell("overlap 扫描：25% 最优（MRR 0.529→0.658 FinanceBench；0.735→0.833 TableQuest），0% 最差，50% 回落", W[4]),
        bcell("C 级（标准分级）；载体：ICSE-SEIP 2026 正式发表。局限：单领域英文金融 PDF，仅检索层指标，与端到端结论不可直接互斥", W[5]),
        bcell("§六 overlap 三方对照之“检索层有收益”一极", W[6]),
      ],
    }),
    new TableRow({
      children: [
        bcell("Legal-DC（arXiv:2603.11772；2026-03 预印本）", W[0]),
        bcell("中文法律文书；480 份法律/行政法规/部门规章（市场监管、合同管理等）＋2,475 条条款级标注 QA（Kappa≥0.83）", W[1]),
        bcell("向量＋关键词混合（ColBERT-Legal 等）＋Qwen3-Rerank", W[2]),
        bcell("检索层（Recall / MRR@5）＋端到端生成（LLM 评分）", W[3]),
        bcell("检索粒度是法律 RAG 的关键因素；条文级（article-level）结构切分与语义 chunk 切分并行的多粒度检索优于任一单一粒度（Table 7）；Recall 78.02% vs LightRAG 75.86%，MRR@5 50.23% vs 47.56%", W[4]),
        bcell("C 级（预印本，未经同行评审，引用须注明状态）；目前最贴合中文文书场景的直接证据", W[5]),
        bcell("§五结构感知切分（法规条款拆条）；§三标题路径/结构成分；多粒度检索设计", W[6]),
      ],
    }),
    new TableRow({
      children: [
        bcell("HOPE（A New HOPE: Domain-agnostic Automatic Evaluation of Text Chunking；arXiv:2505.02171；ACM SIGIR 2025，DOI 10.1145/3726302.3729882）", W[0]),
        bcell("通用领域，跨七领域（非中文专属）", W[1]),
        bcell("领域无关 RAG 管线（切片质量自动评估框架）", W[2]),
        bcell("切片质量评估：内在/外在/文档连贯三层特征与 RAG 指标相关性", W[3]),
        bcell("切片间语义独立性是对系统表现影响最大的切片特征：事实正确性提升最高 56.2%、答案正确性最高 21.1%；“切片内概念单一”的传统假设影响甚微", W[4]),
        bcell("C 级（标准分级）；载体：ACM SIGIR 2025 正式收录。局限：通用领域实验，非中文专属，数值不可直接平移", W[5]),
        bcell("§四语义完整性/自包含性维度；§五切片边界质量原则", W[6]),
      ],
    }),
  ],
});

const children = [
  para(run(T`RAG 切片证据速查表（评审对照用）`, { bold: true, size: 28, color: palette.dark }), {
    heading: HeadingLevel.TITLE,
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 60 },
  }),
  para(run(T`2026-11-17 · 配套《RAG 切片标准平台冲突核查与文献适用性评审意见（修订版 V2）》与《V3.1 修订草案》使用`, { size: 16, color: palette.primary }), {
    alignment: AlignmentType.CENTER,
    spacing: { after: 140 },
  }),
  table,
  para(
    run(T`对照基准：arXiv:2601.14123（英文 NQ＋SPLADE 稀疏检索＋Ministral-8B，端到端生成质量）测得 overlap 10–20% 无可测量收益（|ΔBERTScore|≤0.004）。与上表第 1、2 行的差异主要来自检索器类型（稀疏/稠密）、语言与测量层次（端到端/检索层/关键信息层），三者并非平面矛盾。综合表述：overlap 对检索召回与关键信息覆盖可能有收益，对端到端答案质量未必有收益——证据冲突，默认值须自建基线实测。`, { size: 16 }),
    { spacing: { before: 120, after: 60 } },
  ),
  para(run(T`使用说明：等级口径与标准一致（A＝ITU-T F.748.52；B＝GB/Z 236—2026；C＝公开研究/benchmark）——四篇文献按标准分级均属 C 级，载体差异（ACM 收录／会议发表／预印本）已在各行注明；C 级数值不可直接搬用，最终以目标平台自建基线的 A/B 复验为准。网传“信通院 25 能力项＋量化阈值”版本源头不可核验，不在本表收录。`, { size: 16, color: palette.light })),
];

const doc = new Document({
  styles: {},
  sections: [
    {
      properties: {
        page: {
          size: { orientation: PageOrientation.LANDSCAPE },
          margin: { top: 720, bottom: 720, left: 720, right: 720 },
        },
      },
      headers: {
        default: new Header({
          children: [para(run(T`RAG 切片证据速查表 · 2026-11-17`, { bold: true, color: palette.primary, size: 14 }), { alignment: AlignmentType.CENTER })],
        }),
      },
      footers: {
        default: new Footer({ children: [para(new TextRun({ children: [PageNumber.CURRENT], font, size: 14 }), { alignment: AlignmentType.CENTER })] }),
      },
      children,
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync(outputPath, buffer);
