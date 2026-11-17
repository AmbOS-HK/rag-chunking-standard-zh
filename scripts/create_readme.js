import fs from "node:fs";
import {
  AlignmentType,
  Document,
  Footer,
  Header,
  HeadingLevel,
  LevelFormat,
  PageNumber,
  Packer,
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
if (!outputPath) throw new Error("Usage: node create_readme.js <输出文件.docx（绝对或相对路径）>");

const T = String.raw;
const palette = { dark: "263238", primary: "37474F", light: "78909C", fill: "EEF3F6" };
const font = { ascii: "Times New Roman", hAnsi: "Times New Roman", cs: "Times New Roman", eastAsia: "SimSun" };

const run = (text, options = {}) => new TextRun({ text, font, size: 21, ...options });
const para = (children, options = {}) =>
  new Paragraph({ spacing: { after: 110, line: 280 }, ...options, children: Array.isArray(children) ? children : [children] });
const p = (text) => para(run(text), { indent: { firstLine: convertInchesToTwip(0.33) } });
const h2 = (text) =>
  para(run(text, { bold: true, size: 23, color: palette.primary }), { heading: HeadingLevel.HEADING_2, spacing: { before: 180, after: 90 } });
const bullet = (text) => para(run(text), { numbering: { reference: "bullets", level: 0 } });

const cell = (content, options = {}) =>
  new TableCell({
    children: Array.isArray(content) ? content : [para(run(content, { size: 19 }), { spacing: { after: 20, line: 240 } })],
    margins: { top: 70, bottom: 70, left: 80, right: 80 },
    ...options,
  });
const hcell = (text, width) =>
  cell([para(run(text, { bold: true, size: 19 }), { spacing: { after: 20, line: 240 } })], {
    shading: { type: ShadingType.CLEAR, fill: palette.fill },
    width: { size: width, type: WidthType.DXA },
  });
const bcell = (text, width) => cell([para(run(text, { size: 19 }), { spacing: { after: 20, line: 240 } })], { width: { size: width, type: WidthType.DXA } });

const W = [700, 3300, 5000];
const fileTable = new Table({
  width: { size: 9000, type: WidthType.DXA },
  columnWidths: W,
  rows: [
    new TableRow({ children: [hcell("顺序", W[0]), hcell("文件", W[1]), hcell("用途与关键内容", W[2])] }),
    new TableRow({ children: [
      bcell("1", W[0]),
      bcell("RAG切片标准平台冲突核查与文献适用性评审意见_V2_修订版_2026-11-17.docx", W[1]),
      bcell("先读。评审结论：三平台（四入口）无硬冲突，F-1~F-6 差异点清单；arXiv:2601.14123 结论分层（定性可借鉴、数值不可套用）；中文证据链（CRUD-RAG／Legal-DC／HOPE）；修改建议与局限性声明（含 V2 更正记录）", W[2]),
    ] }),
    new TableRow({ children: [
      bcell("2", W[0]),
      bcell("RAG切片质量标准与验收清单_V3.2_修订草案_2026-11-17.docx", W[1]),
      bcell("主文档。标准本体＋验收清单；棕褐色为新增/修订（R1–R9 修订对照）；含中文定位声明（范围段）与附录 A 中文数值基线候选清单（B1–B7，全部【待验证】）", W[2]),
    ] }),
    new TableRow({ children: [
      bcell("3", W[0]),
      bcell("RAG切片证据速查表_2026-11-17.docx", W[1]),
      bcell("评审会现场对照用。A4 横排单页：四篇文献的语料/检索器/测量层次/核心数值/等级/可支撑条款；附 overlap 三方证据的综合表述与引用口径说明", W[2]),
    ] }),
    new TableRow({ children: [
      bcell("4", W[0]),
      bcell("RAG自建基线评测集建集模板_V1.0_2026-11-17.docx", W[1]),
      bcell("执行工具。评审通过后用于落地：字段设计、标注口径、覆盖配比、六步流程、B1–B7 验证前置条件、回填留档规范", W[2]),
    ] }),
  ],
});

const children = [
  para(run(T`README · RAG 切片标准评审交付包`, { bold: true, size: 32, color: palette.dark }), {
    heading: HeadingLevel.TITLE,
    alignment: AlignmentType.CENTER,
    spacing: { before: 360, after: 160 },
  }),
  para(run(T`2026-11-17 · 评审对象：《RAG 知识库切片质量标准与验收清单》V3.0（正式评审版）`, { color: palette.primary }), {
    alignment: AlignmentType.CENTER,
    spacing: { after: 300 },
  }),

  h2("一、文件清单与阅读顺序"),
  fileTable,

  h2("二、版本与依据关系"),
  bullet(T`评审对象为 V3.0；V3.2 为修订草案（未生效），修订依据为评审意见（修订版 V2）：第 6 章第 1–6 条建议分别落入 R1／R6／R4／R3／R5／R2，第 7 条为评审记录待办（无 R 编号），R7–R9 为一致性同步与新增，R10 为评审后勘误（GB/Z 236—2026 发布，引用更新）。R10 已随 V3.2 归位；`),
  bullet(T`上位依据：ITU-T F.748.52（A 级）、GB/Z 236—2026《人工智能 检索增强生成系统评价指标与方法》（B 级，2026-08-27 发布）；文献证据全部为 C 级公开研究；`),
  bullet(T`勘误立场（V2.1 起维持）：网传“信通院 25 能力项＋量化阈值”版本源头不可核验，不采纳其外部标准出处主张。`),
  bullet(T`版本纪律：版本对应矩阵、升版检查清单与命名/归档规则见仓库根目录 VERSIONING.md；本文件由 scripts/create_readme.js 生成，与仓库 README.md 保持对齐——改动任一边须同步另一边后重新生成。`),

  h2("三、使用要点"),
  bullet(T`所有数值合格线均为 C 级经验值，不是强制阈值；附录 A 候选基线全部【待验证】，须用建集模板实测后方可回填；`),
  bullet(T`本标准面向中文语料场景；英文/多语言语料的数值基线应另行校准（V3.2 范围段定位声明）；`),
  bullet(T`overlap 条款证据冲突（稀疏端到端无收益 vs 稠密检索层有收益），以目标平台自建基线 A/B 复验为准；`),
  bullet(T`评审意见与 V3.2 中目录页码为估计值，在 Word/WPS 中右键目录“更新域”刷新。`),
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
            style: { paragraph: { indent: { left: convertInchesToTwip(0.35), hanging: convertInchesToTwip(0.18) } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: { page: { margin: { top: 1200, bottom: 1200, left: 1300, right: 1300 } } },
      headers: {
        default: new Header({
          children: [para(run(T`README · RAG 切片标准评审交付包 · 2026-11-17`, { bold: true, color: palette.primary, size: 16 }), { alignment: AlignmentType.CENTER })],
        }),
      },
      footers: {
        default: new Footer({ children: [para(new TextRun({ children: [PageNumber.CURRENT], font, size: 16 }), { alignment: AlignmentType.CENTER })] }),
      },
      children,
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync(outputPath, buffer);
