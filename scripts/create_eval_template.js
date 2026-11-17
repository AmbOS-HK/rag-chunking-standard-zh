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
if (!outputPath) throw new Error("Usage: node create_eval_template.js <输出文件.docx（绝对或相对路径）>");

const T = String.raw;
const palette = { dark: "263238", primary: "37474F", light: "78909C", fill: "EEF3F6" };
const font = { ascii: "Times New Roman", hAnsi: "Times New Roman", cs: "Times New Roman", eastAsia: "SimSun" };

const run = (text, options = {}) => new TextRun({ text, font, size: 22, ...options });
const para = (children, options = {}) =>
  new Paragraph({ spacing: { after: 140, line: 300 }, ...options, children: Array.isArray(children) ? children : [children] });
const p = (text) => para(run(text), { indent: { firstLine: convertInchesToTwip(0.33) } });
const h1 = (text) =>
  para(run(text, { bold: true, size: 28, color: palette.dark }), { heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 140 } });
const h2 = (text) =>
  para(run(text, { bold: true, size: 24, color: palette.primary }), { heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 110 } });
const bullet = (text) => para(run(text), { numbering: { reference: "bullets", level: 0 } });

const cell = (content, options = {}) =>
  new TableCell({
    children: Array.isArray(content) ? content : [para(run(content, { size: 20 }), { spacing: { after: 30, line: 250 } })],
    margins: { top: 90, bottom: 90, left: 90, right: 90 },
    ...options,
  });
const hcell = (text, width) =>
  cell([para(run(text, { bold: true, size: 20 }), { spacing: { after: 30, line: 250 } })], {
    shading: { type: ShadingType.CLEAR, fill: palette.fill },
    width: { size: width, type: WidthType.DXA },
  });
const bcell = (text, width) => cell([para(run(text, { size: 20 }), { spacing: { after: 30, line: 250 } })], { width: { size: width, type: WidthType.DXA } });
const makeTable = (widths, headerRow, bodyRows) =>
  new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ children: headerRow.map((t, i) => hcell(t, widths[i])) }),
      ...bodyRows.map((r) => new TableRow({ children: r.map((c, i) => bcell(c, widths[i])) })),
    ],
  });

// 字段设计表
const fieldTable = makeTable(
  [1500, 1300, 6200],
  ["字段", "必填", "填写说明"],
  [
    ["id", "必填", "编号规则：EVAL-<文书类>-<序号>，如 EVAL-FL-001（FL=法律类）"],
    ["query（问题）", "必填", "真实业务口吻的中文问题；禁止从原文整句照搬，避免词面泄漏"],
    ["reference（标准答案）", "必填", "信息点式答案（列要点），不追求文学通顺；真实性、准确性为“应”（强制，GB/Z 236—2026 §8.3）"],
    ["source_locator（答案出处）", "必填", "文档名＋章节/条款路径＋关键原文片段（100 字内）；检索判分就以此为准"],
    ["answer_span_type", "必填", "单切片可答 / 需多切片合成 —— 决定用命中率/MRR 还是召回率/精确度口径"],
    ["scenario（场景）", "必填", "按实际业务问题类型填：事实问答 / 条款引用 / 汇总归纳 / 多文档综合 / 计算类"],
    ["difficulty（难度）", "必填", "易（原文直接可查）/ 中（需同义改写匹配）/ 难（需跨段推理）"],
    ["noise_type", "条件必填", "仅噪声/拒答测试题填：混入无关切片 / 概念近似干扰 / 库中无答案（应拒答）"],
    ["annotator / reviewer", "必填", "标注人、复核人分离；分歧由第三人仲裁"],
    ["version", "必填", "所属评测集版本号（如 v1.0）；评测集冻结后只增不改，改即升版"],
  ],
);

// 覆盖维度配比表
const coverTable = makeTable(
  [2200, 3400, 3400],
  ["维度", "取值", "建议配比（100 条规模）"],
  [
    ["文书类型", "按实际入库结构：法规 / 合同 / 公文 / 手册 / 其他", "按库内文档占比分配，任一类型 ≥10 条"],
    ["场景", "事实问答 / 条款引用 / 汇总归纳 / 多文档综合", "事实问答约 40%，其余各约 20%"],
    ["难度", "易 / 中 / 难", "3 : 5 : 2"],
    ["答案跨度", "单切片可答 / 需多切片合成", "约 7 : 3"],
    ["噪声与拒答", "正常题 / 噪声题 / 应拒答题", "正常约 80%，噪声＋拒答约 20%（B5 验证必需）"],
  ],
);

// B1–B7 对接表
const mapTable = makeTable(
  [700, 2900, 5400],
  ["基线", "验证内容", "评测集需具备的条件"],
  [
    ["B1", "chunk size 128/256/512 三档复测", "通用问答即可；题目在三种切分配置下共用同一份评测集"],
    ["B2", "overlap 0%/10%/20% 三档 A/B", "同 B1；判分用“前 K 含答案文本”，需 source_locator 可定位"],
    ["B3", "embedding 模型对照", "含足够“中/难”题（同义改写匹配），才能拉开模型差异"],
    ["B4", "法规双粒度切分", "法规类文书题目 ≥15 条，且 source_locator 精确到“第×条”"],
    ["B5", "噪声鲁棒与拒答", "噪声题＋应拒答题合计约 20 条；应拒答题答案必须确不在库中"],
    ["B6", "金融文书检索器排序", "金融类文书题目 ≥10 条（无金融类则跳过此项）"],
    ["B7", "稀疏 vs 稠密对照", "含关键词精确匹配型问题（条款号、专有名词）约 15 条"],
  ],
);

// 空白记录示例表
const sampleTable = makeTable(
  [1100, 2300, 2300, 1900, 1400],
  ["id", "query", "reference（要点）", "source_locator", "场景/难度/跨度"],
  [
    ["EVAL-FL-001", "（示例）××办法规定的注册资本认缴期限是多久？", "（示例）要点 1：…；要点 2：…", "《××办法》第三章第 12 条；“……”", "条款引用 / 易 / 单切片"],
    ["EVAL-HT-014", "", "", "", ""],
    ["EVAL-GW-027", "", "", "", ""],
  ],
);

const children = [
  para(run(T`RAG 知识库自建基线评测集建集模板`, { bold: true, size: 34, color: palette.dark }), {
    heading: HeadingLevel.TITLE,
    alignment: AlignmentType.CENTER,
    spacing: { before: 480, after: 240 },
  }),
  para(run(T`V1.0 · 2026-11-17 · 配套《RAG 知识库切片质量标准与验收清单 V3.2（修订草案）》附录 A 使用`, { color: palette.primary }), {
    alignment: AlignmentType.CENTER,
    spacing: { after: 480 },
  }),

  h1("一、用途与定位"),
  p(T`本模板用于建设本知识库的自建基线评测集，承接 V3.2 附录 A“中文数值基线候选清单”B1–B7 的验证动作：清单中的候选数值全部标注【待验证】，验证通过并留档后方可回填为本库实测基线。建集口径对齐 GB/Z 236—2026 §8.3：参考答案的真实性、准确性为“应”（强制）；场景/题型/难度覆盖、偏差分析、构建流程记录、动态更新机制为“宜”（推荐）。`),
  p(T`三条铁律：（1）评测集固定后才允许调参——没有冻结的评测集，任何指标对比都无效；（2）单变量原则——策略对比实验固定检索与重排参数，只动切分；（3）版本管理——评测集冻结后只增不改，修改即升版并留档。`),

  h1("二、记录字段设计"),
  p(T`每条评测记录占一行，建议用电子表格维护（每行一题）。字段如下：`),
  fieldTable,

  h1("三、标注口径"),
  bullet(T`判据统一为“前 K 个召回切片含正确答案文本”（以 source_locator 的原文片段比对为准），不是“有召回就算过”；`),
  bullet(T`文本比对前做归一化：空白、换行、全半角统一；PDF 来源切片尤其应如此；`),
  bullet(T`单切片可答题用命中率/MRR 口径；需多切片合成题才做召回率/精确度全标注；`),
  bullet(T`噪声题与应拒答题单独标记（noise_type），不计入常规命中率，单独算拒答正确率；`),
  bullet(T`两人独立标注＋分歧第三人仲裁；仲裁记录留档。`),

  h1("四、覆盖维度与配比"),
  p(T`规模 50–100 条起步（C 级经验值）；下表按 100 条给配比，50 条规模同比缩减：`),
  coverTable,

  h1("五、构建流程（六步）"),
  bullet(T`1. 抽文档：从实际入库文档中分层抽样（覆盖各文书类型），记录抽样清单；`),
  bullet(T`2. 出题：按覆盖矩阵配额出题，问题用业务口吻改写，禁止整句照搬原文；`),
  bullet(T`3. 标注：两人独立填 reference 与 source_locator，标注答案跨度类型；`),
  bullet(T`4. 复核与仲裁：交叉复核，分歧仲裁，记录一致性情况；`),
  bullet(T`5. 偏差分析：检查场景/难度/文书类型是否失衡，补足短板项；`),
  bullet(T`6. 冻结入库：定版号（如 v1.0）、留档（构建记录＋抽样清单＋仲裁记录），之后只增不改。`),

  h1("六、与附录 A 基线验证动作（B1–B7）的对接"),
  p(T`建集时按下表检查评测集是否满足各验证项的前置条件；不满足的验证项应暂缓，不得用不合格数据强行验证：`),
  mapTable,

  h1("七、评测记录空白模板（示例一行）"),
  sampleTable,
  para(run(T`说明：电子表格落地时以第二节全部字段为列；本表仅示核心五列的填写样式。`, { italics: true, color: palette.light, size: 18 })),

  h1("八、回填与留档"),
  bullet(T`每项基线验证完成后，将实测值回填 V3.2 附录 A 对应行，状态由【待验证】改为【已验证 vX.X】，并注明评测集版本、平台、embedding 模型、日期；`),
  bullet(T`验证记录至少包含：切分策略、chunk_size、overlap、检索/重排参数、各指标数值、评测集版本；`),
  bullet(T`知识库内容更新后应重跑评测集，防止旧知识污染新回答（GB/Z 236—2026 §8.3 f) 将其列为“宜”）；评测集本身按季度增补新题并升版。`),
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
            style: { paragraph: { indent: { left: convertInchesToTwip(0.4), hanging: convertInchesToTwip(0.2) } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      headers: {
        default: new Header({
          children: [para(run(T`RAG 自建基线评测集建集模板 V1.0`, { bold: true, color: palette.primary, size: 18 }), { alignment: AlignmentType.CENTER })],
        }),
      },
      footers: {
        default: new Footer({ children: [para(new TextRun({ children: [PageNumber.CURRENT], font }), { alignment: AlignmentType.CENTER })] }),
      },
      children,
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync(outputPath, buffer);
