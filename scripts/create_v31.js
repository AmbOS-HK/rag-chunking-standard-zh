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
if (!outputPath) throw new Error("Usage: node create_v31.js <输出文件.docx（绝对或相对路径）>");

const T = String.raw;
const title = T`RAG 知识库切片质量标准与验收清单 V3.1（修订草案）`;

const palette = {
  dark: "263238",
  primary: "37474F",
  light: "78909C",
  fill: "EEF3F6",
  warn: "FDF3E7",
  added: "8C6D3F",
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
    spacing: { after: 140, line: 300 },
    ...options,
    children: Array.isArray(children) ? children : [children],
  });
const p = (text) => para(run(text), { indent: { firstLine: convertInchesToTwip(0.33) } });
const pAdd = (text, refs) =>
  para(
    [run(T`【V3.1 新增/修订】`, { bold: true, color: palette.added }), run(text, { color: palette.added }), ...(refs || [])],
    { indent: { firstLine: convertInchesToTwip(0.33) } },
  );
const h1 = (text) =>
  para(run(text, { bold: true, size: 30, color: palette.dark }), {
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 140 },
  });
const h2 = (text) =>
  para(run(text, { bold: true, size: 26, color: palette.primary }), {
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
  });
const bullet = (children, options = {}) =>
  para(Array.isArray(children) ? children : [run(children)], {
    numbering: { reference: "bullets", level: 0 },
    ...options,
  });
const bulletAdd = (text, refs) =>
  bullet([run(T`【V3.1 修订】`, { bold: true, color: palette.added }), run(text, { color: palette.added }), ...(refs || [])]);
const check = (text, added) =>
  para(
    [run("☐ ", { bold: true }), ...(added ? [run(text, { color: palette.added })] : [run(text)])],
    { indent: { left: convertInchesToTwip(0.2) } },
  );

const cell = (content, options = {}) =>
  new TableCell({
    children: Array.isArray(content) ? content : [para(run(content), { spacing: { after: 30, line: 260 } })],
    margins: { top: 90, bottom: 90, left: 90, right: 90 },
    ...options,
  });
const hcell = (text, width, fill = palette.fill) =>
  cell([para(run(text, { bold: true, size: 21 }), { spacing: { after: 30, line: 250 } })], {
    shading: { type: ShadingType.CLEAR, fill },
    width: { size: width, type: WidthType.DXA },
  });
const bcell = (text, width, opts = {}) =>
  cell([para(run(text, { size: 21, ...(opts.added ? { color: palette.added } : {}) }), { spacing: { after: 30, line: 250 } })], {
    width: { size: width, type: WidthType.DXA },
    ...(opts.fill ? { shading: { type: ShadingType.CLEAR, fill: opts.fill } } : {}),
  });
const makeTable = (widths, headerRow, bodyRows) =>
  new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ children: headerRow.map((t, i) => hcell(t, widths[i])) }),
      ...bodyRows.map((r) => new TableRow({ children: r.map((c, i) => (typeof c === "object" && c !== null && c.text !== undefined ? bcell(c.text, widths[i], c) : bcell(c, widths[i]))) })),
    ],
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

// ============ 表格数据 ============

// 修订对照表
const revTable = makeTable(
  [700, 1500, 3700, 3100],
  ["序号", "章节", "修改内容", "理由与依据"],
  [
    ["R1", "§六 尺寸与重叠", "overlap 条款重写：由“默认不设 overlap（o=0）”改为“默认 0 起步 + 目标平台稠密链路 A/B 复验（建议扫描 0%/10%/20%）后固化”；补入三方证据对照表", "2601.14123（英文/稀疏/端到端：无收益）与 CRUD-RAG（中文/稠密：QA 关键信息有收益）、arXiv:2604.12047（金融 PDF/检索层：25% 最优）结论冲突；三个目标平台默认均为稠密检索", { added: true }],
    ["R2", "§六 单片大小", "注明 300–800 字为中文场景基线而非上限；新增各平台默认值对照表", "百炼默认 600 token、千帆默认 1200 字、盘古助手默认 300、AgentArts 默认 500；原区间易被误读为硬上限", { added: true }],
    ["R3", "§三 构成规则", "新增规则 4：平台不暴露某成分/参数的配置能力时，应登记为“平台约束”并留档，不判为不达标", "盘古助手等平台无标题路径/元数据/overlap 配置项，原条款无执行对象", { added: true }],
    ["R4", "§五 特殊内容处置", "表格条目增加平台例外：平台固定按数据行切分+表头拼接的（如百炼），可登记为等效合规路径，最终以检索验收判效", "百炼表格行级切分为平台固定行为，与本条目的相同（防语义混淆）、路径不同", { added: true }],
    ["R5", "§十 处置路径", "“变更后指标无变化”一行补充：百炼类平台切片方式与元数据配置创建后不可更改，“升档重导”等于新建知识库并全量重新导入", "平台约束导致回归成本高于原文措辞暗示", { added: true }],
    ["R6", "参考依据", "补入 CRUD-RAG（arXiv:2401.17043，中文场景主锚点，C 级）、arXiv:2604.12047（检索层反例，C 级）、Legal-DC（arXiv:2603.11772，中文法律文书基准，C 级预印本）与 HOPE（ACM SIGIR 2025 切片评估框架，C 级、ACM 正式收录）；§五结构感知切分条目同步补注 Legal-DC 证据", "证据链闭合到中文+稠密检索的真实基座；中文法律文书场景获得直接实验支撑", { added: true }],
    ["R7", "第二部分清单", "同步修改：overlap 检查项改为“默认 0 起步 + A/B 复验”；新增“平台约束已登记留档”检查项；表格处置项增加平台例外", "与第一部分保持一致", { added: true }],
    ["R8", "第一部分 范围", "新增定位声明：本标准面向中文语料场景制定；英文或多语言语料的数值基线应另行校准，不应直接套用本标准的中文经验值", "评审确认标准初衷即面向中文场景；英文实验数值（如 150–300 tokens、2.5k 拐点）不可平移，在正文钉死定位防止误用", { added: true }],
    ["R9", "附录 A（新增）", "新增资料性附录“中文数值基线候选清单（全部待验证）”：7 项候选基线（chunk size 区间、overlap 区间、embedding 对照、法律双粒度、噪声拒答、金融检索器排序、稀疏/稠密对照），全部标注【待验证】", "中文数值基线无本库一手实测；以 ACM TOIS（CRUD-RAG）为主干汇总候选起点，供后期专门验证后回填", { added: true }],
    ["R10", "参考依据/全文引用", "上位依据更新：GB/Z 征求意见稿 → 正式发布版 GB/Z 236—2026《人工智能 检索增强生成系统评价指标与方法》；全文引用同步，§8.3 动态更新条款引用由 e) 更正为 f)", "该标准已于 2026-08-27 正式发布（国家标准公告 2026 年第 37 号），名称与征求意见稿不同；评审后勘误", { added: true }],
  ].map((r) => [r[0], r[1], { text: r[2], added: true }, { text: r[3], added: true }]),
);

// §一 来源分级表
const srcTable = makeTable(
  [700, 4600, 3700],
  ["级别", "来源", "效力说明"],
  [
    ["A", "ITU-T F.748.52《Requirements and evaluation methods for retrieval augmented generation of large scale pre-trained models》（2025-03-16 批准，现行有效，中国信通院牵头制定）", "现行国际标准；定性要求＋公式化评估方法，不设数值阈值"],
    ["B", "GB/Z 236—2026《人工智能 检索增强生成系统评价指标与方法》（2026-08-27 发布，国家标准公告 2026 年第 37 号）", "确立“端到端评价＋模块化评价”框架与指标公式，不设数值阈值；已正式发布"],
    ["C", "行业通行做法与公开研究/benchmark（LangChain/LlamaIndex 生态、RAGAS/ARES/TruLens 评估体系、arXiv 公开论文、工程实践）", "经验值，非强制；不同业务场景容忍度不同，应先建基线再看趋势"],
  ],
);

// §三 切片构成表
const compTable = makeTable(
  [1700, 1700, 2600, 3000],
  ["成分", "必填性", "定义", "说明与示例"],
  [
    ["正文（content）", "必填", "切片承载的语义完整文本单元", "句读收尾、无腰斩；是检索与生成的直接对象"],
    ["来源文档（source_doc）", "必填", "切片出处文档的标识", "文档名或文档 ID；检索结果可溯源的前提"],
    ["标题路径（title_path）", "必填（无标题结构文档可置空，但应登记原因）", "切片所属章节标题的层级路径", "示例：“第三章 §3.2”；质量要求见§八"],
    ["元数据（metadata）", "必填（字段集合按库定义）", "支持过滤检索的属性字段", "文档类型、时间、地域/主题、关键词等；质量口径见§八"],
    ["摘要（summary）", "可选", "切片内容的概括", "长切片或弱结构件宜携带"],
    ["实现层字段（chunk_id、向量、版本号等）", "不在本规范范围", "索引与检索实现所需的技术字段", "由平台/自研索引自定，本规范不约束"],
  ],
);

// §四 质量维度表
const dimTable = makeTable(
  [1500, 3600, 3900],
  ["维度", "核心要求", "不合格的典型表现"],
  [
    ["语义完整性", "一片一个完整语义单元", "条文/段落拦腰切断、表格断行、半句话收尾"],
    ["粒度适当", "不碎不大，长度分布合理", "过碎丢上下文、过大稀释向量精度并挤占召回名额"],
    ["结构对齐", "边界锚在文档天然结构（章节、条款、段落、页）", "固定长度硬截、结构锚点失效退化为硬切"],
    ["自包含性", "切片脱离原文可独立理解、可溯源", "不知出自哪份文件、哪个章节"],
    ["特殊内容适配", "表格/PDF/代码/弱结构长文分别处置", "表行级切断、扫描件零切片、代码块破碎"],
  ],
);

// §五 策略选型表
const stratTable = makeTable(
  [2300, 2600, 4100],
  ["策略", "适用文档", "说明"],
  [
    ["结构感知切分（按标题层级/条款/章节，多级回退）", "有明确结构的文档（法规、标准、手册、Markdown）", { text: "单源公开实践中覆盖率最高、命中率最稳的一类（C 级）；中文法律文书基准 Legal-DC（arXiv:2603.11772）实验表明：条文级结构切分与语义 chunk 切分并行的多粒度方案优于任一单一粒度（C 级，预印本）", added: true }],
    ["递归字符/句子切分（段落→句子逐级回退）", "一般自然语言文档", "通用保底策略，业界默认基线；arXiv:2601.14123 实测梯度：句子≈语义>定长 token≫代码（文本任务，C 级）"],
    ["语义切分（按 embedding 相似度骤降点）", "无标题纯长文（合同、论文）或超大上下文场景", "收益有限：句子切分在约 5k tokens 内与语义切分效果持平（arXiv:2601.14123）；另有工程博客报告其计算成本约为递归切分的 50 倍（C 级）"],
    ["固定长度切分", "仅作对照基线", "产生硬截断，不宜单独用于生产"],
    ["父子切分（Parent-Child）", "精度与上下文矛盾突出的场景", "子片负责检索命中、父片负责提供上下文，生产级常见配置（C 级）"],
  ],
);

// §六 平台默认值对照表（新增）
const platTable = makeTable(
  [1800, 2200, 2200, 2800],
  ["平台", "默认分段长度", "单位", "备注"],
  [
    [{ text: "阿里云百炼", added: true }, { text: "600（范围 10–6000）", added: true }, { text: "token", added: true }, { text: "官方另建议：FAQ 256、长篇报告 2048、法律合同按标题 512", added: true }],
    [{ text: "百度千帆", added: true }, { text: "1200（上限 15 万字）", added: true }, { text: "字符", added: true }, { text: "整文件切片单切片最高 15 万字，属全文总结场景", added: true }],
    [{ text: "CodeArts 盘古助手", added: true }, { text: "300（范围 50–1000）", added: true }, { text: "未明示，入库前应确认", added: true }, { text: "只支持中文切分器", added: true }],
    [{ text: "华为 AgentArts 系", added: true }, { text: "500（范围 1–6000）", added: true }, { text: "以平台界面为准", added: true }, { text: "支持自动/长度/层级分段", added: true }],
  ],
);

// §六 overlap 三方证据对照表（新增）
const overlapTable = makeTable(
  [2200, 2000, 2300, 2500],
  ["研究", "语料/检索器", "测量层次", "overlap 结论"],
  [
    [{ text: "arXiv:2601.14123（2026-01，ECIR 短篇）", added: true }, { text: "英文 NQ；SPLADE 稀疏检索", added: true }, { text: "端到端生成质量（BERTScore/EM）", added: true }, { text: "10–20% 无可测量收益（|ΔBERTScore|≤0.004）；索引按 1/(1−o) 膨胀", added: true }],
    [{ text: "CRUD-RAG（arXiv:2401.17043，ACM 收录）", added: true }, { text: "中文新闻；稠密检索（bge 等）", added: true }, { text: "关键信息指标（RAGQuestEval）", added: true }, { text: "单文档 QA 有正向收益：精确度 52.67→54.06、召回 70.82→73.94（0%→70%）", added: true }],
    [{ text: "arXiv:2604.12047（ICSE-SEIP 2026）", added: true }, { text: "英文金融 PDF；混合检索", added: true }, { text: "检索层（P@1/R@3/MRR）", added: true }, { text: "25% 最优（MRR 0.529→0.658），0% 最差，50% 回落", added: true }],
  ],
);

// §九 检索指标表
const metricTable = makeTable(
  [2400, 4200, 2400],
  ["指标", "合格线", "来源级别"],
  [
    ["Hit Rate@5", "≥ 0.85（@5 截断为行业用法）", "C"],
    ["Recall@5", "≥ 0.90（生产口径）；宽松基线 0.75", "C"],
    ["MRR", "≥ 0.75", "C"],
  ],
);

// 附录 A 中文数值基线候选清单（全部待验证）
const baselineTable = makeTable(
  [600, 2500, 2100, 1900, 1900],
  ["编号", "候选基线（数值/区间）", "来源与载体（标准分级均 C 级）", "适用场景", "验证动作与状态"],
  [
    [{ text: "B1", added: true }, { text: "中文 chunk size 候选区间 64–512（token 计，稠密检索系统扫描）；任务相关：单文档问答偏好小切片（128 优于 512），文本续写/多文档问答偏好大切片", added: true }, { text: "CRUD-RAG（ACM TOIS 43(2)，2025；近 30 万篇中文新闻）", added: true }, { text: "通用中文问答/写作类文书", added: true }, { text: "本库文书以 128/256/512 三档复测 HR@5/MRR【待验证】", added: true }],
    [{ text: "B2", added: true }, { text: "overlap 候选区间 0%–70%；单文档 QA 关键信息精确度 52.67→54.06、召回 70.82→73.94（0%→70%）", added: true }, { text: "CRUD-RAG（同上）", added: true }, { text: "问答为主场景", added: true }, { text: "0%/10%/20% 三档 A/B 复验【待验证】", added: true }],
    [{ text: "B3", added: true }, { text: "中文稠密 embedding 系统对照（bge-base 等四款）", added: true }, { text: "CRUD-RAG（同上）", added: true }, { text: "全部文书", added: true }, { text: "换 embedding 模型全量复测评测集【待验证】", added: true }],
    [{ text: "B4", added: true }, { text: "法律文书双粒度切分：条文级结构切分＋语义 chunk 并行，优于单一粒度；Recall 78.02%、MRR@5 50.23%（最强基线 75.86%/47.56%）", added: true }, { text: "Legal-DC（arXiv:2603.11772 预印本；480 份中文法律法规＋2,475 条 QA）", added: true }, { text: "法规、合同等条文类文书", added: true }, { text: "法规类文书启用双粒度索引后复测【待验证】", added: true }],
    [{ text: "B5", added: true }, { text: "噪声鲁棒与拒答：噪声环境下中文 LLM 最高拒答率仅 43.33%（英文 45%，2023 年模型实测）；评测集配比可参考其 600 基础＋200 整合＋200 反事实结构", added: true }, { text: "RGB（AAAI 2024，非 ACM；中英各半语料）", added: true }, { text: "端到端进阶维度", added: true }, { text: "噪声切片混入＋拒答能力评测【待验证】", added: true }],
    [{ text: "B6", added: true }, { text: "中文金融场景检索器排序对照：Jina-zh / BGE-large-zh / BGE-m3 / GTE-Qwen2-1.5B 的 MAP、MRR 对比", added: true }, { text: "人大金融领域中文 RAG 基准（2025-01，机构发布，非 ACM）", added: true }, { text: "金融类文书", added: true }, { text: "含金融文书时按其排序做选型初值再实测【待验证】", added: true }],
    [{ text: "B7", added: true }, { text: "稀疏 vs 稠密对照：BM25 与 BGE-base-zh-v1.5 检索参考下的端到端差异（EM/F1/ROUGE-L）", added: true }, { text: "DomainRAG（arXiv:2406.05654 预印本；中文领域文档）", added: true }, { text: "领域文书", added: true }, { text: "本库跑一次稀疏/稠密对照实验【待验证】", added: true }],
  ],
);

// §十 / 清单§五 处置路径表
const fixTable = makeTable(
  [1700, 3600, 3700],
  ["症状", "先查", "再查"],
  [
    ["截断检测 >0", "该文档切分策略是否选错（结构锚点失效 → 退硬切）", "chunk_size 是否太小（升一档重导验证）"],
    ["命中率低（<70%）", "先查检索侧：数据是否入库、切片完整、embedding 适配", "命中率未达标前不应调生成端（提示词/模型）——行业共识的排查顺序（C 级）"],
    ["命中率高但答案错", "生成侧问题：提示词、上下文矛盾、忠实度；或上下文预算 C 过大触发“上下文悬崖”", "与切片无关，走生成模块/检索预算优化"],
    ["元数据成批抽空", "抽取配置（提示词过长/规则过多会压垮抽取模型）", "简化抽取提示词，注意平台缓存行为"],
    [{ text: "变更后指标无变化", added: true }, { text: "配置是否真正生效（草稿/发布、缓存、版本）", added: true }, { text: "以实际输出为准回归验证。注意：百炼类平台切片方式与元数据配置创建后不可更改，“升档重导”等于新建知识库并全量重新导入，回归成本应按此预估", added: true }],
  ],
);

const tocEntries = [
  { title: "修订说明与对照表", level: 1, page: 2 },
  { title: "第一部分 RAG 知识库切片质量标准", level: 1, page: 3 },
  { title: "一、依据与来源分级", level: 2, page: 3 },
  { title: "二、术语", level: 2, page: 4 },
  { title: "三、切片构成规范", level: 2, page: 4 },
  { title: "四、质量维度", level: 2, page: 5 },
  { title: "五、切分策略选型标准", level: 2, page: 5 },
  { title: "六、尺寸与重叠标准", level: 2, page: 6 },
  { title: "七、完整性检测标准", level: 2, page: 7 },
  { title: "八、自包含与元数据标准", level: 2, page: 7 },
  { title: "九、检索验收标准（最终判据）", level: 2, page: 8 },
  { title: "十、不达标处置路径", level: 2, page: 9 },
  { title: "十一、质量闭环与回归纪律", level: 2, page: 9 },
  { title: "参考依据", level: 2, page: 9 },
  { title: "附录 A（资料性）中文数值基线候选清单（全部待验证）", level: 2, page: 10 },
  { title: "第二部分 RAG 知识库切片验收清单", level: 1, page: 10 },
  { title: "一、切分策略选型（切之前）", level: 2, page: 10 },
  { title: "二、切片结果核查（切之后）", level: 2, page: 11 },
  { title: "三、自包含与元数据", level: 2, page: 11 },
  { title: "四、检索验收（最终判据）", level: 2, page: 12 },
  { title: "五、不达标处置路径", level: 2, page: 12 },
  { title: "六、收尾与回归纪律", level: 2, page: 12 },
];

const children = [
  // 封面
  para(run(title, { bold: true, size: 34, color: palette.dark }), {
    heading: HeadingLevel.TITLE,
    alignment: AlignmentType.CENTER,
    spacing: { before: 480, after: 240 },
  }),
  para(run("通用版 · V3.1 修订草案 · 2026-11-17", { color: palette.primary, bold: true }), {
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
  }),
  para(run("基线版本：V3.0（2026-11-17，正式评审版）　　修订依据：《平台冲突核查与文献适用性评审意见（修订版 V2）》", { color: palette.primary }), {
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
  }),
  para(run("凡例：棕褐色文字为 V3.1 新增/修订内容，其余与 V3.0 相同。", { color: palette.added }), {
    alignment: AlignmentType.CENTER,
    spacing: { after: 480 },
  }),
  para(run("目录", { bold: true, size: 30, color: palette.dark }), { heading: HeadingLevel.HEADING_1 }),
  para(run("页码为估计值，可在 Word/WPS 中右键目录选择“更新域”刷新。", { italics: true, color: palette.light })),
  toc(tocEntries),

  // 修订说明
  h1("修订说明与对照表"),
  p(T`V3.1 修订记录：依据外部评审意见，（1）§六 overlap 条款由“默认不设”改为“默认 0 起步＋目标平台稠密链路 A/B 复验后固化”，补入三方证据对照；（2）§六单片大小区间注明为基线非上限，新增平台默认值对照；（3）§三新增平台约束登记规则；（4）§五表格处置增加平台例外；（5）§十补充百炼类平台重导成本提示；（6）参考依据补入 CRUD-RAG、arXiv:2604.12047、Legal-DC（中文法律文书基准，预印本）与 HOPE（ACM SIGIR 2025），§五结构感知切分条目同步补注 Legal-DC 证据；（7）第二部分清单同步修改；（8）第一部分范围段新增定位声明：本标准面向中文语料场景制定，英文或多语言语料的数值基线应另行校准；（9）新增附录 A“中文数值基线候选清单（全部待验证）”，以 ACM TOIS 收录的 CRUD-RAG 为主干汇总 7 项候选基线，供后期专门验证后回填；（10）上位依据更新：GB/Z 征求意见稿已发布为 GB/Z 236—2026（2026-08-27），全文引用同步更新，§8.3 子项引用 e)→f) 更正。既有修订记录（V2.1 勘误网传“信通院量化口径”；V2.2–V3.0 版本对齐等）保持不变。`),
  revTable,

  // ================= 第一部分 =================
  h1("第一部分 RAG 知识库切片质量标准"),
  p(T`版本：V3.1（2026-11-17，修订草案，与《RAG 切片验收清单》版本对齐）。性质：厂商中立的通用标准，不绑定任何具体平台、产品或项目。标准用语参照 GB/T 1.1—2020 的“应/不应/宜/不宜/可”体系。`),
  p(T`范围：本文件规定 RAG 知识库切片的逻辑构成、质量维度、切分策略选型、尺寸与重叠、完整性检测、自包含与元数据、检索验收的判定准则及不达标处置路径，适用于 RAG 知识库建设中的切片质量管理与验收。本文件不规定切片的物理存储结构与实现层字段，不设强制性数值阈值（数值合格线均为 C 级经验值，使用方宜按场景自建基线）。`),
  pAdd(T`定位声明：本标准面向中文语料场景制定。英文或多语言语料的数值基线应另行校准，不应直接套用本标准的中文经验值；英文/通用场景文献的结论仅作方向性参考，数值不可平移。`),

  h2("一、依据与来源分级"),
  p(T`本标准条款按以下来源分级引用，使用时请注意效力层级：`),
  srcTable,
  p(T`两级标准均不设数值阈值——阈值是留给使用方按场景自定的。本标准中的所有数值合格线一律标为 C 级经验值。`),
  p(T`勘误说明（V2.1 起，V2.7 修订措辞）：网络流传的“信通院《检索增强生成技术要求与评估方法》25 能力项＋量化阈值”版本（元数据完整性 ≥90%、MRR≥0.8、幻觉率 <15% 等），经与 ITU-T F.748.52 标准原文逐字核验不符；该版本是否另有所本（如信通院内部评估文件）未经核验，不下断言。本标准不采纳其外部标准出处主张；其中个别数值（元数据完整性/准确性、幻觉率红线）经评估保留为本标准自定经验起点，已在相应条款明示“无外部标准出处”。`),
  pAdd(T`V3.1 注：截至 2026-11，网传“25 能力项＋量化阈值”版本仍在公开渠道传播，本标准立场不变——不采纳其外部标准出处主张，相关数值维持自定经验起点地位。`),

  h2("二、术语"),
  p(T`切片（Chunk）：文档经切分后进入索引的最小检索单元。`),
  p(T`单片大小（S）与上下文预算（C）：S 是单个切片的长度；C 是检索后送入生成模型的上下文总 token 数。两者是不同的调参维度，不可混用——“上下文悬崖”等结论针对的是 C。`),
  p(T`召回率（Recall）：检索返回的相关文档占全部相关文档的比例（公式定义见 GB/Z 236—2026公式 (5)、ITU-T F.748.52 §8.1）。`),
  p(T`命中率（Hit Rate, HR）：检索模块在所有查询中至少返回一个相关结果的比例（GB/Z 236—2026公式 (3)；“@K”截断为行业用法，C 级）。`),
  p(T`MRR（平均倒数排名）：首个相关结果排名倒数的均值（GB/Z 236—2026公式 (6)）。`),
  p(T`NDCG（归一化折损累积增益）：Top-K 结果列表的排序质量（GB/Z 236—2026公式 (7)(8)）。`),
  p(T`上下文忠实度（Faithfulness）：生成答案中能从检索上下文推断出的信息点占比（GB/Z 236—2026公式 (10)，信息点交叉检查法）；其反面为幻觉率。ITU-T F.748.52 附录 A 另对答案相关性、忠实度、正确性给出 5 级分级要求（定性）。`),

  h2("三、切片构成规范"),
  p(T`本节规定切片的逻辑构成（必备成分与必填性），不规定物理存储形式（平台内嵌头、元数据字段、拼接到正文等载体由实现自定）——与上位标准“定要求、不定结构”的写法保持一致。`),
  compTable,
  p(T`构成规则：`),
  bullet(T`1. 必填成分缺失即判不合格（清单 §三对应检查项）；`),
  bullet(T`2. 正文是唯一检索与生成对象，其余成分为溯源、过滤与上下文服务——成分增减不得损害正文语义完整性；`),
  bullet(T`3. 各成分的载体形式（内嵌、外挂、拼接）由实现决定，但应保证检索结果可回显来源与标题路径。`),
  bulletAdd(T`4. 平台约束登记：当所用平台不暴露某成分或参数的配置能力（如无标题路径/元数据/overlap 配置项）时，应将该限制登记为“平台约束”并留档，不判为不达标；平台升级暴露该能力后应补办配置并复核。`),

  h2("四、质量维度"),
  dimTable,

  h2("五、切分策略选型标准"),
  p(T`ITU-T F.748.52 §7.2（数据分段）要求支持多种分段策略（定长、递归、标点、字符数等）、可配置（自定义策略、最大分段长度、最大重叠长度）、多数据类型（文本、文档、表格数据）——选型标准与此对齐，按文档结构选型，不一刀切：`),
  stratTable,
  p(T`特殊内容处置：`),
  bullet(T`表格：整块成切分单元，防表行级切断；超大表格按逻辑分组拆分并保留表头；`),
  bulletAdd(T`表格平台例外：部分平台对表格文件固定按数据行切分并将表头拼接到每行切片（如阿里云百炼）——目的与本条相同（防语义混淆），路径不同，可登记为等效合规路径，最终以检索验收结果判效。`),
  bullet(T`PDF/扫描件：先做版面还原再按语义段落切，注意平台间解析能力差异，“本地判无文本层”不等于平台切不出；`),
  bullet(T`代码：按函数/类等语法单元切，不打断语句块。`),
  p(T`选型纪律：`),
  bullet(T`切片方式以实际切片输出为唯一判定依据，不以配置界面/接口回显为准；`),
  bullet(T`使用正则/分隔符切分前，应验证锚点在目标文档正文中真实可见（经解析器处理后的文本），锚点失效会静默退化为按长度硬切。`),

  h2("六、尺寸与重叠标准"),
  p(T`单片大小（S）：`),
  bullet(T`先确认单位：不同平台 chunk_size 单位可能是字符也可能是 token，调参前应先确认；`),
  bullet(T`经验区间：中文通用场景单片 300–800 字（C 级）；arXiv:2601.14123 的默认配方为 150–300 tokens（英文场景，单位不同，不应直接换算套用）；法规条款类等短语义单元顺其自然，不强凑长度；`),
  bulletAdd(T`区间性质：300–800 字为中文场景基线而非上限。各平台默认值对照如下（入库前应核实当期平台文档）：`),
  platTable,
  bullet(T`软上限意识：多数智能切分的 chunk_size 是软上限不是硬截断，允许单片小幅超出（10%–20% 幅度属常见）；`),
  bullet(T`调档纪律：出现硬切时优先升一档验证，够用即停——更大的上限只摊薄检索粒度；`),
  bullet(T`分布检查：看长度分布的 min/max 而非均值——min 个位数说明有碎片，max 远超上限说明切分器失效。`),
  pAdd(T`重叠（overlap）——V3.1 重写。证据现状（三方对照，均 C 级）：`, [new FootnoteReferenceRun(1), new FootnoteReferenceRun(2), new FootnoteReferenceRun(3)]),
  overlapTable,
  bulletAdd(T`结论判定：三方证据冲突——overlap 的收益与检索器类型（稀疏/稠密）、语言、测量层次（端到端/检索层）相关，不存在可平移的通用默认值。索引体积按 1/(1−o) 膨胀为算术事实，与场景无关（o=20% 即 1.25 倍）。`),
  bulletAdd(T`推荐做法：默认从 0 起步；在目标平台的稠密检索链路上用固定评测集做 A/B 复验（建议扫描 0%/10%/20% 三档），有收益证据再固化设值；工程上若保守设值，~10% 边界保底即可，不宜继续加大（行业惯例上限 20%）。复验记录应留档（策略、指标、数值、评测集版本）。`),
  p(T`上下文预算（C，检索侧，不应与 S 混淆，arXiv:2601.14123 F3/F4，C 级）：`),
  bullet(T`“上下文悬崖”：质量在 0.5k–2.5k tokens 区间稳定，超过 ~2.5k tokens 后下降（10k 时相对降约 4–5%）——更多上下文不等于更好；`),
  bullet(T`按目标调 C：语义质量在小上下文（~500 tokens）达峰；事实精确匹配（EM）在 ~2.5k tokens 达峰；问答默认 ~2.5k、摘要/解释类 ~500；`),
  bullet(T`模型相关：论文实测基于 Ministral-8B-Instruct，拐点位置应按所用 LLM 复调，不应照搬数值。`),

  h2("七、完整性检测标准"),
  p(T`截断检测（主判据）：统计“长度恰等于 chunk_size 且末字符非句读标点”的切片数，合格线 = 0。只计恰长的片，避免误伤合法满片。`),
  p(T`执行细则：`),
  bullet(T`检测应覆盖全量切片：接口分页返回时应翻页到底，只查首页会漏检；`),
  bullet(T`判定前剥离页眉、页脚、页码等版面噪声（如 PDF 页边界片尾的“句读＋页码数字”属合法边界，非硬切）；`),
  bullet(T`长度分布抽查：最长/最短各抽若干片人工通读，确认无“腰斩片”（以半句话开头/结尾）；`),
  bullet(T`目录对照抽查：按文档目录/标题索引逐节核对切片覆盖，专抓腰斩、散行这类检索测试测不出的问题。`),

  h2("八、自包含与元数据标准"),
  p(T`（字段构成定义见 §三，本节规定质量要求。）`),
  bullet(T`每个切片应携带来源信息：文档名、章节标题路径，必要时附摘要——公开实践中“切片自带标题路径”对向量质量有显著提升（C 级）；`),
  bullet(T`元数据字段（文档类型、时间、地域/主题、关键词等）应自动抽取并支持过滤检索；`),
  bullet(T`元数据质量参考口径（C 级，本标准自定经验起点，无外部标准出处）：完整性 ≥90%、准确性 ≥95%，期望值与实际值逐字段对账；`),
  bullet(T`抽取配置变更后应验证抽取结果真实更新（注意平台可能存在按“文件 × 配置”缓存抽取结果的行为）。`),

  h2("九、检索验收标准（最终判据）"),
  p(T`切片质量的最终判据是检索效果。验收分两层：`),
  p(T`1. 检索模块验收（模块化评价，B 级框架）。GB/Z 236—2026 §7.1 规定的检索模块指标共六项：命中率、精确度、召回率、MRR、NDCG、检索平均响应时间（公式 3–9，均无数值阈值）。本标准的合格线为 C 级经验值：`),
  metricTable,
  bullet(T`判据是“前 K 个召回切片含正确答案文本”，不是“有召回就算过”；`),
  bullet(T`评测集建设参照 GB/Z 236—2026 §8.3：参考答案的真实性、准确性为“应”（强制）；场景/题型/难度覆盖、偏差分析、构建流程记录、动态更新机制为“宜”（推荐）；规模上行业经验 50–100 条起步（C 级）；`),
  bullet(T`标注口径统一：单答案问题用命中率/MRR，多切片答案问题才需要召回率/精确度全标注；`),
  bullet(T`文本比对前做归一化（空白、换行、全半角），PDF 来源切片尤其应如此；`),
  bullet(T`不达标项应分类：评测集真值错误（改评测集，GB/Z §8.3 b) 要求确保参考答案真实性准确性）与检索缺陷（回切片/召回链路）分开判，不混为一谈。`),
  p(T`2. 端到端验收（B 级框架）。端到端三指标（GB/Z §6）：答案相关性、答案正确性、答案流畅度——有参考答案时按公式 (1)(2) 计算（信息点覆盖法），无参考答案时按表 1–3 的 1–5 分人工评分表打分（可直接用作评分标尺）；上下文忠实度按 GB/Z 公式 (10)（信息点交叉检查法）计算；幻觉率 <15% 为本标准自定参考红线（C 级，无外部标准出处）——RAG 的核心价值是基于检索回答，幻觉率居高不下即失去意义；进阶维度（GB/Z 附录 A，资料性）：拒答能力、信息整合能力、噪声鲁棒性、反事实鲁棒性。`),

  h2("十、不达标处置路径"),
  fixTable,

  h2("十一、质量闭环与回归纪律"),
  bullet(T`1. 固定一份回归评测集，每次改切分策略、调 chunk_size、换 embedding 模型都全量重跑——无量化度量则无法优化；`),
  bullet(T`2. 策略对比实验应固定检索与重排参数、只动切分，否则无法归因（公开实验：递归分块较固定长度 Recall@5 由 0.62 升至 0.74，绝对提升约 12 个百分点、相对提升约 19%，C 级）；`),
  bullet(T`3. 知识库内容更新后重跑评测集，防止旧知识污染新回答（GB/Z 236—2026 §8.3 f) 将其列为“宜”）；`),
  bullet(T`4. 上线后持续监控：空召回率、低分命中占比、答案忠实度，设阈值告警；`),
  bullet(T`5. 先跑基线数字，再看优化趋势——经验阈值是参考，不同业务场景（如医疗 vs 客服）容忍度差异很大。`),

  h2("参考依据"),
  bullet(T`1. ITU-T F.748.52 (03/25)，2025-03-16 批准、现行有效；§7.2 数据分段要求、§8.1 检索评估（Recall/MAP 公式）、附录 A 生成能力 5 级分级。`),
  bullet(T`2. GB/Z 236—2026《人工智能 检索增强生成系统评价指标与方法》（2026-08-27 发布，国家标准公告 2026 年第 37 号；全国标准信息公共服务平台）；§5 评价参考框架、§6 端到端三指标及表 1–3 评分表、§7 模块化指标公式 (3)–(10)、§8 评价过程、附录 A 其他评价维度。`),
  bullet(T`3. GB/T 45288.1-2025《人工智能 大模型 第 1 部分：通用要求》。`),
  bullet(T`4. GB/T 1.1—2020《标准化工作导则 第 1 部分：标准化文件的结构和起草规则》。`),
  bullet(T`5. arXiv:2601.14123（2026-01-20，工业场景 RAG 分块系统实验，NQ 数据集 + SPLADE + Ministral-8B）：F1 overlap 无可测量收益、索引按 1/(1−o) 膨胀；F2 方法梯度 句子≈语义>定长≫代码；F3 上下文悬崖（~2.5k tokens 拐点，模型相关）；F4 按目标调上下文预算。`),
  bulletAdd(T`6. CRUD-RAG：中文 RAG 基准（arXiv:2401.17043；ACM TOIS 43(2) 正式收录，2025，DOI 10.1145/3701228；近 30 万篇中文新闻语料 + 稠密检索 bge 等）：chunk size 64–512 与 overlap 0%–70% 系统扫描；最优 chunk size 因任务而异；overlap 对单文档 QA 关键信息指标有正向收益。C 级；局限：2024 年初、新闻语料、当时代模型，数值不可直接搬用。`),
  bulletAdd(T`7. arXiv:2604.12047（ICSE-SEIP 2026，金融 PDF 解析与分块实证，卢森堡大学 + BGL 法国巴黎银行）：overlap 扫描中 25% 检索层最优（MRR 0.529→0.658）、0% 最差。C 级；测量层次为检索层，与端到端结论不可直接互斥。`),
  bulletAdd(T`8. Legal-DC：中文法律文书 RAG 基准（arXiv:2603.11772，2026-03，预印本，未经同行评审）：480 份中国法律、行政法规与部门规章（覆盖市场监管、合同管理等）＋2,475 条条款级标注问答（两轮人工标注，Cohen's Kappa ≥ 0.83）；实验表明检索粒度是法律 RAG 的关键因素，条文级结构切分与语义 chunk 切分并行的多粒度检索优于单一粒度（Recall 78.02% vs LightRAG 75.86%，MRR@5 50.23% vs 47.56%）。C 级；引用时须注明预印本状态。`),
  bulletAdd(T`9. HOPE（A New HOPE: Domain-agnostic Automatic Evaluation of Text Chunking，ACM SIGIR 2025 正式收录，DOI 10.1145/3726302.3729882）：领域无关的切片质量自动评估框架；跨七领域实证发现“切片间语义独立性”对事实正确性影响最大（提升最高 56.2%）。C 级（ACM 正式收录，载体等级高于预印本）；通用领域实验，非中文专属。`),
  bullet(T`10. RAGAS / ARES / TruLens 评估体系（GB/Z 236—2026参考文献 [1][2] 亦收录 RAGAS）。`),
  bullet(T`11. 公开工程实践与 benchmark（结构感知分块对比实验等，C 级，多为单源）。`),

  h2("附录 A（资料性）中文数值基线候选清单（全部待验证）"),
  pAdd(T`本清单为 V3.1 新增的资料性附录：从 ACM TOIS 及其他可溯源渠道汇总面向中文的 RAG 数值基线，供自建基线时作候选起点。经核验，ACM TOIS 上面向中文的 RAG 基准仅 CRUD-RAG 一篇（TOIS 43(2) 目录确认）；清单以它为唯一 ACM 主干，其余中文来源（AAAI、机构发布、arXiv 预印本）一并列入并注明载体。全部条目按本标准分级均为 C 级公开研究，数值均非本库实测——一律标注【待验证】，验证方法：按 GB/Z 236—2026 §8.3 建评测集、固定检索与重排参数只动切分（单变量原则）、结果留档后回填本清单。`),
  baselineTable,

  // ================= 第二部分 =================
  h1("第二部分 RAG 知识库切片验收清单"),
  p(T`版本：V3.1（2026-11-17，修订草案，与《RAG 切片质量标准》版本对齐）。用途：每次建库/重切/新文档入库后照单过一遍。厂商中立，不绑定具体平台。来源分级：A＝ITU-T F.748.52；B＝GB/Z 236—2026；C＝行业经验值与本标准自定经验起点。所有数值合格线均为 C 级。用法：逐项打勾；任何一项不过，先定位再改，改完从该项起重新回归。`),

  h2("一、切分策略选型（切之前）"),
  check(T`每份文档已按结构选型，不一刀切：有标题/条款/章节结构 → 结构感知切分（多级回退）；一般自然语言文档 → 递归字符/句子切分（保底基线）；无标题纯长文（合同、论文）→ 可考虑语义切分（收益有限：句子/递归切分在约 5k tokens 内与语义切分持平，arXiv:2601.14123；C 级）；精度与上下文矛盾突出 → 父子切分；固定长度切分仅作对照基线，不单独用于生产`),
  check(T`切分能力覆盖 ITU-T F.748.52 §7.2 要求（A 级）：多种分段策略可选、最大分段长度与重叠长度可配置、表格数据可分段`),
  check(T`特殊内容已分别处置：表格整块成单元（保留表头）——平台固定按行切分+表头拼接的，已登记为等效合规路径并以检索验收判效（V3.1）；PDF/扫描件先版面还原再按段落切；代码按语法单元切`, true),
  check(T`若用正则/分隔符切分：锚点已在“经解析器处理后的正文”中验证真实可见，失效会静默退化为硬切`),
  check(T`若用标题层级切分：标题锚点已单独验证——标题标记（如 Markdown 的 #）可能被解析器剥除，剥除后锚不到、退硬切`),
  check(T`已确认平台 chunk_size 的单位（字符还是 token）与性质（软上限还是硬截断）`),
  check(T`overlap 设置（V3.1 修订）：默认从 0 起步；已在目标平台稠密检索链路上用固定评测集做 A/B 复验（建议 0%/10%/20% 三档）并留档；保守设值 ≤10% 边界保底，上限 20%。注意证据冲突：2601.14123（英文/稀疏/端到端）无收益，CRUD-RAG（中文/稠密）与 2604.12047（检索层）有收益——以自建基线为准`, true),
  check(T`平台不暴露的配置项（标题路径、元数据、overlap 等）已登记为“平台约束”并留档，不判为不达标（V3.1 新增）`, true),

  h2("二、切片结果核查（切之后）"),
  check(T`以实际切片输出为唯一判定依据，不信配置界面/接口回显`),
  check(T`截断检测已跑：统计“长度恰等于 chunk_size 且末字符非句读标点”的切片数，结果为 0`),
  check(T`检测覆盖全量切片：分页接口已翻页到底，未只查首页`),
  check(T`判定前已剥离页眉/页脚/页码噪声（如“句读＋页码数字”的页边界属合法边界，不应误报）`),
  check(T`长度分布合理：min 不是个位数碎片、max 未远超软上限；最长/最短各抽 3 片人工通读`),
  check(T`抽查确认无“腰斩片”（以半句话开头/结尾）；已按文档目录/标题索引逐节核对切片覆盖`),
  check(T`表格件无表行级切断（平台按行切分的例外见第一部分§五，V3.1）；如有 → 升一档 chunk_size 重导验证，够用即停`, true),
  check(T`重切后元数据已复核（注意平台可能缓存抽取结果，配置变更应验证真实生效）`),

  h2("三、自包含与元数据"),
  check(T`切片构成符合标准《切片构成规范》（§三）：必填成分（正文、来源文档、标题路径、元数据）无缺失；无标题结构件的标题路径置空已登记原因；平台不暴露的成分已登记为平台约束（V3.1）`, true),
  check(T`每个切片携带来源信息：文档名＋章节标题路径（必要时附摘要）`),
  check(T`随机抽 5 片，脱离原文能判断“出自哪份文件、哪个章节”`),
  check(T`元数据字段齐全且支持过滤检索`),
  check(T`元数据质量达标（C 级，自定经验起点）：完整性 ≥90%、准确性 ≥95%，期望值 vs 实际值逐字段对账`),

  h2("四、检索验收（最终判据）"),
  check(T`评测集就绪：50–100 条起步，覆盖不同问题类型与难度；标注口径统一`),
  check(T`Hit Rate@5 ≥ 0.85（C 级经验值）`),
  check(T`Recall@5 ≥ 0.90（生产口径，C 级；宽松基线 0.75）`),
  check(T`MRR ≥ 0.75（C 级经验值；公式定义见 GB/Z 236—2026公式 (6)，B 级）`),
  check(T`（可选）NDCG@5 用于观察 Top-K 排序质量（GB/Z 公式 (7)(8)；无数值阈值，看优化趋势）`),
  check(T`判据为“前 K 个召回切片含正确答案文本”，不是“有召回”`),
  check(T`检索上下文预算 C 已按任务调优（问答 ~2.5k tokens 起调、摘要 ~500；警惕超 ~2.5k 的“上下文悬崖”——拐点模型相关，应按所用 LLM 复调，arXiv:2601.14123，C 级）`),
  check(T`文本比对前已归一化（空白、换行、全半角），PDF 来源切片尤其应如此`),
  check(T`不达标项已分类：评测集真值错误（改评测集）vs 检索缺陷（回切片/召回链路）`),
  check(T`端到端复测（B 级框架）：答案相关性、正确性、流畅度按 GB/Z §6 公式 (1)(2) 或表 1–3 的 1–5 分人工评分表打分；忠实度按 GB/Z 公式 (10) 计算，幻觉率 <15% 为自定参考红线（C 级，无外部标准出处）`),
  check(T`评测集符合 GB/Z §8.3（B 级）：参考答案真实准确（“应”，强制）；覆盖场景/题型/难度、偏差分析与构建记录、动态更新机制（“宜”，推荐）`),

  h2("五、不达标处置路径"),
  fixTable,

  h2("六、收尾与回归纪律"),
  check(T`本次验收数据已留档：切分策略、chunk_size、overlap 及 A/B 复验记录（V3.1）、各指标数值、评测集版本`, true),
  check(T`回归评测集已固定；下次改切分/调参数/换模型时全量重跑`),
  check(T`策略对比实验遵守单变量原则：固定检索与重排参数，只动切分`),
  check(T`测试环境/临时资源已清理，无孤儿服务与残留数据`),
  check(T`上线后监控已配置：空召回率、低分命中占比、忠实度阈值告警`),
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
    1: { children: [para(run(T`Bennani & Moslonka, “A Systematic Analysis of Chunking Strategies for Reliable Question Answering”, arXiv:2601.14123, 2026-01；英文 NQ + SPLADE 稀疏检索 + Ministral-8B；overlap 10–20% 对 BERTScore/EM 无可测量收益。`, { size: 18 }))] },
    2: { children: [para(run(T`Lyu et al., “CRUD-RAG: A Comprehensive Chinese Benchmark for Retrieval-Augmented Generation of Large Language Models”, arXiv:2401.17043，ACM TOIS 43(2) 收录（DOI 10.1145/3701228）；中文新闻语料 + 稠密检索；单文档 QA 关键信息精确度 52.67→54.06、召回 70.82→73.94（0%→70% overlap）。`, { size: 18 }))] },
    3: { children: [para(run(T`El Bachyr et al., “Empirical Evaluation of PDF Parsing and Chunking for Financial Question Answering with RAG”, arXiv:2604.12047, ICSE-SEIP 2026；金融 PDF，检索层指标，25% overlap 最优（MRR 0.529→0.658），0% 最差。`, { size: 18 }))] },
  },
  sections: [
    {
      properties: {
        page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
      },
      headers: {
        default: new Header({
          children: [
            para(run(T`RAG 知识库切片质量标准与验收清单 V3.1（修订草案）`, { bold: true, color: palette.primary, size: 18 }), {
              alignment: AlignmentType.CENTER,
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [para(new TextRun({ children: [PageNumber.CURRENT] }), { alignment: AlignmentType.CENTER })],
        }),
      },
      children,
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync(outputPath, buffer);
