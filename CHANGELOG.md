# 更新日志

本项目的版本规则：标准版本号（V3.x）随文档内容演进，修订记录以 V3.1 文档内"修订说明与对照表"（R 编号）为准，本文件只记里程碑。

## [未发布]

### 变更
- 交付文档文件名按 VERSIONING.md §4 统一：评审意见补版本号（`_V2_修订版_2026-11-17`）、标准本体补日期（`_V3.1_修订草案_2026-11-17`）；README、`build:all`、Word 版交付包 README 同步更新
- 新增 `VERSIONING.md`：三层版本载体、版本对应矩阵、升版检查清单、命名与归档规则
- `tools/validate_docx.py`：支持 glob 通配符（修复 Windows 下 `npm run validate` 失败）、排除 Word/WPS `~$` 锁文件、校验不再短路、空校验显式失败
- `package.json`：补 `license: CC-BY-4.0` 与 `engines.node >=22`（与锁定依赖对齐），README 环境说明同步更正

## [v3.1.0] - 2026-11-17

首个公开版本，评审交付包。

### 新增
- 《RAG 切片标准平台冲突核查与文献适用性评审意见（修订版 V2）》：百炼/千帆/盘古（四入口）冲突核查，F-1～F-6 差异点清单；arXiv:2601.14123 适用性分层；中文证据链（CRUD-RAG／Legal-DC／HOPE）
- 《RAG 知识库切片质量标准与验收清单 V3.1（修订草案）》：R1–R9 修订；中文定位声明；附录 A 中文数值基线候选清单 B1–B7（全部【待验证】）
- 《RAG 切片证据速查表》（A4 横排单页）
- 《RAG 自建基线评测集建集模板 V1.0》
- 文档生成脚本（docx-js，单一事实源）、docx 完整性校验工具与 CI

[v3.1.0]: https://github.com/AmbOS-HK/rag-chunking-standard-zh/releases/tag/v3.1.0
