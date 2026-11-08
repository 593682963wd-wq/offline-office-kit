# 离线办公文件核心

版本 0.1.0。本项目源码采用 AGPL-3.0 许可证，见 LICENSE；第三方依赖仍适用各自上游许可。

这是从已有实现筛选出的通用算法和必要离线适配，保留实用计算路径；并非任何原应用的完整迁移。未包含原始业务数据、既有文档模板、账号、网络适配、模型权重或历史仓库。来源追溯保存在独立的私有审查材料中。

## 运行

Python 3.13 已实测；源码使用 Python 3.11+ 语法，其他运行时尚未验证。在本目录执行：

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python demo.py
.venv/bin/python -m unittest discover -s tests -v
```

`demo.py` 是可直接运行的库使用示例，输入完全新造且明显标注为合成。`python demo.py --help` 查看说明。演示生成的文件/数据库都放临时目录并清理；需要处理自己的数据时，参照示例显式调用下面模块，不存在自动发现本机目录或登录操作。

## 能力与边界

时刻只接受明确 HH:MM/HHMM（扩展小时至47:59）。DOCX数据行含嵌套表、合并格或内容控件时拒绝变更；只支持规则矩形表。PDF仅可检索文字，不含OCR/自动翻译；提取表格降级会返回告警。PDF元数据配对仅支持代码可识别且唯一1:1的输入，当前元数据语法针对以Z开头的AD2文档，其他文档会明确无法配对。PPT复制不宣称支持所有媒体/图表关系；测试覆盖文本页。slide_xml接收调用方可信的XML片段。PyMuPDF依赖仍适用其上游许可；本项目的 AGPL-3.0 许可不改变依赖条款。

仅离线合成测试和文件回读已验证。没有真实业务数据验收、当前法规适用性验证、生产部署或外部发送测试。读取自己的文件不等于可以公开其内容；输出和使用参数仍由调用方审核。

## 模块接口

接口参数以源码类型标注和测试用例为准。下表包含实现接口，带下划线的函数是保留的内部辅助接口。

| 模块 | 主要符号 |
| --- | --- |
| `office/date_counts.py` | `weekday_pattern`, `execution_count` |
| `office/extract.py` | `_openpyxl_to_text`, `_cell_to_text` |
| `office/pdf_metadata.py` | `PdfMetadata`, `PairingError`, `_normalise_date`, `_clean_airport_name`, `extract_metadata`, `pair_pdf_paths` |
| `office/pdf_overlay.py` | `_normalize_text`, `_load_base_translation`, `generate_translation_template`, `_translation_font_size`, `_translation_required_height`, `render_bilingual_pdf` |
| `office/pdf_text.py` | `_format_text_line`, `_inside_table`, `_table_to_markdown`, `_pdf_to_markdown_with_warnings` |
| `office/ppt_edit.py` | `_p_text`, `_set_p_text`, `set_text_keep_format`, `set_multi_paragraphs`, `set_cell`, `remove_shape`, `delete_slides_after`, `clone_slide_after` |
| `office/scenario_dates.py` | `_exec_in_window` |
| `office/scenario_overlay.py` | `overlay` |
| `office/section_diff.py` | `SectionDiffItem`, `_clean_compare_line`, `markdown_sections`, `_normalise_for_compare`, `compare_markdown` |
| `office/slide_xml.py` | `emu`, `rr`, `para`, `Builder` |
| `offline_office_core/compare.py` | `parse_clock`, `duration_minutes`, `signed_clock_delta`, `format_minutes`, `TableSchema`, `Change`, `read_table`, `compare_files`, `write_changes`, `main` |
| `offline_office_core/docx_table.py` | `_normal_header`, `fill_table` |
| `offline_office_core/excel.py` | `center_worksheet`, `center_workbook`, `save_centered_workbook` |

## 2026-11-08 协作修订

- 比较保留布尔与整数类型差异。
- PPT清空单元格后仍保留有效段落。
- Markdown表格转义反斜线和竖线。
- 纯标题变化也纳入章节比较。

集合根目录 `verify_release.py --project domain-office` 可复跑合成测试和首次使用示例；本轮以集合验证报告为准。
