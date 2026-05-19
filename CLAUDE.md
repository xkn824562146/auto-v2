# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PDF 检测报告智能提取与 Word 模板回填工具。用户上传 PDF 检测报告 + Word 模板，AI 视觉提取 PDF 中的字段值，模糊映射后回填到 Word 模板的空单元格中，保留原始格式。

## Running the Application

```bash
pip install -r requirements.txt
python app.py
# Server starts at http://localhost:5000
```

No build step. No test suite. Python 3.10+ required (uses `X | Y` type union syntax).

## Architecture

Linear data pipeline — the orchestration lives entirely in `app.py`, not in a separate pipeline class:

```
WordScanner → VisionExtractor → (inline mapping in app.py) → ReportFiller
```

- **`core/word_scanner.py`** — Scans `.docx` tables for empty cells. Determines labels by finding the header row (via keyword detection for "序号"/"检验项目"/"标准要求" etc.), then combining column header + row label. Handles merged cells, sub-items (经向/纬横向), unit labels, and "以下空白" stop markers. Returns `list[FieldSlot]` with `(table_idx, row_idx, cell_idx, label, standard, fmt)`.
- **`core/vision_extractor.py`** — Converts PDF pages to base64 PNG via PyMuPDF, sends to OpenAI-compatible multimodal API, parses JSON response with regex + fuzzy key matching. Supports field aliases (from `config.yaml`) for fields that appear under different names in PDFs. Exits early once all fields found.
- **`core/field_mapper.py`** — Standalone mapper (exact + fuzzy match via `thefuzz`). **Not used in the current `app.py` flow** — mapping is done inline in the upload route with additional label-stripping logic.
- **`core/report_filler.py`** — Writes values into exact cell coordinates preserving format (bold, italic, font name/size, color, alignment, East Asian fonts via `w:rFonts`).

**Flask routes** (`app.py`):
- `/` — Upload UI (`templates/index.html`)
- `/convert_doc` (POST) — `.doc` → `.docx` via Windows COM (`win32com`); Windows-only
- `/upload` (POST) — Scans Word template, calls AI extraction, returns fields for user confirmation
- `/confirm` (POST) — User confirms values, executes fill, returns download URL
- `/download/<job_id>` (GET) — Downloads generated `.docx`, then auto-deletes the job directory via `@after_this_request`
- `/pdf_image/<job_id>/<page>` / `/pdf_pages/<job_id>` — PDF preview endpoints

Uploads stored in `uploads/<job_id>/`.

**Pass/fail auto-comparison** (`app.py:_compare`): After extraction, measured values in "检验结果" columns are compared against standard requirements parsed from the same row (e.g. `≤0.16`). The "单项结论" column is auto-filled with "合格"/"不合格" without user confirmation. The `_parse_standard` function handles operators: ≤, ≥, <, >, =, and `tf=N` notation.

**Label construction** (in `app.py` upload route): For "检验结果"/"单项结论" columns, the column suffix is stripped to create a `base_name` for AI extraction (e.g. `拉伸断裂强力-经向-检验结果` → `拉伸断裂强力-经向`). After extraction, values are mapped back to full labels.

**Logging** (`core/log_config.py`): Dual output to stderr + `logs/app.log`. Suppresses noisy libraries (urllib3, httpx, openai) to WARNING level. Uses a guard flag to prevent duplicate handler registration.

## Configuration

`config.yaml` keys:
- `ai.base_url` / `ai.api_key` / `ai.model` — OpenAI-compatible API settings
- `extraction.dpi` — PDF rendering resolution (default 300)
- `extraction.fuzzy_threshold` — Fuzzy match threshold (default 60)
- `aliases` — Field name aliases for AI extraction (key = Word label, value = list of PDF aliases). Example: `内焰尖高度: [焰尖高度是否达到150mm刻度线, 焰尖高度]`
- `extraction.null_handling` — Defined but unused in code
- `storage.ttl_hours` — TTL for auto-cleanup of stale job directories (default 24)

**Storage cleanup** (`app.py`): Two mechanisms prevent `uploads/` from growing unbounded:
1. After `/download/<job_id>` serves the file, `@after_this_request` deletes the entire `uploads/<job_id>/` directory.
2. On startup, `_cleanup_stale_jobs()` scans `uploads/` and deletes any directory older than `storage.ttl_hours` (default 24). This catches abandoned uploads and any failed post-download cleanup.

## Key Dependencies

- `python-docx` — Word read/write
- `PyMuPDF` (`fitz`) — PDF rendering
- `openai` — Multimodal AI API client (OpenAI-compatible)
- `thefuzz` + `python-Levenshtein` — Fuzzy string matching
- `flask` — Web server
- `pyyaml` — Config parsing
- `pywin32` — Windows COM automation for `.doc` → `.docx` conversion (Windows-only, lazy imported)

## Language

All user-facing text (UI, logs) is in Chinese. Code comments and variable names are in English.
