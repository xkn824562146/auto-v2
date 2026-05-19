"""Flask Web UI：上传 PDF + Word 模板，AI 提取后确认填充。"""

import logging
import os
import re
import shutil
import time
import uuid
import yaml
from flask import Flask, render_template, request, send_file, jsonify, after_this_request

from core.log_config import setup_logging

LOG_PATH = setup_logging()
logger = logging.getLogger(__name__)

from core.word_scanner import WordScanner
from core.vision_extractor import VisionExtractor
from core.field_mapper import FieldMapper
from core.report_filler import ReportFiller

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 加载配置
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.yaml")
with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    CONFIG = yaml.safe_load(f)


def _cleanup_stale_jobs():
    """启动时删除超过 TTL 的历史任务目录。"""
    ttl_hours = (CONFIG.get("storage") or {}).get("ttl_hours", 24)
    cutoff = time.time() - ttl_hours * 3600
    removed = 0
    if not os.path.isdir(UPLOAD_DIR):
        return
    for name in os.listdir(UPLOAD_DIR):
        job_dir = os.path.join(UPLOAD_DIR, name)
        if not os.path.isdir(job_dir):
            continue
        try:
            if os.path.getmtime(job_dir) < cutoff:
                shutil.rmtree(job_dir)
                removed += 1
        except OSError as e:
            logger.warning("清理过期任务目录失败 %s: %s", name, e)
    if removed:
        logger.info("启动清理：删除了 %d 个过期任务目录（TTL=%dh）", removed, ttl_hours)


_cleanup_stale_jobs()


def _parse_standard(label: str):
    """从标签中解析标准要求，返回 (比较运算符, 阈值) 或 None。"""
    # 匹配 ± 公差（如 ±0.04、±5）— 无法直接比较，标记为特殊类型
    m = re.search(r'[±]\s*([\d.]+)', label)
    if m:
        return '±', float(m.group(1))
    # 匹配 ≤ >= < > = 等运算符后跟数字
    m = re.search(r'([≤≥＜＞﹤≧≦<>=]+)\s*([\d.]+)', label)
    if m:
        return m.group(1), float(m.group(2))
    # 匹配 tf=0s 这种等号条件
    m = re.search(r'tf\s*=\s*([\d.]+)', label)
    if m:
        return '=', float(m.group(1))
    return None


def _compare(measured: str | None, standard_text: str) -> str:
    """将测量值与标准要求对比，返回合格/不合格。"""
    # 标准要求为空或纯 /，表示无实际要求，直接合格
    std_stripped = standard_text.strip()
    if not std_stripped or std_stripped == '/':
        return '合格'
    # 标准以 / 开头（如 "/ 60s内焰尖高度\nFS≤150mm"），去掉 / 前缀继续解析
    if std_stripped.startswith('/'):
        std_stripped = std_stripped[1:].strip()

    if measured is None:
        return '不合格'

    # 处理非数字测量值（AI 可能返回文字判断）
    measured_str = str(measured).strip()
    m = re.search(r'[\d.]+', measured_str)
    if not m:
        # 常见文字判断：否/无/合格/通过 = 合格，是/有/不合格/不通过 = 不合格
        _PASS_WORDS = {"否", "无", "合格", "通过", "正常", "符合"}
        _FAIL_WORDS = {"是", "有", "不合格", "不通过", "异常", "不符合"}
        if measured_str in _PASS_WORDS:
            return '合格'
        if measured_str in _FAIL_WORDS:
            return '不合格'
        return '不合格'
    measured_num = float(m.group())

    std = _parse_standard(std_stripped)
    if std is None:
        logger.warning("无法解析标准要求: %r", std_stripped)
        return '不合格'

    op, threshold = std
    if op in ('≤', '≦', '<=', '<', '＜', '﹤'):
        result = measured_num <= threshold
    elif op in ('≥', '≧', '>=', '>', '＞'):
        result = measured_num >= threshold
    elif op in ('=', '＝'):
        result = measured_num == threshold
    elif op == '±':
        result = -threshold <= measured_num <= threshold
    else:
        return '不合格'

    return '合格' if result else '不合格'


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert_doc", methods=["POST"])
def convert_doc():
    """将 .doc 文件转换为 .docx 格式。"""
    word_file = request.files.get("word")
    if not word_file:
        return jsonify({"error": "请上传 Word 文件"}), 400

    filename = word_file.filename.lower()
    if not filename.endswith(".doc"):
        return jsonify({"error": "文件不是 .doc 格式"}), 400

    # 保存上传的 .doc 文件
    job_id = uuid.uuid4().hex[:8]
    job_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    doc_path = os.path.join(job_dir, word_file.filename)
    word_file.save(doc_path)

    # 转换为 .docx
    docx_path = os.path.splitext(doc_path)[0] + ".docx"
    try:
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False

        doc = word_app.Documents.Open(os.path.abspath(doc_path))
        doc.SaveAs2(os.path.abspath(docx_path), FileFormat=16)  # 16 = docx format
        doc.Close()
        word_app.Quit()

        pythoncom.CoUninitialize()
    except Exception as e:
        logger.error("doc 转 docx 失败: %s", e)
        return jsonify({"error": f"文件转换失败: {str(e)}"}), 500

    # 删除原 .doc 文件
    try:
        os.remove(doc_path)
    except OSError:
        pass

    return jsonify({
        "job_id": job_id,
        "docx_filename": os.path.basename(docx_path),
    })


@app.route("/upload", methods=["POST"])
def upload():
    """处理上传的 PDF 和 Word 模板，返回 AI 提取结果。"""
    logger.info("收到上传请求")
    pdf_file = request.files.get("pdf")

    if not pdf_file:
        return jsonify({"error": "请上传 PDF 文件"}), 400

    # 检查是否是已转换的 .docx 文件
    word_job_id = request.form.get("word_job_id")
    word_filename = request.form.get("word_filename")

    if word_job_id and word_filename:
        # 使用已转换的文件
        job_id = word_job_id
        job_dir = os.path.join(UPLOAD_DIR, job_id)
        word_path = os.path.join(job_dir, word_filename)

        if not os.path.exists(word_path):
            return jsonify({"error": "转换后的 Word 文件不存在，请重新上传"}), 400

        # 保存 PDF 到同一个 job 目录
        pdf_path = os.path.join(job_dir, pdf_file.filename)
        pdf_file.save(pdf_path)
    else:
        # 直接上传 .docx 文件
        word_file = request.files.get("word")
        if not word_file:
            return jsonify({"error": "请上传 Word 模板文件"}), 400

        job_id = uuid.uuid4().hex[:8]
        job_dir = os.path.join(UPLOAD_DIR, job_id)
        os.makedirs(job_dir, exist_ok=True)

        pdf_path = os.path.join(job_dir, pdf_file.filename)
        word_path = os.path.join(job_dir, word_file.filename)
        pdf_file.save(pdf_path)
        word_file.save(word_path)

    # 第一步：扫描 Word 模板
    scanner = WordScanner()
    all_slots = scanner.scan(word_path)

    if not all_slots:
        return jsonify({"error": "未在 Word 模板中找到待填充的空字段"}), 400

    # 只保留"检验结果"和"单项结论"列的字段（注意表头可能有空格如"单项 结论"）
    TARGET_COLUMNS = {"检验结果", "单项结论"}
    slots = [
        s for s in all_slots
        if s.label.rsplit('-', 1)[-1].replace(" ", "") in TARGET_COLUMNS
    ]

    if not slots:
        # 回退：如果目标列过滤后无结果，使用全部 slots 并记录日志
        app.logger.warning(
            "目标列过滤无结果（未找到列头含'检验结果'/'单项结论'的字段），"
            "回退使用全部 %d 个 slots。labels: %s",
            len(all_slots),
            [s.label for s in all_slots],
        )
        slots = all_slots

    # 第二步：AI 视觉提取（去掉最后一个列标题后缀如"检验结果"，保留方向等中间信息）
    _TAIL = {"检验结果", "单项结论", "检测结果", "结论"}
    base_names = list(set(
        s.label.rsplit('-', 1)[0] if '-' in s.label and s.label.rsplit('-', 1)[-1].replace(' ', '') in _TAIL else s.label
        for s in slots
    ))
    aliases = CONFIG.get("aliases") or {}
    extractor = VisionExtractor(CONFIG["ai"], aliases=aliases)
    dpi = CONFIG["extraction"].get("dpi", 300)
    ai_result = extractor.extract(pdf_path, base_names, dpi=dpi)

    # 第三步：将 AI 结果映射回完整标签
    mapped = {}
    for s in slots:
        base = s.label.rsplit('-', 1)[0] if '-' in s.label and s.label.rsplit('-', 1)[-1].replace(' ', '') in _TAIL else s.label
        if s.label not in mapped and base in ai_result:
            mapped[s.label] = ai_result[base]

    # 构造 slots 序列化数据（前端展示用）
    # 只展示原始测量值字段，单项结论自动计算不展示
    slots_data = []
    for s in slots:
        raw_value = mapped.get(s.label)
        # 检验结果列：填入原始测量值
        # 单项结论列（标签含"单项"）：与标准要求对比，生成合格/不合格
        if '单项' in s.label:
            auto_value = _compare(raw_value, s.standard)
            logger.info("[upload] %s: raw_value=%r standard=%r -> %s", s.label, raw_value, s.standard, auto_value)
            mapped[s.label] = auto_value
        elif '检验结果' in s.label:
            # 检验结果直接使用原始测量值，不做比较
            auto_value = raw_value
        else:
            auto_value = raw_value

        # 前端不展示单项结论（自动计算），其余字段展示供用户确认
        if '单项' not in s.label:
            slots_data.append({
                "table_idx": s.table_idx,
                "row_idx": s.row_idx,
                "cell_idx": s.cell_idx,
                "label": s.label,
                "value": raw_value,
            })

    return jsonify({
        "job_id": job_id,
        "pdf_filename": pdf_file.filename,
        "word_filename": word_filename if word_job_id else word_file.filename,
        "fields": mapped,
        "slots": slots_data,
    })


@app.route("/confirm", methods=["POST"])
def confirm():
    """用户确认后执行回填，返回生成的 Word 文件。"""
    data = request.json
    job_id = data.get("job_id")
    fields = data.get("fields", {})

    if not job_id:
        return jsonify({"error": "缺少 job_id"}), 400

    job_dir = os.path.join(UPLOAD_DIR, job_id)
    if not os.path.isdir(job_dir):
        return jsonify({"error": "任务不存在"}), 404

    # 找到上传的文件
    pdf_path = None
    word_path = None
    for f in os.listdir(job_dir):
        if f.lower().endswith(".pdf"):
            pdf_path = os.path.join(job_dir, f)
        elif f.lower().endswith(".docx"):
            word_path = os.path.join(job_dir, f)

    if not word_path:
        return jsonify({"error": "找不到 Word 模板文件"}), 400

    # 重新扫描获取 slots（带格式信息），只保留目标列
    scanner = WordScanner()
    all_slots = scanner.scan(word_path)
    TARGET_COLUMNS = {"检验结果", "单项结论"}
    slots = [
        s for s in all_slots
        if s.label.rsplit('-', 1)[-1].replace(" ", "") in TARGET_COLUMNS
    ]
    if not slots:
        slots = all_slots

    # 自动计算"单项结论"列（前端只发回了检验结果的值）
    for s in slots:
        if '单项' in s.label and s.label not in fields:
            base = s.label.rsplit('-', 1)[0]
            result_key = f"{base}-检验结果"
            raw_value = fields.get(result_key)
            result = _compare(raw_value, s.standard)
            logger.info("[confirm] %s: raw_value=%r standard=%r -> %s", s.label, raw_value, s.standard, result)
            fields[s.label] = result

    # 执行回填
    output_path = os.path.join(job_dir, "完成版_检测报告.docx")
    filler = ReportFiller()
    logs = filler.fill(word_path, output_path, slots, fields)

    return jsonify({
        "download_url": f"/download/{job_id}",
        "logs": logs,
    })


@app.route("/download/<job_id>")
def download(job_id):
    """下载生成的 Word 文件，下载完成后自动清理任务目录。"""
    output_path = os.path.join(UPLOAD_DIR, job_id, "完成版_检测报告.docx")
    if not os.path.exists(output_path):
        return jsonify({"error": "文件不存在"}), 404

    @after_this_request
    def _schedule_cleanup(response):
        job_dir = os.path.join(UPLOAD_DIR, job_id)
        try:
            shutil.rmtree(job_dir)
            logger.info("已清理任务目录: %s", job_id)
        except OSError as e:
            logger.warning("清理任务目录失败 %s: %s", job_id, e)
        return response

    return send_file(output_path, as_attachment=True, download_name="完成版_检测报告.docx")


@app.route("/pdf_image/<job_id>/<int:page>")
def pdf_image(job_id, page):
    """返回 PDF 指定页的图片（供前端预览）。"""
    import fitz
    from io import BytesIO

    job_dir = os.path.join(UPLOAD_DIR, job_id)
    pdf_path = None
    for f in os.listdir(job_dir):
        if f.lower().endswith(".pdf"):
            pdf_path = os.path.join(job_dir, f)
            break

    if not pdf_path:
        return jsonify({"error": "找不到 PDF"}), 404

    doc = fitz.open(pdf_path)
    if page < 0 or page >= len(doc):
        doc.close()
        return jsonify({"error": "页码超出范围"}), 400

    page_obj = doc[page]
    mat = fitz.Matrix(1.5, 1.5)  # 1.5x 缩放
    pix = page_obj.get_pixmap(matrix=mat)
    img_bytes = pix.tobytes("png")
    doc.close()

    return send_file(BytesIO(img_bytes), mimetype="image/png")


@app.route("/pdf_pages/<job_id>")
def pdf_pages(job_id):
    """返回 PDF 总页数。"""
    import fitz

    job_dir = os.path.join(UPLOAD_DIR, job_id)
    pdf_path = None
    for f in os.listdir(job_dir):
        if f.lower().endswith(".pdf"):
            pdf_path = os.path.join(job_dir, f)
            break

    if not pdf_path:
        return jsonify({"error": "找不到 PDF"}), 404

    doc = fitz.open(pdf_path)
    count = len(doc)
    doc.close()
    return jsonify({"pages": count})


def _warn_if_port_busy(port: int) -> None:
    import socket

    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.bind(("127.0.0.1", port))
    except OSError:
        logger.warning(
            "端口 %s 已被占用，可能已有其它 python app.py 在运行；"
            "请求会打到别的进程，本终端可能看不到日志。请先结束旧进程再启动。",
            port,
        )
    finally:
        sock.close()


if __name__ == "__main__":
    import os

    port = int(os.environ.get("PORT", "5000"))
    # use_reloader=False：Windows 下重载子进程 stdout 常不可见
    _warn_if_port_busy(port)
    logger.info(
        "启动服务 http://127.0.0.1:%s  PID=%s  日志文件: %s",
        port, os.getpid(), LOG_PATH,
    )
    app.run(host="0.0.0.0", port=port, debug=True, use_reloader=False)
