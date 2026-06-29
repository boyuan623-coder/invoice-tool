"""
电子发票识别工具 - Web 服务
启动: python app.py  或  双击 invoice_tool.bat  →  访问 http://127.0.0.1:5000
"""

import os, sys, io, shutil, datetime, threading, uuid, glob, re, webbrowser
from flask import Flask, render_template, request, jsonify, send_file

# 导入核心逻辑
_stdout_backup = sys.stdout
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
from invoice_extractor import (
    _discover_folders, extract_all_invoices, write_to_excel, InvoiceInfo,
    reset_ocr_if_failed, _safe_float, set_progress_callback,
)
sys.stdout = sys.__stdout__

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB

UPLOAD_DIR = os.path.join(SCRIPT_DIR, 'uploads')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'excel')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

MAX_TASKS = 50
TASK_ID_RE = re.compile(r'^\d{8}_\d{6}_[a-f0-9]{6}$')

tasks = {}
tasks_lock = threading.Lock()


def _validate_task_id(task_id):
    return bool(task_id and TASK_ID_RE.match(task_id))


def _get_task(task_id):
    with tasks_lock:
        return tasks.get(task_id)


def _init_task(task_id, data):
    with tasks_lock:
        tasks[task_id] = data
    _cleanup_old_tasks()


def _update_task(task_id, **kwargs):
    with tasks_lock:
        if task_id in tasks:
            tasks[task_id].update(kwargs)


def _cleanup_old_tasks():
    """保留最近 MAX_TASKS 个任务，移除更早任务的明细数据以节省内存。"""
    with tasks_lock:
        if len(tasks) <= MAX_TASKS:
            return
        sorted_ids = sorted(tasks.keys())
        for old_id in sorted_ids[: len(tasks) - MAX_TASKS]:
            task = tasks.get(old_id)
            if task and task.get('status') in ('done', 'error'):
                task.pop('invoices', None)


def _safe_output_path(path):
    """确保输出文件位于 OUTPUT_DIR 内。"""
    abs_path = os.path.abspath(path)
    abs_output = os.path.abspath(OUTPUT_DIR)
    try:
        return os.path.commonpath([abs_output, abs_path]) == abs_output
    except ValueError:
        return False


def _safe_upload_dest(batch_dir, filename):
    """规范化上传路径，防止目录穿越并限制在 batch_dir 内。"""
    rel_path = filename.replace('\\', '/').lstrip('/')
    if not rel_path or rel_path.startswith('..') or '/../' in f'/{rel_path}/':
        raise ValueError(f'非法路径: {filename}')
    if re.match(r'^[a-zA-Z]:', rel_path):
        raise ValueError(f'非法路径: {filename}')

    parts = [p for p in rel_path.split('/') if p and p not in ('.', '..')]
    if not parts or not parts[-1].lower().endswith('.pdf'):
        raise ValueError(f'无效的上传路径: {filename}')

    dest = os.path.normpath(os.path.join(batch_dir, *parts))
    batch_abs = os.path.abspath(batch_dir)
    try:
        if os.path.commonpath([batch_abs, dest]) != batch_abs:
            raise ValueError(f'非法路径: {filename}')
    except ValueError:
        raise ValueError(f'非法路径: {filename}') from None
    return dest


def _unique_dest_if_exists(dest):
    """同名 PDF 并存时自动追加序号，避免覆盖。"""
    if not os.path.exists(dest):
        return dest
    folder, basename = os.path.dirname(dest), os.path.basename(dest)
    stem, ext = os.path.splitext(basename)
    for i in range(1, 10000):
        candidate = os.path.join(folder, f'{stem}_{i}{ext}')
        if not os.path.exists(candidate):
            return candidate
    raise ValueError(f'同名文件过多: {basename}')


def _cleanup_old_uploads():
    """保留最近 50 个上传批次"""
    dirs = sorted(
        [d for d in glob.glob(os.path.join(UPLOAD_DIR, '*')) if os.path.isdir(d)],
        key=os.path.getctime
    )
    while len(dirs) > 50:
        old_dir = dirs.pop(0)
        try:
            shutil.rmtree(old_dir)
        except OSError as e:
            print(f'  [警告] 无法删除旧上传目录 {old_dir}: {e}', file=sys.stderr)


def run_extraction(input_dir, output_file, task_id):
    """后台提取"""
    old_stdout = sys.stdout
    sys.stdout = io.StringIO()
    try:
        reset_ocr_if_failed()
        _update_task(task_id, status='running')
        folders = _discover_folders(input_dir)
        if not folders:
            _update_task(task_id, status='error', error='未找到 PDF 文件')
            return

        folder_results = {}
        total_pdfs = sum(len(p) for p in folders.values())
        processed = 0

        def _on_page_progress(message, page_progress=None):
            overall = processed
            if page_progress is not None and total_pdfs:
                overall = (processed + page_progress / 100.0) / total_pdfs
            pct = int(overall * 100) if total_pdfs else 0
            _update_task(task_id, message=message, progress=min(pct, 99))

        set_progress_callback(_on_page_progress)

        for folder_name, pdf_paths in folders.items():
            invoices = []
            for pdf_path in pdf_paths:
                _update_task(
                    task_id,
                    message=f'准备识别: {os.path.basename(pdf_path)}',
                    progress=int(processed / total_pdfs * 100) if total_pdfs else 0,
                )
                results = extract_all_invoices(pdf_path, source_folder=folder_name)
                invoices.extend(results)
                processed += 1
                progress = int(processed / total_pdfs * 100) if total_pdfs else 0
                _update_task(
                    task_id,
                    progress=progress,
                    message=f'识别中 ({processed}/{total_pdfs}): {os.path.basename(pdf_path)}',
                )

            folder_results[folder_name] = [{
                'invoice_number': inv.invoice_number,
                'invoice_date': inv.invoice_date,
                'buyer_name': inv.buyer_name,
                'seller_name': inv.seller_name,
                'subtotal': inv.subtotal,
                'total_with_tax': inv.total_with_tax,
                'phone_number': inv.phone_number,
                'billing_period': inv.billing_period,
                'extraction_method': inv.extraction_method,
                'file_name': inv.file_name,
                'source_folder': inv.source_folder,
            } for inv in invoices]

        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        folder_groups = {}
        for folder_name, inv_dicts in folder_results.items():
            objs = []
            for d in inv_dicts:
                obj = InvoiceInfo()
                for k, v in d.items():
                    setattr(obj, k, v)
                objs.append(obj)
            folder_groups[folder_name] = objs

        inv_objects = [inv for invs in folder_groups.values() for inv in invs]
        is_multi = len(folder_groups) > 1
        write_to_excel(
            inv_objects, output_file,
            show_folder_col=False,
            folder_groups=folder_groups if is_multi else None,
        )

        flat = [inv for invs in folder_results.values() for inv in invs]
        _update_task(
            task_id,
            status='done',
            progress=100,
            message='识别完成',
            output_file=output_file,
            invoices=flat,
            total=sum(_safe_float(inv['total_with_tax']) for inv in flat),
            success=sum(1 for inv in flat if inv['invoice_number']),
        )
    except Exception as e:
        _update_task(task_id, status='error', error=str(e))
    finally:
        set_progress_callback(None)
        sys.stdout = old_stdout


# ==================== 页面 ====================

@app.route('/')
def index():
    return render_template('index.html')


# ==================== 上传文件（扁平） ====================

@app.route('/api/upload-files', methods=['POST'])
def upload_files():
    """上传 PDF 文件（所有文件视为同一批）"""
    files = request.files.getlist('files')
    if not files:
        return jsonify({'ok': False, 'error': '未选择文件'})

    batch_id = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_') + uuid.uuid4().hex[:6]
    batch_dir = os.path.join(UPLOAD_DIR, batch_id)
    os.makedirs(batch_dir, exist_ok=True)

    count = 0
    for f in files:
        if not f.filename or not f.filename.lower().endswith('.pdf'):
            continue
        try:
            dest = _safe_upload_dest(batch_dir, os.path.basename(f.filename))
            dest = _unique_dest_if_exists(dest)
            f.save(dest)
            count += 1
        except ValueError:
            continue

    if count == 0:
        shutil.rmtree(batch_dir, ignore_errors=True)
        return jsonify({'ok': False, 'error': '未上传任何 PDF 文件'})

    _cleanup_old_uploads()

    output_file = os.path.join(OUTPUT_DIR, f'发票识别结果_{batch_id}.xlsx')
    _init_task(batch_id, {'status': 'pending', 'progress': 0, 'message': '准备中...'})
    threading.Thread(target=run_extraction, args=(batch_dir, output_file, batch_id), daemon=True).start()

    return jsonify({'ok': True, 'task_id': batch_id, 'file_count': count})


# ==================== 上传文件夹（保留子目录结构） ====================

@app.route('/api/upload-folder', methods=['POST'])
def upload_folder():
    """上传整个文件夹（保留子目录结构）"""
    files = request.files.getlist('files')
    if not files:
        return jsonify({'ok': False, 'error': '未选择文件'})

    batch_id = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_') + uuid.uuid4().hex[:6]
    batch_dir = os.path.join(UPLOAD_DIR, batch_id)
    os.makedirs(batch_dir, exist_ok=True)

    count = 0
    for f in files:
        if not f.filename or not f.filename.lower().endswith('.pdf'):
            continue
        try:
            dest = _safe_upload_dest(batch_dir, f.filename)
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            dest = _unique_dest_if_exists(dest)
            f.save(dest)
            count += 1
        except ValueError:
            continue

    if count == 0:
        shutil.rmtree(batch_dir, ignore_errors=True)
        return jsonify({'ok': False, 'error': '未上传任何 PDF 文件'})

    _cleanup_old_uploads()

    output_file = os.path.join(OUTPUT_DIR, f'发票识别结果_{batch_id}.xlsx')
    _init_task(batch_id, {'status': 'pending', 'progress': 0, 'message': '准备中...'})
    threading.Thread(target=run_extraction, args=(batch_dir, output_file, batch_id), daemon=True).start()

    return jsonify({'ok': True, 'task_id': batch_id, 'file_count': count})


# ==================== 通用接口 ====================

@app.route('/api/status/<task_id>')
def task_status(task_id):
    if not _validate_task_id(task_id):
        return jsonify({'ok': False, 'error': '无效的任务 ID'})
    task = _get_task(task_id)
    if not task:
        return jsonify({'ok': False, 'error': '任务不存在'})
    return jsonify({
        'ok': True,
        'status': task.get('status'),
        'progress': task.get('progress', 0),
        'message': task.get('message', ''),
        'error': task.get('error', ''),
        'output_file': os.path.basename(task.get('output_file', '')),
        'invoices': task.get('invoices', []),
        'total': task.get('total', 0),
        'success': task.get('success', 0),
    })


@app.route('/api/download/<task_id>')
def download_file(task_id):
    if not _validate_task_id(task_id):
        return jsonify({'ok': False, 'error': '无效的任务 ID'})
    task = _get_task(task_id)
    if not task or not task.get('output_file'):
        return jsonify({'ok': False, 'error': '文件不存在'})
    path = task['output_file']
    if not _safe_output_path(path):
        return jsonify({'ok': False, 'error': '文件不存在'})
    if os.path.exists(path):
        return send_file(path, as_attachment=True, download_name=os.path.basename(path))
    return jsonify({'ok': False, 'error': '文件不存在'})


def _start_server(host='0.0.0.0', port=5000):
    print('=' * 50)
    print('  电子发票识别工具')
    print(f'  本机: http://127.0.0.1:{port}')
    print(f'  局域网: http://<你的IP>:{port}')
    print('=' * 50)

    try:
        from waitress import serve
        serve(app, host=host, port=port, threads=4)
    except ImportError:
        print('  [提示] 未安装 waitress，使用 Flask 内置服务器（建议 pip install waitress）')
        app.run(host=host, port=port, debug=False)


if __name__ == '__main__':
    if os.environ.get('INVOICE_TOOL_NO_BROWSER') != '1':
        threading.Timer(1.5, lambda: webbrowser.open('http://127.0.0.1:5000')).start()
    _start_server()
