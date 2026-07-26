"""
电子发票识别工具 - Web 服务
启动:
  python app.py                 → 本机 http://127.0.0.1:5000
  python app.py --share         → 额外生成外网临时链接（无需云服务器）
  python app.py --cloud          → 腾讯云等服务器部署（不弹浏览器、默认 CPU）
  双击 invoice_tool.bat         → 本机使用
  双击 invoice_tool_share.bat   → 分享给非本机用户
  bash deploy/start_cloud.sh    → 云服务器一键安装并启动

环境变量（云部署常用）:
  INVOICE_TOOL_CLOUD=1
  INVOICE_TOOL_ACCESS_TOKEN=口令   → 开启访问保护（推荐公网开启）
  INVOICE_TOOL_PORT=5000
  INVOICE_TOOL_HOST=0.0.0.0
  INVOICE_TOOL_DEVICE=cpu
  INVOICE_TOOL_NO_BROWSER=1
  INVOICE_TOOL_TRUST_PROXY=1      → 前面有 Nginx 时开启
"""

import os, sys, io, shutil, datetime, threading, uuid, glob, re, webbrowser, atexit, argparse, time, hashlib, zipfile, hmac
from flask import Flask, render_template, request, jsonify, send_file

# 导入核心逻辑
_stdout_backup = sys.stdout
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
from invoice_extractor import (
    _discover_folders, extract_all_invoices, write_to_excel, InvoiceInfo,
    reset_ocr_if_failed, _safe_float, set_progress_callback,
    INVOICE_EXTENSIONS, is_supported_invoice_file,
    set_ocr_mode, normalize_ocr_mode, OCR_MODE_BALANCED,
)
from table_extractor import process_table_images, is_table_image_file
sys.stdout = sys.__stdout__

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB
app.config['TEMPLATES_AUTO_RELOAD'] = True  # 改前端后刷新即可生效，无需重启

# 公网访问口令（空 = 不校验）。Header: X-Access-Token 或 ?token= / Bearer
ACCESS_TOKEN = (os.environ.get('INVOICE_TOOL_ACCESS_TOKEN') or '').strip()
_AUTH_EXEMPT_PATHS = frozenset({'/api/health', '/api/auth-config'})


def _apply_proxy_fix():
    """Nginx 反代时修正 scheme/host，便于生成正确链接。"""
    if os.environ.get('INVOICE_TOOL_TRUST_PROXY', '').strip() not in ('1', 'true', 'yes'):
        return
    try:
        from werkzeug.middleware.proxy_fix import ProxyFix
        app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
    except Exception:
        pass


_apply_proxy_fix()


def _extract_access_token():
    auth = (request.headers.get('Authorization') or '').strip()
    if auth.lower().startswith('bearer '):
        return auth[7:].strip()
    token = (request.headers.get('X-Access-Token') or '').strip()
    if token:
        return token
    return (request.args.get('token') or '').strip()


def _token_ok(provided):
    if not ACCESS_TOKEN:
        return True
    if not provided:
        return False
    return hmac.compare_digest(provided, ACCESS_TOKEN)


@app.before_request
def _require_access_token():
    if not ACCESS_TOKEN:
        return None
    path = request.path or ''
    if path in _AUTH_EXEMPT_PATHS or path == '/':
        return None
    if not path.startswith('/api/'):
        return None
    if _token_ok(_extract_access_token()):
        return None
    return jsonify({'ok': False, 'error': '未授权：请提供正确的访问口令'}), 401


@app.errorhandler(400)
@app.errorhandler(404)
@app.errorhandler(405)
@app.errorhandler(413)
@app.errorhandler(500)
def _api_json_error(err):
    """保证 /api/* 始终返回 JSON，避免前端 JSON.parse 失败。"""
    code = getattr(err, 'code', 500) or 500
    if not request.path.startswith('/api/'):
        # 非 API 走 Werkzeug 默认 HTML 响应
        if hasattr(err, 'get_response'):
            return err.get_response()
        return ('Error', code)

    if code == 413:
        msg = '文件过大，请压缩后重试或分批上传'
    elif code == 404:
        msg = '接口不存在，请刷新页面后重试'
    elif code == 405:
        msg = '请求方法不允许'
    elif code == 400:
        msg = getattr(err, 'description', None) or '请求无效'
    else:
        msg = '服务器内部错误，请稍后重试'
    return jsonify({'ok': False, 'error': msg}), code

UPLOAD_DIR = os.path.join(SCRIPT_DIR, 'uploads')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'excel')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

MAX_TASKS = 50
TASK_ID_RE = re.compile(r'^\d{8}_\d{6}_[a-f0-9]{6}$')

# 生成的 Excel 保留天数；启动时清理一次，之后每周再清理一次
EXCEL_RETENTION_DAYS = 7
EXCEL_CLEANUP_INTERVAL_SEC = 7 * 24 * 3600

tasks = {}
tasks_lock = threading.Lock()
_excel_cleanup_stop = threading.Event()


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
    ext = os.path.splitext(parts[-1])[1].lower() if parts else ''
    if not parts or ext not in INVOICE_EXTENSIONS:
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


def cleanup_old_excel_files(retention_days=EXCEL_RETENTION_DAYS):
    """删除 excel/ 目录中超过保留天数的生成文件，防止堆积。"""
    if not os.path.isdir(OUTPUT_DIR):
        return 0

    cutoff = time.time() - retention_days * 24 * 3600
    removed = 0
    patterns = ('*.xlsx', '*.xls', '*.xlsm')

    for pattern in patterns:
        for path in glob.glob(os.path.join(OUTPUT_DIR, '**', pattern), recursive=True):
            if not os.path.isfile(path):
                continue
            try:
                mtime = os.path.getmtime(path)
                if mtime >= cutoff:
                    continue
                os.remove(path)
                removed += 1
                print(f'  [清理] 已删除过期 Excel: {os.path.relpath(path, OUTPUT_DIR)}', file=sys.stderr)
            except OSError as e:
                print(f'  [警告] 无法删除 Excel {path}: {e}', file=sys.stderr)

    # 顺带去掉清理后留下的空子目录
    for root, dirnames, filenames in os.walk(OUTPUT_DIR, topdown=False):
        if root == OUTPUT_DIR:
            continue
        try:
            if not os.listdir(root):
                os.rmdir(root)
        except OSError:
            pass

    if removed:
        print(f'  [清理] 共删除 {removed} 个超过 {retention_days} 天的 Excel 文件', file=sys.stderr)
    return removed


def _excel_cleanup_loop():
    """后台定时器：启动先清一次，之后每隔一周再清一次。"""
    try:
        cleanup_old_excel_files()
    except Exception as e:
        print(f'  [警告] Excel 定时清理失败: {e}', file=sys.stderr)

    while not _excel_cleanup_stop.wait(EXCEL_CLEANUP_INTERVAL_SEC):
        try:
            cleanup_old_excel_files()
        except Exception as e:
            print(f'  [警告] Excel 定时清理失败: {e}', file=sys.stderr)


def start_excel_cleanup_timer():
    """启动 Excel 定期清理后台线程（daemon，随进程退出）。"""
    t = threading.Thread(target=_excel_cleanup_loop, name='excel-cleanup', daemon=True)
    t.start()
    print(
        f'  [清理] 已启用 Excel 定期清理：保留最近 {EXCEL_RETENTION_DAYS} 天，'
        f'每 {EXCEL_CLEANUP_INTERVAL_SEC // 86400} 天检查一次',
        file=sys.stderr,
    )
    return t


def run_extraction(input_dir, output_file, task_id):
    """后台提取"""
    old_stdout = sys.stdout
    sys.stdout = io.StringIO()
    try:
        task = _get_task(task_id) or {}
        ocr_mode = normalize_ocr_mode(task.get('ocr_mode') or OCR_MODE_BALANCED)
        set_ocr_mode(ocr_mode)
        mode_label = {'fast': '速度', 'balanced': '平衡', 'accurate': '精度'}.get(ocr_mode, '平衡')
        _update_task(task_id, status='running', message=f'正在扫描文件（{mode_label}模式）...', progress=1)
        reset_ocr_if_failed()
        folders = _discover_folders(input_dir)
        if not folders:
            _update_task(task_id, status='error', error='未找到 PDF 或图片文件')
            return

        folder_results = {}
        total_files = sum(len(p) for p in folders.values())
        processed = 0

        def _on_page_progress(message, page_progress=None):
            # 无 page_progress 的日志只更新文案，避免把进度条打回 0%
            if page_progress is not None and total_files:
                overall = (processed + page_progress / 100.0) / total_files
                pct = int(overall * 100)
                _update_task(task_id, message=message, progress=min(pct, 99))
            else:
                _update_task(task_id, message=message)

        set_progress_callback(_on_page_progress)

        for folder_name, file_paths in folders.items():
            invoices = []
            for file_path in file_paths:
                _update_task(
                    task_id,
                    message=f'准备识别: {os.path.basename(file_path)}',
                    progress=int(processed / total_files * 100) if total_files else 0,
                )
                results = extract_all_invoices(file_path, source_folder=folder_name)
                invoices.extend(results)
                processed += 1
                progress = int(processed / total_files * 100) if total_files else 0
                _update_task(
                    task_id,
                    progress=progress,
                    message=f'识别中 ({processed}/{total_files}): {os.path.basename(file_path)}',
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
    return render_template(
        'index.html',
        access_token_required=bool(ACCESS_TOKEN),
    )


@app.route('/api/health')
def health():
    """云监控 / 负载均衡探活（无需口令）。"""
    return jsonify({'ok': True, 'service': 'invoice-tool'})


@app.route('/api/auth-config')
def auth_config():
    """前端据此决定是否弹出入口令。"""
    return jsonify({'ok': True, 'required': bool(ACCESS_TOKEN)})


# ==================== 分片并行上传（加快外网隧道传输） ====================

@app.route('/api/upload-init', methods=['POST'])
def upload_init():
    """创建上传批次，供前端并行分片上传。"""
    data = request.get_json(silent=True) or {}
    mode = (data.get('mode') or request.form.get('mode') or 'files').strip()
    if mode not in ('files', 'folder'):
        mode = 'files'
    ocr_mode = normalize_ocr_mode(data.get('ocr_mode') or OCR_MODE_BALANCED)

    batch_id = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_') + uuid.uuid4().hex[:6]
    batch_dir = os.path.join(UPLOAD_DIR, batch_id)
    os.makedirs(batch_dir, exist_ok=True)
    _init_task(batch_id, {
        'status': 'uploading',
        'progress': 0,
        'message': '正在上传...',
        'mode': mode,
        'ocr_mode': ocr_mode,
        'file_count': 0,
    })
    return jsonify({'ok': True, 'task_id': batch_id, 'ocr_mode': ocr_mode})


@app.route('/api/upload-part/<task_id>', methods=['POST'])
def upload_part(task_id):
    """接收一批文件（可被前端并行多次调用）。"""
    try:
        if not _validate_task_id(task_id):
            return jsonify({'ok': False, 'error': '无效的任务 ID'}), 400
        task = _get_task(task_id)
        if not task or task.get('status') != 'uploading':
            return jsonify({'ok': False, 'error': '任务不可上传'}), 400

        batch_dir = os.path.join(UPLOAD_DIR, task_id)
        if not os.path.isdir(batch_dir):
            return jsonify({'ok': False, 'error': '上传目录不存在'}), 400

        mode = task.get('mode') or 'files'
        try:
            files = request.files.getlist('files')
        except Exception as e:
            return jsonify({'ok': False, 'error': f'解析上传数据失败: {e}'}), 400
        if not files:
            return jsonify({'ok': False, 'error': '未选择文件'}), 400

        count = 0
        last_err = None
        for f in files:
            if not f.filename or not is_supported_invoice_file(f.filename):
                continue
            try:
                if mode == 'folder':
                    dest = _safe_upload_dest(batch_dir, f.filename)
                    os.makedirs(os.path.dirname(dest), exist_ok=True)
                else:
                    dest = _safe_upload_dest(batch_dir, os.path.basename(f.filename))
                dest = _unique_dest_if_exists(dest)
                f.save(dest)
                count += 1
            except ValueError as e:
                last_err = str(e)
                continue
            except OSError as e:
                last_err = str(e)
                continue

        if count == 0:
            return jsonify({
                'ok': False,
                'error': last_err or '未上传任何有效文件',
            }), 400

        with tasks_lock:
            t = tasks.get(task_id)
            if t and t.get('status') == 'uploading':
                t['file_count'] = int(t.get('file_count') or 0) + count
                total = t['file_count']
            else:
                total = count
        return jsonify({'ok': True, 'added': count, 'file_count': total})
    except Exception as e:
        return jsonify({'ok': False, 'error': f'上传失败: {e}'}), 500


@app.route('/api/upload-chunk/<task_id>', methods=['POST'])
def upload_chunk(task_id):
    """大文件分块上传（外网隧道用）：每块快速回包，最后一块组装成完整文件。"""
    try:
        if not _validate_task_id(task_id):
            return jsonify({'ok': False, 'error': '无效的任务 ID'}), 400
        task = _get_task(task_id)
        if not task or task.get('status') != 'uploading':
            return jsonify({'ok': False, 'error': '任务不可上传'}), 400

        batch_dir = os.path.join(UPLOAD_DIR, task_id)
        if not os.path.isdir(batch_dir):
            return jsonify({'ok': False, 'error': '上传目录不存在'}), 400

        filename = (request.form.get('filename') or '').strip()
        try:
            chunk_index = int(request.form.get('chunk_index', -1))
            chunk_total = int(request.form.get('chunk_total', -1))
        except (TypeError, ValueError):
            return jsonify({'ok': False, 'error': '分块参数无效'}), 400

        chunk = request.files.get('chunk')
        if not filename or chunk is None:
            return jsonify({'ok': False, 'error': '缺少文件名或分块数据'}), 400
        if chunk_index < 0 or chunk_total < 1 or chunk_index >= chunk_total:
            return jsonify({'ok': False, 'error': '分块序号无效'}), 400
        if not is_supported_invoice_file(filename):
            return jsonify({'ok': False, 'error': '不支持的文件类型'}), 400

        mode = task.get('mode') or 'files'
        # 分块临时目录（按文件名哈希，避免中文路径问题）
        chunk_key = hashlib.md5(filename.encode('utf-8')).hexdigest()
        chunk_dir = os.path.join(batch_dir, '.chunks', chunk_key)
        os.makedirs(chunk_dir, exist_ok=True)
        part_path = os.path.join(chunk_dir, f'{chunk_index:06d}.part')
        chunk.save(part_path)

        # 未到最后一块：立刻返回，避免 Cloudflare 100s 超时
        if chunk_index < chunk_total - 1:
            return jsonify({
                'ok': True,
                'chunk_index': chunk_index,
                'chunk_total': chunk_total,
                'done': False,
            })

        # 最后一块：校验并组装
        missing = [
            i for i in range(chunk_total)
            if not os.path.isfile(os.path.join(chunk_dir, f'{i:06d}.part'))
        ]
        if missing:
            return jsonify({
                'ok': False,
                'error': f'分块不完整，缺少: {missing[:8]}',
            }), 400

        try:
            if mode == 'folder':
                dest = _safe_upload_dest(batch_dir, filename)
                os.makedirs(os.path.dirname(dest), exist_ok=True)
            else:
                dest = _safe_upload_dest(batch_dir, os.path.basename(filename))
            dest = _unique_dest_if_exists(dest)
            with open(dest, 'wb') as out:
                for i in range(chunk_total):
                    p = os.path.join(chunk_dir, f'{i:06d}.part')
                    with open(p, 'rb') as inp:
                        shutil.copyfileobj(inp, out, length=1024 * 1024)
        finally:
            shutil.rmtree(chunk_dir, ignore_errors=True)

        with tasks_lock:
            t = tasks.get(task_id)
            if t and t.get('status') == 'uploading':
                t['file_count'] = int(t.get('file_count') or 0) + 1
                total = t['file_count']
            else:
                total = 1
        return jsonify({
            'ok': True,
            'chunk_index': chunk_index,
            'chunk_total': chunk_total,
            'done': True,
            'added': 1,
            'file_count': total,
        })
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    except Exception as e:
        return jsonify({'ok': False, 'error': f'分块上传失败: {e}'}), 500


@app.route('/api/upload-commit/<task_id>', methods=['POST'])
def upload_commit(task_id):
    """全部文件上传完成后开始识别。"""
    if not _validate_task_id(task_id):
        return jsonify({'ok': False, 'error': '无效的任务 ID'})
    task = _get_task(task_id)
    if not task or task.get('status') != 'uploading':
        return jsonify({'ok': False, 'error': '任务状态无效'})

    file_count = int(task.get('file_count') or 0)
    if file_count <= 0:
        return jsonify({'ok': False, 'error': '未上传任何文件'})

    batch_dir = os.path.join(UPLOAD_DIR, task_id)
    threading.Thread(target=_cleanup_old_uploads, daemon=True).start()
    output_file = os.path.join(OUTPUT_DIR, f'发票识别结果_{task_id}.xlsx')
    _update_task(
        task_id,
        status='pending',
        progress=0,
        message='上传完成，正在开始识别...',
    )
    threading.Thread(target=run_extraction, args=(batch_dir, output_file, task_id), daemon=True).start()
    return jsonify({'ok': True, 'task_id': task_id, 'file_count': file_count})


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
        if not f.filename or not is_supported_invoice_file(f.filename):
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
        return jsonify({'ok': False, 'error': '未上传任何 PDF 或图片文件'})

    # 清理放到后台，避免拖慢上传响应（外网隧道场景更明显）
    threading.Thread(target=_cleanup_old_uploads, daemon=True).start()

    ocr_mode = normalize_ocr_mode(request.form.get('ocr_mode') or OCR_MODE_BALANCED)
    output_file = os.path.join(OUTPUT_DIR, f'发票识别结果_{batch_id}.xlsx')
    _init_task(batch_id, {
        'status': 'pending',
        'progress': 0,
        'message': '上传完成，正在开始识别...',
        'ocr_mode': ocr_mode,
    })
    threading.Thread(target=run_extraction, args=(batch_dir, output_file, batch_id), daemon=True).start()

    return jsonify({'ok': True, 'task_id': batch_id, 'file_count': count, 'ocr_mode': ocr_mode})


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
        if not f.filename or not is_supported_invoice_file(f.filename):
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
        return jsonify({'ok': False, 'error': '未上传任何 PDF 或图片文件'})

    threading.Thread(target=_cleanup_old_uploads, daemon=True).start()

    ocr_mode = normalize_ocr_mode(request.form.get('ocr_mode') or OCR_MODE_BALANCED)
    output_file = os.path.join(OUTPUT_DIR, f'发票识别结果_{batch_id}.xlsx')
    _init_task(batch_id, {
        'status': 'pending',
        'progress': 0,
        'message': '上传完成，正在开始识别...',
        'ocr_mode': ocr_mode,
    })
    threading.Thread(target=run_extraction, args=(batch_dir, output_file, batch_id), daemon=True).start()

    return jsonify({'ok': True, 'task_id': batch_id, 'file_count': count, 'ocr_mode': ocr_mode})


# ==================== 表格转 Excel（独立功能） ====================

def run_table_extraction(input_dir, output_dir, task_id):
    """后台：每张图片生成同名 Excel。"""
    try:
        task = _get_task(task_id) or {}
        ocr_mode = normalize_ocr_mode(task.get('ocr_mode') or OCR_MODE_BALANCED)
        _update_task(task_id, status='running', message='正在识别表格...', progress=1)

        def _cb(message, pct):
            _update_task(task_id, message=message, progress=min(int(pct), 99))

        result = process_table_images(input_dir, output_dir, ocr_mode=ocr_mode, progress_cb=_cb)
        if not result.get('ok'):
            _update_task(task_id, status='error', error=result.get('error') or '表格识别失败', progress=0)
            return

        files = []
        for item in result.get('files') or []:
            files.append({
                'ok': bool(item.get('ok')),
                'source': item.get('source') or '',
                'excel': item.get('excel') or '',
                'rows': item.get('rows') or 0,
                'cols': item.get('cols') or 0,
                'error': item.get('error') or '',
            })

        # 打包 zip，方便一次下载
        zip_path = os.path.join(output_dir, f'表格识别结果_{task_id}.zip')
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for item in result.get('files') or []:
                p = item.get('excel_path') or ''
                if item.get('ok') and p and os.path.isfile(p):
                    zf.write(p, arcname=os.path.basename(p))

        _update_task(
            task_id,
            status='done',
            progress=100,
            message=f"完成：成功 {result.get('success')}/{result.get('total')}",
            feature='table',
            output_file=zip_path,
            output_dir=output_dir,
            table_files=files,
            success=result.get('success') or 0,
            total=result.get('total') or 0,
            invoices=[],
        )
    except Exception as e:
        _update_task(task_id, status='error', error=str(e))


@app.route('/api/table/upload-init', methods=['POST'])
def table_upload_init():
    data = request.get_json(silent=True) or {}
    mode = (data.get('mode') or 'files').strip()
    if mode not in ('files', 'folder'):
        mode = 'files'
    ocr_mode = normalize_ocr_mode(data.get('ocr_mode') or OCR_MODE_BALANCED)
    batch_id = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_') + uuid.uuid4().hex[:6]
    batch_dir = os.path.join(UPLOAD_DIR, batch_id)
    os.makedirs(batch_dir, exist_ok=True)
    _init_task(batch_id, {
        'status': 'uploading',
        'progress': 0,
        'message': '正在上传表格图片...',
        'mode': mode,
        'ocr_mode': ocr_mode,
        'feature': 'table',
        'file_count': 0,
    })
    return jsonify({'ok': True, 'task_id': batch_id, 'ocr_mode': ocr_mode})


@app.route('/api/table/upload-part/<task_id>', methods=['POST'])
def table_upload_part(task_id):
    """复用发票上传落盘逻辑，但只接受图片。"""
    # 临时切换校验：仅图片
    if not _validate_task_id(task_id):
        return jsonify({'ok': False, 'error': '无效的任务 ID'}), 400
    task = _get_task(task_id)
    if not task or task.get('status') != 'uploading' or task.get('feature') != 'table':
        return jsonify({'ok': False, 'error': '任务不可上传'}), 400

    batch_dir = os.path.join(UPLOAD_DIR, task_id)
    if not os.path.isdir(batch_dir):
        return jsonify({'ok': False, 'error': '上传目录不存在'}), 400

    mode = task.get('mode') or 'files'
    try:
        files = request.files.getlist('files')
    except Exception as e:
        return jsonify({'ok': False, 'error': f'解析上传数据失败: {e}'}), 400
    if not files:
        return jsonify({'ok': False, 'error': '未选择文件'}), 400

    count = 0
    last_err = None
    for f in files:
        if not f.filename or not is_table_image_file(f.filename):
            continue
        try:
            if mode == 'folder':
                dest = _safe_upload_dest(batch_dir, f.filename)
                os.makedirs(os.path.dirname(dest), exist_ok=True)
            else:
                dest = _safe_upload_dest(batch_dir, os.path.basename(f.filename))
            dest = _unique_dest_if_exists(dest)
            f.save(dest)
            count += 1
        except ValueError as e:
            last_err = str(e)
            continue
        except OSError as e:
            last_err = str(e)
            continue

    if count == 0:
        return jsonify({'ok': False, 'error': last_err or '未上传任何图片文件'}), 400

    with tasks_lock:
        t = tasks.get(task_id)
        if t and t.get('status') == 'uploading':
            t['file_count'] = int(t.get('file_count') or 0) + count
            total = t['file_count']
        else:
            total = count
    return jsonify({'ok': True, 'added': count, 'file_count': total})


@app.route('/api/table/upload-chunk/<task_id>', methods=['POST'])
def table_upload_chunk(task_id):
    """表格大图分块上传（与发票分块逻辑一致，仅校验图片）。"""
    if not _validate_task_id(task_id):
        return jsonify({'ok': False, 'error': '无效的任务 ID'}), 400
    task = _get_task(task_id)
    if not task or task.get('feature') != 'table':
        return jsonify({'ok': False, 'error': '任务不可上传'}), 400
    filename = (request.form.get('filename') or '').strip()
    if filename and not is_table_image_file(filename):
        return jsonify({'ok': False, 'error': '表格模式仅支持图片文件'}), 400
    return upload_chunk(task_id)


@app.route('/api/table/upload-commit/<task_id>', methods=['POST'])
def table_upload_commit(task_id):
    if not _validate_task_id(task_id):
        return jsonify({'ok': False, 'error': '无效的任务 ID'})
    task = _get_task(task_id)
    if not task or task.get('status') != 'uploading' or task.get('feature') != 'table':
        return jsonify({'ok': False, 'error': '任务状态无效'})

    file_count = int(task.get('file_count') or 0)
    if file_count <= 0:
        return jsonify({'ok': False, 'error': '未上传任何文件'})

    batch_dir = os.path.join(UPLOAD_DIR, task_id)
    out_dir = os.path.join(OUTPUT_DIR, f'table_{task_id}')
    os.makedirs(out_dir, exist_ok=True)
    threading.Thread(target=_cleanup_old_uploads, daemon=True).start()
    _update_task(
        task_id,
        status='pending',
        progress=0,
        message='上传完成，正在开始表格识别...',
        output_dir=out_dir,
    )
    threading.Thread(target=run_table_extraction, args=(batch_dir, out_dir, task_id), daemon=True).start()
    return jsonify({'ok': True, 'task_id': task_id, 'file_count': file_count})


@app.route('/api/table/download/<task_id>/<path:filename>')
def table_download_one(task_id, filename):
    if not _validate_task_id(task_id):
        return jsonify({'ok': False, 'error': '无效的任务 ID'})
    task = _get_task(task_id)
    if not task or task.get('feature') != 'table':
        return jsonify({'ok': False, 'error': '文件不存在'})
    out_dir = task.get('output_dir') or ''
    if not out_dir or not os.path.isdir(out_dir):
        return jsonify({'ok': False, 'error': '文件不存在'})
    # 仅允许文件名，禁止路径穿越
    safe_name = os.path.basename(filename)
    path = os.path.join(out_dir, safe_name)
    if not os.path.isfile(path):
        return jsonify({'ok': False, 'error': '文件不存在'})
    abs_out = os.path.abspath(out_dir)
    abs_path = os.path.abspath(path)
    try:
        if os.path.commonpath([abs_out, abs_path]) != abs_out:
            return jsonify({'ok': False, 'error': '文件不存在'})
    except ValueError:
        return jsonify({'ok': False, 'error': '文件不存在'})
    return send_file(path, as_attachment=True, download_name=safe_name)


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
        'feature': task.get('feature') or 'invoice',
        'output_file': os.path.basename(task.get('output_file', '')),
        'invoices': task.get('invoices', []),
        'table_files': task.get('table_files', []),
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


_share_tunnel = None


def _env_truthy(name):
    return (os.environ.get(name) or '').strip().lower() in ('1', 'true', 'yes', 'on')


def _apply_low_mem_runtime():
    """2G 小机优化：限制 CPU 线程，减轻 Paddle/OpenMP 抢内存。"""
    os.environ.setdefault('OMP_NUM_THREADS', '1')
    os.environ.setdefault('MKL_NUM_THREADS', '1')
    os.environ.setdefault('OPENBLAS_NUM_THREADS', '1')
    os.environ.setdefault('FLAGS_use_mkldnn', '0')


def _apply_cloud_defaults(low_mem=False):
    """云部署默认：不弹浏览器、无 GPU 时用 CPU。"""
    os.environ.setdefault('INVOICE_TOOL_NO_BROWSER', '1')
    if not (os.environ.get('INVOICE_TOOL_DEVICE') or '').strip():
        os.environ['INVOICE_TOOL_DEVICE'] = 'cpu'
    if low_mem or _env_truthy('INVOICE_TOOL_LOW_MEM'):
        os.environ['INVOICE_TOOL_LOW_MEM'] = '1'
        _apply_low_mem_runtime()


def _start_server(host='0.0.0.0', port=5000, enable_share=False, enable_fixed_share=False, cloud_mode=False):
    global _share_tunnel

    low_mem = _env_truthy('INVOICE_TOOL_LOW_MEM')

    print('=' * 50)
    print('  电子发票识别工具')
    print(f'  本机: http://127.0.0.1:{port}')
    if cloud_mode:
        print(f'  监听: http://{host}:{port}  （云部署）')
        print('  公网: http://<服务器公网IP>:' + str(port))
        if low_mem:
            print('  低内存模式: 已开启（适合 2G；建议加 2G Swap，文本发票优先）')
        if ACCESS_TOKEN:
            print('  访问口令: 已启用（INVOICE_TOOL_ACCESS_TOKEN）')
        else:
            print('  [安全] 未设置访问口令，公网可被任何人使用')
            print('         建议: export INVOICE_TOOL_ACCESS_TOKEN=你的口令')
    else:
        try:
            from tunnel_share import print_lan_hints
            print_lan_hints(port)
        except Exception:
            print(f'  局域网: http://<你的IP>:{port}')
    print('=' * 50)

    start_excel_cleanup_timer()

    # 2G 机器预热易 OOM：改为首次识别图片时再加载模型
    skip_warmup = low_mem or _env_truthy('INVOICE_TOOL_SKIP_OCR_WARMUP')

    def _warm_ocr():
        try:
            time.sleep(3)
            from invoice_extractor import _get_ocr, reset_ocr_if_failed, set_ocr_mode, OCR_MODE_BALANCED, OCR_MODE_FAST
            reset_ocr_if_failed()
            # 低内存默认用速度档，模型与渲染更省
            mode = OCR_MODE_FAST if low_mem else OCR_MODE_BALANCED
            set_ocr_mode(mode)
            print(f'  [OCR] 正在预热模型（{mode}，后台）...', flush=True)
            _get_ocr()
            print('  [OCR] 预热完成，可立即识别', flush=True)
        except Exception as e:
            print(f'  [OCR] 预热跳过（不影响文本型 PDF）: {e}', flush=True)

    if skip_warmup:
        print('  [OCR] 已跳过启动预热（低内存/显式关闭），首次图片识别时再加载', flush=True)
    else:
        threading.Thread(target=_warm_ocr, daemon=True, name='ocr-warmup').start()

    if enable_share or enable_fixed_share:
        def _boot_share():
            global _share_tunnel
            time.sleep(1.2)
            try:
                if enable_fixed_share:
                    from tunnel_fixed import start_fixed_share_tunnel
                    _share_tunnel = start_fixed_share_tunnel(port)
                else:
                    from tunnel_share import start_share_tunnel
                    _share_tunnel = start_share_tunnel(port)

                def _cleanup_tunnel():
                    if _share_tunnel:
                        _share_tunnel.stop()

                atexit.register(_cleanup_tunnel)
                if os.environ.get('INVOICE_TOOL_NO_BROWSER') != '1' and getattr(_share_tunnel, 'public_url', None):
                    try:
                        webbrowser.open(_share_tunnel.public_url)
                    except Exception:
                        pass
            except Exception as e:
                print(f'  [分享] 外网链接创建失败: {e}')
                if enable_fixed_share:
                    print('  [分享] 请先双击 setup_fixed_tunnel.bat 完成一次初始化')
                print('  [分享] 仍可使用本机 / 局域网地址')

        threading.Thread(target=_boot_share, daemon=True).start()

    try:
        from waitress import serve
        # 2G / 云小机少线程，避免 OCR 时并发把内存打满
        if low_mem:
            threads = 4
        elif cloud_mode:
            threads = 8
        else:
            threads = 16
        serve(app, host=host, port=port, threads=threads, channel_timeout=300, connection_limit=100 if low_mem else 200)
    except ImportError:
        print('  [提示] 未安装 waitress，使用 Flask 内置服务器（建议 pip install waitress）')
        app.run(host=host, port=port, debug=False)


def _parse_args():
    parser = argparse.ArgumentParser(description='电子发票识别工具')
    parser.add_argument(
        '--share', action='store_true',
        help='开启临时公网链接（每次网址会变化）',
    )
    parser.add_argument(
        '--share-fixed', action='store_true',
        help='开启固定域名分享（需先运行 setup_fixed_tunnel.bat）',
    )
    parser.add_argument(
        '--cloud', action='store_true',
        help='云服务器模式（不弹浏览器，默认 CPU，打印公网访问提示）',
    )
    parser.add_argument(
        '--low-mem', action='store_true',
        help='低内存模式（适合 2G：跳过 OCR 预热、少线程、限 CPU 并行）',
    )
    parser.add_argument('--host', default=None, help='监听地址，默认 0.0.0.0')
    parser.add_argument('--port', type=int, default=None, help='服务端口，默认 5000')
    return parser.parse_args()


if __name__ == '__main__':
    args = _parse_args()
    cloud_mode = bool(args.cloud) or _env_truthy('INVOICE_TOOL_CLOUD')
    low_mem = bool(args.low_mem) or _env_truthy('INVOICE_TOOL_LOW_MEM')
    if cloud_mode:
        _apply_cloud_defaults(low_mem=low_mem)
    elif low_mem:
        os.environ['INVOICE_TOOL_LOW_MEM'] = '1'
        _apply_low_mem_runtime()

    enable_fixed = args.share_fixed or os.environ.get('INVOICE_TOOL_SHARE_FIXED') == '1'
    enable_share = (not enable_fixed) and (not cloud_mode) and (
        args.share or os.environ.get('INVOICE_TOOL_SHARE') == '1'
    )
    host = (args.host or os.environ.get('INVOICE_TOOL_HOST') or '0.0.0.0').strip() or '0.0.0.0'
    port = args.port
    if port is None:
        try:
            port = int((os.environ.get('INVOICE_TOOL_PORT') or '5000').strip())
        except ValueError:
            port = 5000

    if os.environ.get('INVOICE_TOOL_NO_BROWSER') != '1':
        threading.Timer(1.5, lambda: webbrowser.open(f'http://127.0.0.1:{port}')).start()
    _start_server(
        host=host,
        port=port,
        enable_share=enable_share,
        enable_fixed_share=enable_fixed,
        cloud_mode=cloud_mode,
    )
