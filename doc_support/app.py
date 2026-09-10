import io
import os
import shutil
import subprocess
import tempfile
import traceback
import zipfile
from pathlib import Path

from flask import Flask, request, send_file, render_template_string, make_response
from werkzeug.exceptions import RequestEntityTooLarge, ClientDisconnected, BadRequest

from generator import build_report
from footnote_fix import fix_footnote_format

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 60 * 1024 * 1024

HTML = '''<!doctype html>
<html lang="vi"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>Tạo báo cáo tuần VKSND tỉnh Ninh Bình</title>
<style>
*{box-sizing:border-box}body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;color:#1f2937}.wrap{max-width:780px;margin:24px auto;padding:0 14px}.card{background:#fff;border-radius:14px;box-shadow:0 8px 28px rgba(0,0,0,.08);padding:24px}.brand{font-weight:700;color:#991b1b;font-size:22px}.sub{color:#6b7280;margin:8px 0 20px;line-height:1.5}.field{margin:15px 0}.field label{display:block;font-weight:700;margin-bottom:7px}.field input{width:100%;border:1px solid #d1d5db;padding:12px;border-radius:8px;background:#fff;font-size:15px}.file-name{font-size:13px;color:#166534;margin-top:6px;min-height:16px;word-break:break-word}.hint{font-size:13px;color:#6b7280;margin-top:5px;line-height:1.45}.btn{margin-top:20px;width:100%;border:0;border-radius:9px;padding:15px 18px;min-height:52px;background:#991b1b;color:#fff;font-weight:700;font-size:16px;cursor:pointer;touch-action:manipulation}.btn:disabled{opacity:.65;cursor:wait}.status{display:none;margin-top:14px;padding:12px;border-radius:8px;background:#eff6ff;color:#1e3a8a;font-size:14px;line-height:1.45}.note{margin-top:18px;padding:12px;border-radius:8px;background:#fff7ed;font-size:13px;line-height:1.5}.err{background:#fef2f2;color:#991b1b;padding:12px;border-radius:8px;margin-bottom:16px;line-height:1.45}@media(max-width:560px){.wrap{margin:10px auto}.card{padding:18px}.brand{font-size:20px}.btn{font-size:15px}}
</style></head><body><div class="wrap"><div class="card">
<div class="brand">TẠO BÁO CÁO TUẦN</div>
<div class="sub">VKSND tỉnh Ninh Bình — hệ thống dùng chính báo cáo gốc làm nền, cập nhật số liệu từ Excel và báo cáo Phòng 7, Phòng 9; số liệu tự động điền hiển thị màu đỏ và footnote ở dạng số mũ.</div>
{% if error %}<div class="err"><b>Không thể tạo báo cáo.</b><br>{{ error }}</div>{% endif %}
<form id="reportForm" method="post" enctype="multipart/form-data" action="/">
<div class="field"><label>1. Bảng Excel báo cáo tuần</label><input id="excel" type="file" name="excel" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required><div id="excelName" class="file-name"></div></div>
<div class="field"><label>2. Báo cáo Phòng 7</label><input id="p7" type="file" name="p7" accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required><div id="p7Name" class="file-name"></div></div>
<div class="field"><label>3. Báo cáo Phòng 9</label><input id="p9" type="file" name="p9" accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required><div id="p9Name" class="file-name"></div></div>
<div class="field"><label>4. Báo cáo gốc dùng làm mẫu</label><input id="template" type="file" name="template" accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required><div id="templateName" class="file-name"></div><div class="hint">Chấp nhận cả .doc và .docx. File .doc được chuyển sang .docx trên máy chủ trước khi điền số liệu để giữ bố cục và định dạng tốt nhất.</div></div>
<button id="submitBtn" class="btn" type="submit">TẠO VÀ TẢI BÁO CÁO WORD</button><div id="status" class="status">Đang tải tệp và tạo báo cáo Word…</div>
</form>
<div class="note"><b>Nguyên tắc:</b> .docx được dùng trực tiếp. .doc được chuyển đổi bằng LibreOffice rồi mới cập nhật số liệu; hệ thống giữ thiết lập trang, căn đoạn, tab, giãn dòng và cấu trúc footnote của bản chuyển đổi.</div>
</div></div>
<script>(function(){const ids=['excel','p7','p9','template'];ids.forEach(id=>{const el=document.getElementById(id),label=document.getElementById(id+'Name');el.addEventListener('change',()=>{label.textContent=el.files&&el.files[0]?'Đã chọn: '+el.files[0].name:'';});});const form=document.getElementById('reportForm'),btn=document.getElementById('submitBtn'),status=document.getElementById('status');form.addEventListener('submit',function(e){const missing=ids.filter(id=>!document.getElementById(id).files.length);if(missing.length){e.preventDefault();status.style.display='block';status.textContent='Vui lòng chọn đủ 4 tệp trước khi tạo báo cáo.';return;}btn.disabled=true;btn.textContent='ĐANG TẢI VÀ TẠO BÁO CÁO...';status.style.display='block';});window.addEventListener('pageshow',function(){btn.disabled=false;btn.textContent='TẠO VÀ TẢI BÁO CÁO WORD';});})();</script>
</body></html>'''

def render_page(error=None, status=200):
    resp = make_response(render_template_string(HTML, error=error), status)
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    return resp

def valid_word_name(name):
    return Path(name or '').suffix.lower() in {'.doc', '.docx'}

def validate_xlsx(path):
    try:
        if path.stat().st_size < 1000:
            return False
        with zipfile.ZipFile(path, 'r') as z:
            names = set(z.namelist())
            return 'xl/workbook.xml' in names and '[Content_Types].xml' in names
    except (zipfile.BadZipFile, OSError):
        return False

def find_soffice():
    return shutil.which('libreoffice') or shutil.which('soffice')

def prepare_template(src: Path, workdir: Path) -> Path:
    ext = src.suffix.lower()
    dst = workdir / 'report_template.docx'
    if ext == '.docx':
        shutil.copy2(src, dst)
        return dst
    if ext != '.doc':
        raise ValueError('Báo cáo gốc phải là file .doc hoặc .docx.')
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError('Máy chủ chưa có bộ chuyển đổi Word .doc.')
    profile = workdir / 'lo_profile'
    profile.mkdir(exist_ok=True)
    cmd = [soffice, '--headless', f'-env:UserInstallation=file://{profile}', '--convert-to', 'docx:Office Open XML Text', '--outdir', str(workdir), str(src)]
    cp = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=120)
    candidates = [workdir / (src.stem + '.docx')]
    candidates += [p for p in workdir.glob('*.docx') if p not in candidates]
    converted = next((p for p in candidates if p.exists() and p.stat().st_size > 1000), None)
    if cp.returncode != 0 or converted is None:
        msg = (cp.stdout or '').strip()
        raise RuntimeError('Không chuyển được file .doc sang .docx.' + (f' Chi tiết: {msg[-300:]}' if msg else ''))
    if converted != dst:
        shutil.copy2(converted, dst)
    return dst

@app.after_request
def no_cache(resp):
    resp.headers.setdefault('Cache-Control', 'no-store')
    return resp

@app.get('/health')
def health():
    return {'ok': True, 'template_doc': bool(find_soffice()), 'template_docx': True, 'footnote_superscript': True}

@app.errorhandler(RequestEntityTooLarge)
def too_large(e):
    return render_page('Tổng dung lượng các tệp vượt quá 60 MB.', 413)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_page()
    try:
        ex = request.files.get('excel'); p7 = request.files.get('p7'); p9 = request.files.get('p9'); tpl = request.files.get('template')
        if not all([ex and ex.filename, p7 and p7.filename, p9 and p9.filename, tpl and tpl.filename]):
            return render_page('Vui lòng chọn đủ 4 tệp: Excel, Phòng 7, Phòng 9 và báo cáo gốc.', 400)
        if Path(ex.filename).suffix.lower() != '.xlsx':
            return render_page('Tệp thứ nhất phải là Excel .xlsx.', 400)
        if not valid_word_name(p7.filename) or not valid_word_name(p9.filename):
            return render_page('Báo cáo Phòng 7 và Phòng 9 phải là .doc hoặc .docx.', 400)
        if not valid_word_name(tpl.filename):
            return render_page('Báo cáo gốc phải là .doc hoặc .docx.', 400)
        with tempfile.TemporaryDirectory() as td0:
            td = Path(td0)
            ep = td / 'input.xlsx'
            p7p = td / ('p7' + Path(p7.filename).suffix.lower())
            p9p = td / ('p9' + Path(p9.filename).suffix.lower())
            tpl_src = td / ('template_source' + Path(tpl.filename).suffix.lower())
            out = td / 'Bao_cao_tuan_VKSND_tinh_Ninh_Binh.docx'
            ex.save(ep); p7.save(p7p); p9.save(p9p); tpl.save(tpl_src)
            if not validate_xlsx(ep):
                return render_page('Tệp Excel tải lên chưa đầy đủ hoặc không phải .xlsx chuẩn.', 400)
            if p7p.stat().st_size == 0 or p9p.stat().st_size == 0 or tpl_src.stat().st_size == 0:
                return render_page('Có tệp Word tải lên bị rỗng. Vui lòng chọn lại.', 400)
            template_docx = prepare_template(tpl_src, td)
            build_report(str(template_docx), str(ep), str(p7p), str(p9p), str(out))
            fix_footnote_format(out)
            if not out.exists() or out.stat().st_size < 1000:
                raise RuntimeError('Không tạo được file Word đầu ra.')
            data = out.read_bytes()
        resp = send_file(io.BytesIO(data), as_attachment=True, download_name='Bao_cao_tuan_VKSND_tinh_Ninh_Binh.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        resp.headers['Cache-Control'] = 'no-store'
        return resp
    except ClientDisconnected:
        return render_page('Kết nối bị ngắt trong lúc tải tệp. Vui lòng thử lại.', 400)
    except BadRequest:
        return render_page('Trình duyệt không gửi trọn vẹn dữ liệu tải lên. Vui lòng thử lại.', 400)
    except subprocess.TimeoutExpired:
        return render_page('Quá thời gian chuyển đổi file .doc. Vui lòng thử lại với file khác.', 500)
    except Exception as e:
        traceback.print_exc()
        return render_page('Lỗi xử lý: ' + str(e), 500)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', '10000')))
