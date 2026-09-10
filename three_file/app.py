import io
import os
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
BASE = Path(__file__).resolve().parent
TEMPLATE = BASE / 'report_template.docx'

HTML = '''<!doctype html>
<html lang="vi"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>Tạo báo cáo tuần VKSND tỉnh Ninh Bình</title>
<style>
*{box-sizing:border-box}body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;color:#1f2937}.wrap{max-width:780px;margin:24px auto;padding:0 14px}.card{background:#fff;border-radius:14px;box-shadow:0 8px 28px rgba(0,0,0,.08);padding:24px}.brand{font-weight:700;color:#991b1b;font-size:22px}.sub{color:#4b5563;margin:8px 0 20px;line-height:1.55}.field{margin:16px 0}.field label{display:block;font-weight:700;margin-bottom:7px}.field input{width:100%;border:1px solid #d1d5db;padding:12px;border-radius:8px;background:#fff;font-size:15px}.file-name{font-size:13px;color:#166534;margin-top:6px;min-height:16px;word-break:break-word}.hint{font-size:13px;color:#6b7280;margin-top:5px;line-height:1.45}.btn{margin-top:20px;width:100%;border:0;border-radius:9px;padding:15px 18px;min-height:52px;background:#991b1b;color:#fff;font-weight:700;font-size:16px;cursor:pointer;touch-action:manipulation}.btn:disabled{opacity:.65;cursor:wait}.status{display:none;margin-top:14px;padding:12px;border-radius:8px;background:#eff6ff;color:#1e3a8a;font-size:14px;line-height:1.45}.note{margin-top:18px;padding:12px;border-radius:8px;background:#fff7ed;font-size:13px;line-height:1.55}.err{background:#fef2f2;color:#991b1b;padding:12px;border-radius:8px;margin-bottom:16px;line-height:1.45}.ok{margin-top:12px;padding:11px;border-radius:8px;background:#ecfdf5;color:#166534;font-size:13px;line-height:1.5}@media(max-width:560px){.wrap{margin:10px auto}.card{padding:18px}.brand{font-size:20px}.btn{font-size:15px}}
</style></head><body><div class="wrap"><div class="card">
<div class="brand">TẠO BÁO CÁO TUẦN — BẢN 3 FILE</div>
<div class="sub">Chỉ cần tải 3 tệp: Excel tổng hợp, báo cáo Phòng 7 và báo cáo Phòng 9. Hệ thống dùng mẫu báo cáo Word chuẩn đã tích hợp sẵn, tự cập nhật thân báo cáo, hoạt động kiểm sát, footnote và phụ lục chỉ tiêu.</div>
{% if error %}<div class="err"><b>Không thể tạo báo cáo.</b><br>{{error}}</div>{% endif %}
<form id="reportForm" method="post" enctype="multipart/form-data" action="/">
<div class="field"><label>1. Bảng Excel báo cáo tuần</label><input id="excel" type="file" name="excel" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required><div id="excelName" class="file-name"></div><div class="hint">Dùng đúng bảng Excel tổng hợp của tuần cần lập báo cáo.</div></div>
<div class="field"><label>2. Báo cáo Phòng 7</label><input id="p7" type="file" name="p7" accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required><div id="p7Name" class="file-name"></div></div>
<div class="field"><label>3. Báo cáo Phòng 9</label><input id="p9" type="file" name="p9" accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required><div id="p9Name" class="file-name"></div></div>
<button id="submitBtn" class="btn" type="submit">TẠO VÀ TẢI BÁO CÁO WORD</button><div id="status" class="status">Đang đọc 3 tệp và tạo báo cáo Word. Vui lòng giữ nguyên trang này.</div>
</form>
<div class="ok"><b>Đã tích hợp:</b> mẫu Word chuẩn cố định; số liệu tự động điền màu đỏ; footnote dạng số mũ; cập nhật số liệu Phòng 7, Phòng 9; cập nhật phụ lục chỉ tiêu cuối báo cáo.</div>
<div class="note"><b>Nguyên tắc an toàn số liệu:</b> hệ thống chỉ điền nội dung có căn cứ từ 3 tệp đầu vào. Chỗ không có nguồn dữ liệu sẽ không tự tạo số liệu mới.</div>
</div></div>
<script>(function(){const ids=['excel','p7','p9'];ids.forEach(id=>{const el=document.getElementById(id),lab=document.getElementById(id+'Name');el.addEventListener('change',()=>lab.textContent=el.files&&el.files[0]?'Đã chọn: '+el.files[0].name:'')});const f=document.getElementById('reportForm'),b=document.getElementById('submitBtn'),s=document.getElementById('status');f.addEventListener('submit',e=>{const miss=ids.filter(id=>!document.getElementById(id).files.length);if(miss.length){e.preventDefault();s.style.display='block';s.textContent='Vui lòng chọn đủ 3 tệp.';return}b.disabled=true;b.textContent='ĐANG TẠO BÁO CÁO...';s.style.display='block'});window.addEventListener('pageshow',()=>{b.disabled=false;b.textContent='TẠO VÀ TẢI BÁO CÁO WORD'})})();</script></body></html>'''

def render_page(error=None, status=200):
    r = make_response(render_template_string(HTML, error=error), status)
    r.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    r.headers['Pragma'] = 'no-cache'
    return r

def valid_word_name(name):
    return Path(name or '').suffix.lower() in {'.doc', '.docx'}

def validate_xlsx(path):
    try:
        if path.stat().st_size < 1000:
            return False
        with zipfile.ZipFile(path, 'r') as z:
            n = set(z.namelist())
            return 'xl/workbook.xml' in n and '[Content_Types].xml' in n
    except Exception:
        return False

def validate_docx(path):
    try:
        with zipfile.ZipFile(path, 'r') as z:
            n = set(z.namelist())
            return 'word/document.xml' in n and 'word/styles.xml' in n and '[Content_Types].xml' in n
    except Exception:
        return False

@app.get('/health')
def health():
    return {
        'ok': TEMPLATE.exists() and validate_docx(TEMPLATE),
        'version': 'three-file-full-v1',
        'inputs': 3,
        'footnote_superscript': True,
        'embedded_template': True,
        'appendix_update': True,
    }

@app.errorhandler(RequestEntityTooLarge)
def too_large(e):
    return render_page('Tổng dung lượng 3 tệp vượt quá 60 MB.', 413)

@app.route('/', methods=['GET','POST'])
def index():
    if request.method == 'GET':
        return render_page()
    try:
        ex = request.files.get('excel'); p7 = request.files.get('p7'); p9 = request.files.get('p9')
        if not all([ex and ex.filename, p7 and p7.filename, p9 and p9.filename]):
            return render_page('Vui lòng chọn đủ 3 tệp: Excel, báo cáo Phòng 7 và báo cáo Phòng 9.', 400)
        if Path(ex.filename).suffix.lower() != '.xlsx':
            return render_page('Tệp thứ nhất phải là Excel định dạng .xlsx.', 400)
        if not valid_word_name(p7.filename):
            return render_page('Báo cáo Phòng 7 phải là .doc hoặc .docx.', 400)
        if not valid_word_name(p9.filename):
            return render_page('Báo cáo Phòng 9 phải là .doc hoặc .docx.', 400)
        if not TEMPLATE.exists() or not validate_docx(TEMPLATE):
            return render_page('Mẫu Word chuẩn trên máy chủ chưa sẵn sàng.', 500)

        with tempfile.TemporaryDirectory() as td0:
            td = Path(td0)
            ep = td / 'input.xlsx'
            p7p = td / ('p7' + Path(p7.filename).suffix.lower())
            p9p = td / ('p9' + Path(p9.filename).suffix.lower())
            out = td / 'Bao_cao_tuan_VKSND_tinh_Ninh_Binh.docx'
            ex.save(ep); p7.save(p7p); p9.save(p9p)
            if not validate_xlsx(ep):
                return render_page('File Excel không phải .xlsx chuẩn hoặc tải lên chưa đầy đủ.', 400)
            if p7p.stat().st_size == 0 or p9p.stat().st_size == 0:
                return render_page('Báo cáo Phòng 7 hoặc Phòng 9 bị rỗng.', 400)
            build_report(str(TEMPLATE), str(ep), str(p7p), str(p9p), str(out))
            fix_footnote_format(out)
            if not out.exists() or out.stat().st_size < 1000 or not validate_docx(out):
                raise RuntimeError('File Word đầu ra không hợp lệ.')
            data = out.read_bytes()

        return send_file(io.BytesIO(data), as_attachment=True,
            download_name='Bao_cao_tuan_VKSND_tinh_Ninh_Binh.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except ClientDisconnected:
        return render_page('Kết nối bị ngắt trong lúc tải tệp. Vui lòng thử lại.', 400)
    except BadRequest:
        return render_page('Trình duyệt không gửi trọn vẹn dữ liệu. Vui lòng thử lại.', 400)
    except Exception as e:
        traceback.print_exc()
        return render_page('Lỗi xử lý: ' + str(e), 500)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT','10000')))
