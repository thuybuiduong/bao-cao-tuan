import os, tempfile, traceback, io, zipfile
from pathlib import Path
from flask import Flask, request, send_file, render_template_string, make_response
from werkzeug.exceptions import RequestEntityTooLarge, ClientDisconnected, BadRequest
from generator import build_report
from fallback_template import ensure_template

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 40 * 1024 * 1024
BASE = Path(__file__).resolve().parent
TEMPLATE = ensure_template(BASE / 'report_template.docx')

HTML = '''<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Tạo báo cáo tuần VKSND tỉnh Ninh Bình</title>
<style>
*{box-sizing:border-box} body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;color:#1f2937}
.wrap{max-width:760px;margin:28px auto;padding:0 14px}.card{background:#fff;border-radius:14px;box-shadow:0 8px 28px rgba(0,0,0,.08);padding:24px}
.brand{font-weight:700;color:#991b1b;font-size:22px}.sub{color:#6b7280;margin:8px 0 22px;line-height:1.5}
.field{margin:15px 0}.field label{display:block;font-weight:700;margin-bottom:7px}.field input{width:100%;border:1px solid #d1d5db;padding:12px;border-radius:8px;background:#fff;font-size:15px}
.file-name{font-size:13px;color:#166534;margin-top:6px;min-height:16px;word-break:break-word}.hint{font-size:13px;color:#6b7280;margin-top:5px}
.btn{margin-top:20px;width:100%;border:0;border-radius:9px;padding:15px 18px;min-height:52px;background:#991b1b;color:#fff;font-weight:700;font-size:16px;cursor:pointer;touch-action:manipulation;-webkit-tap-highlight-color:transparent}
.btn:disabled{opacity:.65;cursor:wait}.status{display:none;margin-top:14px;padding:12px;border-radius:8px;background:#eff6ff;color:#1e3a8a;font-size:14px;line-height:1.45}
.note{margin-top:18px;padding:12px;border-radius:8px;background:#fff7ed;font-size:13px;line-height:1.5}.err{background:#fef2f2;color:#991b1b;padding:12px;border-radius:8px;margin-bottom:16px;line-height:1.45}
@media(max-width:560px){.wrap{margin:12px auto}.card{padding:18px}.brand{font-size:20px}.btn{font-size:15px}}
</style>
</head>
<body><div class="wrap"><div class="card">
<div class="brand">TẠO BÁO CÁO TUẦN</div>
<div class="sub">VKSND tỉnh Ninh Bình — tải đồng thời Excel tổng hợp, báo cáo Phòng 7 và báo cáo Phòng 9. Hệ thống xuất Word, cập nhật footnote và tô đỏ số liệu tự động điền.</div>
{% if error %}<div class="err"><b>Không thể tạo báo cáo.</b><br>{{ error }}</div>{% endif %}
<form id="reportForm" method="post" enctype="multipart/form-data" action="/">
<div class="field"><label>1. Bảng Excel báo cáo tuần</label><input id="excel" type="file" name="excel" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required><div id="excelName" class="file-name"></div></div>
<div class="field"><label>2. Báo cáo Phòng 7</label><input id="p7" type="file" name="p7" accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required><div id="p7Name" class="file-name"></div></div>
<div class="field"><label>3. Báo cáo Phòng 9</label><input id="p9" type="file" name="p9" accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required><div id="p9Name" class="file-name"></div><div class="hint">Chấp nhận cả .doc và .docx.</div></div>
<button id="submitBtn" class="btn" type="submit">TẠO VÀ TẢI BÁO CÁO WORD</button>
<div id="status" class="status">Đang tải tệp và tạo báo cáo Word. Vui lòng giữ nguyên trang này cho đến khi tệp Word bắt đầu tải xuống.</div>
</form>
<div class="note"><b>Nguyên tắc:</b> số liệu/nội dung có nguồn từ 3 tệp được tự động cập nhật; số liệu điền vào hiển thị màu đỏ và các footnote tương ứng được tạo trong Word.</div>
</div></div>
<script>
(function(){
  const ids=['excel','p7','p9'];
  ids.forEach(id=>{
    const el=document.getElementById(id), label=document.getElementById(id+'Name');
    el.addEventListener('change',()=>{ label.textContent=el.files && el.files[0] ? 'Đã chọn: '+el.files[0].name : ''; });
  });
  const form=document.getElementById('reportForm'), btn=document.getElementById('submitBtn'), status=document.getElementById('status');
  form.addEventListener('submit',function(e){
    const missing=ids.filter(id=>!document.getElementById(id).files.length);
    if(missing.length){e.preventDefault();status.style.display='block';status.textContent='Vui lòng chọn đủ 3 tệp trước khi tạo báo cáo.';return;}
    btn.disabled=true; btn.textContent='ĐANG TẢI VÀ TẠO BÁO CÁO...'; status.style.display='block';
  });
  window.addEventListener('pageshow',function(){btn.disabled=false;btn.textContent='TẠO VÀ TẢI BÁO CÁO WORD';});
})();
</script>
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


@app.after_request
def no_cache(resp):
    resp.headers.setdefault('Cache-Control', 'no-store')
    return resp


@app.get('/health')
def health():
    return {'ok': True, 'upload_ui': 'v2'}


@app.errorhandler(RequestEntityTooLarge)
def too_large(e):
    return render_page('Tổng dung lượng 3 tệp vượt quá 40 MB. Vui lòng kiểm tra lại tệp tải lên.', 413)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_page()
    try:
        ex = request.files.get('excel')
        p7 = request.files.get('p7')
        p9 = request.files.get('p9')
        if not ex or not ex.filename or not p7 or not p7.filename or not p9 or not p9.filename:
            return render_page('Vui lòng chọn đủ 3 tệp: Excel, Phòng 7 và Phòng 9.', 400)
        if Path(ex.filename).suffix.lower() != '.xlsx':
            return render_page('Tệp thứ nhất phải là Excel định dạng .xlsx.', 400)
        if not valid_word_name(p7.filename):
            return render_page('Báo cáo Phòng 7 phải là tệp .doc hoặc .docx.', 400)
        if not valid_word_name(p9.filename):
            return render_page('Báo cáo Phòng 9 phải là tệp .doc hoặc .docx.', 400)

        with tempfile.TemporaryDirectory() as td0:
            td = Path(td0)
            ep = td / 'input.xlsx'
            p7p = td / ('p7' + Path(p7.filename).suffix.lower())
            p9p = td / ('p9' + Path(p9.filename).suffix.lower())
            out = td / 'Bao_cao_tuan_VKSND_tinh_Ninh_Binh.docx'
            ex.save(ep); p7.save(p7p); p9.save(p9p)

            if not validate_xlsx(ep):
                return render_page('Tệp Excel tải lên chưa đầy đủ hoặc không phải tệp .xlsx chuẩn. Hãy chọn lại đúng file Excel và tải lại.', 400)
            if p7p.stat().st_size == 0 or p9p.stat().st_size == 0:
                return render_page('Một trong hai báo cáo Phòng 7/Phòng 9 tải lên bị rỗng. Hãy chọn lại tệp.', 400)

            build_report(str(TEMPLATE), str(ep), str(p7p), str(p9p), str(out))
            if not out.exists() or out.stat().st_size < 1000:
                raise RuntimeError('Không tạo được tệp Word đầu ra.')
            data = out.read_bytes()

        resp = send_file(io.BytesIO(data), as_attachment=True,
                         download_name='Bao_cao_tuan_VKSND_tinh_Ninh_Binh.docx',
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        resp.headers['Cache-Control'] = 'no-store'
        return resp
    except ClientDisconnected:
        return render_page('Kết nối bị ngắt trong lúc tải tệp. Vui lòng chọn lại 3 tệp và thử lại.', 400)
    except BadRequest as e:
        return render_page('Trình duyệt không gửi trọn vẹn dữ liệu tải lên. Vui lòng thử lại.', 400)
    except Exception as e:
        traceback.print_exc()
        return render_page('Lỗi xử lý: ' + str(e), 500)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', '10000')))
