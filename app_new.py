import os, tempfile, traceback, io, zipfile
from pathlib import Path
from flask import Flask, request, send_file, render_template_string, make_response
from werkzeug.exceptions import RequestEntityTooLarge, ClientDisconnected, BadRequest
from generator import build_report
from fallback_template import ensure_template
from footnote_fix import fix_footnote_format

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 60 * 1024 * 1024
BASE = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = ensure_template(BASE / 'report_template.docx')

HTML = '''<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><title>Tạo báo cáo tuần</title><style>
*{box-sizing:border-box}body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;color:#1f2937}.wrap{max-width:760px;margin:24px auto;padding:0 14px}.card{background:#fff;border-radius:14px;box-shadow:0 8px 28px rgba(0,0,0,.08);padding:24px}.brand{font-weight:700;color:#991b1b;font-size:22px}.sub{color:#6b7280;margin:8px 0 22px;line-height:1.5}.field{margin:15px 0}.field label{display:block;font-weight:700;margin-bottom:7px}.field input{width:100%;border:1px solid #d1d5db;padding:12px;border-radius:8px;background:#fff;font-size:15px}.file-name{font-size:13px;color:#166534;margin-top:6px;min-height:16px;word-break:break-word}.hint{font-size:13px;color:#6b7280;margin-top:5px}.btn{margin-top:20px;width:100%;border:0;border-radius:9px;padding:15px 18px;min-height:52px;background:#991b1b;color:#fff;font-weight:700;font-size:16px;cursor:pointer}.btn:disabled{opacity:.65;cursor:wait}.status{display:none;margin-top:14px;padding:12px;border-radius:8px;background:#eff6ff;color:#1e3a8a;font-size:14px}.note{margin-top:18px;padding:12px;border-radius:8px;background:#fff7ed;font-size:13px;line-height:1.5}.err{background:#fef2f2;color:#991b1b;padding:12px;border-radius:8px;margin-bottom:16px;line-height:1.45}@media(max-width:560px){.wrap{margin:10px auto}.card{padding:18px}}
</style></head><body><div class="wrap"><div class="card"><div class="brand">TẠO BÁO CÁO TUẦN</div><div class="sub">Hệ thống điền số liệu vào đúng báo cáo Word gốc, giữ nguyên căn lề, căn đoạn, font, giãn dòng, tab và định dạng footnote.</div>{% if error %}<div class="err"><b>Không thể tạo báo cáo.</b><br>{{error}}</div>{% endif %}<form id="reportForm" method="post" enctype="multipart/form-data">
<div class="field"><label>1. Bảng Excel báo cáo tuần</label><input id="excel" type="file" name="excel" accept=".xlsx" required><div id="excelName" class="file-name"></div></div>
<div class="field"><label>2. Báo cáo Phòng 7</label><input id="p7" type="file" name="p7" accept=".doc,.docx" required><div id="p7Name" class="file-name"></div></div>
<div class="field"><label>3. Báo cáo Phòng 9</label><input id="p9" type="file" name="p9" accept=".doc,.docx" required><div id="p9Name" class="file-name"></div></div>
<div class="field"><label>4. Báo cáo Word gốc cần giữ nguyên form</label><input id="template" type="file" name="template" accept=".docx" required><div id="templateName" class="file-name"></div><div class="hint">Chọn chính báo cáo Word gốc/mẫu của tuần. Hệ thống chỉ thay dữ liệu cần cập nhật, không dựng lại định dạng.</div></div>
<button id="submitBtn" class="btn" type="submit">TẠO VÀ TẢI BÁO CÁO WORD</button><div id="status" class="status">Đang tải và tạo báo cáo. Vui lòng giữ nguyên trang này.</div></form><div class="note"><b>Đã sửa:</b> số footnote được ép về kiểu FootnoteReference và hiển thị dạng số mũ; báo cáo đầu ra sử dụng trực tiếp file Word gốc bạn tải lên để giữ form.</div></div></div>
<script>(function(){const ids=['excel','p7','p9','template'];ids.forEach(id=>{const el=document.getElementById(id),lab=document.getElementById(id+'Name');el.addEventListener('change',()=>lab.textContent=el.files&&el.files[0]?'Đã chọn: '+el.files[0].name:'')});const f=document.getElementById('reportForm'),b=document.getElementById('submitBtn'),s=document.getElementById('status');f.addEventListener('submit',e=>{const miss=ids.filter(id=>!document.getElementById(id).files.length);if(miss.length){e.preventDefault();s.style.display='block';s.textContent='Vui lòng chọn đủ 4 tệp.';return}b.disabled=true;b.textContent='ĐANG TẢI VÀ TẠO BÁO CÁO...';s.style.display='block'});window.addEventListener('pageshow',()=>{b.disabled=false;b.textContent='TẠO VÀ TẢI BÁO CÁO WORD'})})();</script></body></html>'''

def render_page(error=None,status=200):
    r=make_response(render_template_string(HTML,error=error),status)
    r.headers['Cache-Control']='no-store, no-cache, must-revalidate, max-age=0'
    return r

def validate_xlsx(path):
    try:
        with zipfile.ZipFile(path,'r') as z:
            n=set(z.namelist()); return 'xl/workbook.xml' in n and '[Content_Types].xml' in n
    except Exception: return False

def validate_docx(path):
    try:
        with zipfile.ZipFile(path,'r') as z:
            n=set(z.namelist()); return 'word/document.xml' in n and 'word/styles.xml' in n and '[Content_Types].xml' in n
    except Exception: return False

@app.get('/health')
def health(): return {'ok':True,'upload_ui':'v3-exact-template-footnote'}

@app.errorhandler(RequestEntityTooLarge)
def too_large(e): return render_page('Tổng dung lượng tệp vượt quá 60 MB.',413)

@app.route('/',methods=['GET','POST'])
def index():
    if request.method=='GET': return render_page()
    try:
        ex=request.files.get('excel'); p7=request.files.get('p7'); p9=request.files.get('p9'); src=request.files.get('template')
        if not all([ex and ex.filename,p7 and p7.filename,p9 and p9.filename,src and src.filename]):
            return render_page('Vui lòng chọn đủ 4 tệp: Excel, Phòng 7, Phòng 9 và báo cáo Word gốc.',400)
        if Path(ex.filename).suffix.lower()!='.xlsx': return render_page('Tệp thứ nhất phải là .xlsx.',400)
        if Path(p7.filename).suffix.lower() not in {'.doc','.docx'}: return render_page('Báo cáo Phòng 7 phải là .doc hoặc .docx.',400)
        if Path(p9.filename).suffix.lower() not in {'.doc','.docx'}: return render_page('Báo cáo Phòng 9 phải là .doc hoặc .docx.',400)
        if Path(src.filename).suffix.lower()!='.docx': return render_page('Báo cáo Word gốc phải là .docx để giữ nguyên định dạng.',400)
        with tempfile.TemporaryDirectory() as td0:
            td=Path(td0); ep=td/'input.xlsx'; p7p=td/('p7'+Path(p7.filename).suffix.lower()); p9p=td/('p9'+Path(p9.filename).suffix.lower()); tp=td/'template.docx'; out=td/'Bao_cao_tuan_VKSND_tinh_Ninh_Binh.docx'
            ex.save(ep); p7.save(p7p); p9.save(p9p); src.save(tp)
            if not validate_xlsx(ep): return render_page('File Excel không phải .xlsx chuẩn hoặc tải lên chưa đầy đủ.',400)
            if not validate_docx(tp): return render_page('File báo cáo Word gốc không phải .docx chuẩn.',400)
            build_report(str(tp),str(ep),str(p7p),str(p9p),str(out))
            fix_footnote_format(out)
            if not validate_docx(out): raise RuntimeError('File Word đầu ra không hợp lệ.')
            data=out.read_bytes()
        return send_file(io.BytesIO(data),as_attachment=True,download_name='Bao_cao_tuan_VKSND_tinh_Ninh_Binh.docx',mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except ClientDisconnected:
        return render_page('Kết nối bị ngắt khi tải tệp. Vui lòng thử lại.',400)
    except BadRequest:
        return render_page('Trình duyệt không gửi trọn vẹn dữ liệu tải lên. Vui lòng thử lại.',400)
    except Exception as e:
        traceback.print_exc(); return render_page('Lỗi xử lý: '+str(e),500)

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT','10000')))
