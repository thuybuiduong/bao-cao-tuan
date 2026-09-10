import os, tempfile, traceback, io
from pathlib import Path
from flask import Flask, request, send_file, render_template_string
from generator import build_report
from fallback_template import ensure_template

app=Flask(__name__)
app.config['MAX_CONTENT_LENGTH']=40*1024*1024
BASE=Path(__file__).resolve().parent
TEMPLATE=ensure_template(BASE/'report_template.docx')

HTML='''<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Tạo báo cáo tuần VKSND tỉnh Ninh Bình</title><style>body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;color:#1f2937}.wrap{max-width:760px;margin:42px auto;padding:0 16px}.card{background:#fff;border-radius:14px;box-shadow:0 8px 28px rgba(0,0,0,.08);padding:28px}.brand{font-weight:700;color:#991b1b;font-size:22px}.sub{color:#6b7280;margin:8px 0 24px;line-height:1.5}.field{margin:15px 0}.field label{display:block;font-weight:700;margin-bottom:7px}.field input{width:100%;box-sizing:border-box;border:1px solid #d1d5db;padding:12px;border-radius:8px}.hint{font-size:13px;color:#6b7280;margin-top:5px}.btn{margin-top:20px;width:100%;border:0;border-radius:9px;padding:13px 18px;background:#991b1b;color:#fff;font-weight:700;font-size:16px}.note{margin-top:18px;padding:12px;border-radius:8px;background:#fff7ed;font-size:13px;line-height:1.5}.err{background:#fef2f2;color:#991b1b;padding:12px;border-radius:8px;margin-bottom:16px}</style></head><body><div class="wrap"><div class="card"><div class="brand">TẠO BÁO CÁO TUẦN</div><div class="sub">VKSND tỉnh Ninh Bình — tải đồng thời Excel tổng hợp, báo cáo Phòng 7 và báo cáo Phòng 9. Hệ thống xuất Word, cập nhật footnote và tô đỏ số liệu tự động điền.</div>{% if error %}<div class="err">{{error}}</div>{% endif %}<form method="post" enctype="multipart/form-data"><div class="field"><label>1. Bảng Excel báo cáo tuần</label><input type="file" name="excel" accept=".xlsx" required></div><div class="field"><label>2. Báo cáo Phòng 7</label><input type="file" name="p7" accept=".doc,.docx" required></div><div class="field"><label>3. Báo cáo Phòng 9</label><input type="file" name="p9" accept=".doc,.docx" required><div class="hint">Chấp nhận cả .doc và .docx.</div></div><button class="btn" type="submit">TẠO VÀ TẢI BÁO CÁO WORD</button></form><div class="note"><b>Nguyên tắc:</b> số liệu/nội dung có nguồn từ 3 tệp được tự động cập nhật; số liệu điền vào hiển thị màu đỏ và các footnote tương ứng được tạo trong Word.</div></div></div></body></html>'''

@app.get('/health')
def health(): return {'ok':True}

@app.route('/',methods=['GET','POST'])
def index():
    if request.method=='GET': return render_template_string(HTML,error=None)
    try:
        ex=request.files.get('excel'); p7=request.files.get('p7'); p9=request.files.get('p9')
        if not ex or not p7 or not p9: return render_template_string(HTML,error='Vui lòng chọn đủ 3 tệp.'),400
        with tempfile.TemporaryDirectory() as td:
            td=Path(td); ep=td/'input.xlsx'; p7p=td/('p7'+Path(p7.filename).suffix.lower()); p9p=td/('p9'+Path(p9.filename).suffix.lower()); out=td/'Bao_cao_tuan_VKSND_tinh_Ninh_Binh.docx'
            ex.save(ep); p7.save(p7p); p9.save(p9p)
            build_report(str(TEMPLATE),str(ep),str(p7p),str(p9p),str(out)); data=out.read_bytes()
        return send_file(io.BytesIO(data),as_attachment=True,download_name='Bao_cao_tuan_VKSND_tinh_Ninh_Binh.docx',mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        traceback.print_exc(); return render_template_string(HTML,error='Không tạo được báo cáo: '+str(e)),500

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT','10000')))
