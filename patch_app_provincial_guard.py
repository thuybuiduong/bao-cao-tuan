from pathlib import Path

p=Path('/app/app.py')
s=p.read_text(encoding='utf-8')
needle="            if not valid_docx(out): raise RuntimeError('File Word đầu ra không hợp lệ.')\n"
insert="""            if not valid_docx(out): raise RuntimeError('File Word đầu ra không hợp lệ.')
            # Chốt bắt buộc: đầu ra phải vẫn là Báo cáo tuần VKSND tỉnh Ninh Bình.
            _d=Document(out)
            _t=' '.join((x.text or '') for x in _d.paragraphs)
            _required=('TỈNH NINH BÌNH','I. TÌNH HÌNH TỘI PHẠM VÀ VI PHẠM','II. CÔNG TÁC','III. CÔNG TÁC KHÁC','IV. NHIỆM VỤ')
            _missing=[x for x in _required if x not in _t]
            if _missing or len(_d.paragraphs) < 200 or len(_d.tables) < 3:
                raise RuntimeError('Đầu ra không còn đúng cấu trúc Báo cáo tuần VKSND tỉnh Ninh Bình; hệ thống đã chặn file sai.')
"""
if needle not in s:
    raise RuntimeError('Không tìm thấy điểm chèn provincial guard trong app.py')
p.write_text(s.replace(needle,insert,1),encoding='utf-8')
print('PROVINCIAL_OUTPUT_GUARD_INSTALLED')
