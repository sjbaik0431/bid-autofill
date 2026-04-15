# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════╗
║   입찰 정량평가 자동입력 대시보드 v4             ║
║   파일 업로드/삭제 + 회사정보 자동추출 지원      ║
╚══════════════════════════════════════════════════╝
실행: python dashboard.py → 브라우저에서 http://localhost:5000
"""
from flask import Flask, jsonify, request, send_file
import win32com.client as win32
import json, os, shutil, threading, time, re, glob
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# company_master.json 탐색: 스크립트 디렉토리 → 상위 디렉토리 순
_master_local = os.path.join(SCRIPT_DIR, "company_master.json")
_master_parent = os.path.join(os.path.dirname(SCRIPT_DIR), "company_master.json")
MASTER_DB = _master_local if os.path.exists(_master_local) else (_master_parent if os.path.exists(_master_parent) else _master_local)
_config_local = os.path.join(SCRIPT_DIR, "autofill_config.json")
_config_parent = os.path.join(os.path.dirname(SCRIPT_DIR), "autofill_config.json")
CONFIG_FILE = _config_local if os.path.exists(_config_local) else (_config_parent if os.path.exists(_config_parent) else _config_local)
FORMS_DIR = os.path.join(SCRIPT_DIR, "양식")       # 양식 파일 저장소
COMPANY_DIR = os.path.join(SCRIPT_DIR, "회사정보")  # 회사정보 파일 저장소

os.makedirs(FORMS_DIR, exist_ok=True)
os.makedirs(COMPANY_DIR, exist_ok=True)

# 실행 상태 추적
status = {
    "running": False, "progress": 0, "total": 0,
    "current_task": "", "log": [], "result": None
}

def load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_json(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# ── HWP 텍스트 추출 (olefile + zlib, record-level parsing) ──
def extract_hwp_text(hwp_path):
    """HWP 파일에서 텍스트 추출 - HWP record 구조를 파싱하여 PARA_TEXT만 추출"""
    import olefile, zlib, struct
    text_parts = []
    try:
        if not os.path.exists(hwp_path):
            return f"[추출 오류: 파일이 존재하지 않습니다 - {hwp_path}]"
        ole = olefile.OleFileIO(hwp_path)
        streams = ole.listdir()
        body_streams = sorted(["/".join(s) for s in streams if s[0] == "BodyText"])
        if not body_streams:
            ole.close()
            return "[추출 오류: HWP 파일에 BodyText 스트림이 없습니다. 올바른 HWP 파일인지 확인하세요.]"
        for stream_path in body_streams:
            data = ole.openstream(stream_path).read()
            try:
                data = zlib.decompress(data, -15)
            except:
                try:
                    data = zlib.decompress(data)
                except:
                    pass  # use raw data
            # Parse HWP record structure to find PARA_TEXT (tag 0x0043) records
            offset = 0
            while offset + 4 <= len(data):
                header_val = struct.unpack_from('<I', data, offset)[0]
                tag_id = header_val & 0x3FF
                size = (header_val >> 20) & 0xFFF
                offset += 4
                if size == 0xFFF:
                    if offset + 4 > len(data):
                        break
                    size = struct.unpack_from('<I', data, offset)[0]
                    offset += 4
                if offset + size > len(data):
                    break
                if tag_id == 0x0043:  # HWPTAG_PARA_TEXT
                    rec = data[offset:offset + size]
                    text_parts.append(_extract_para_text(rec))
                offset += size
        ole.close()
    except Exception as e:
        return f"[추출 오류: {type(e).__name__} - {e}]"
    return "\n".join(text_parts)

def _extract_para_text(rec):
    """PARA_TEXT 레코드에서 텍스트 추출 (HWP 제어코드 필터링)"""
    import struct
    pos = 0
    chars = []
    while pos + 1 < len(rec):
        cc = struct.unpack_from('<H', rec, pos)[0]
        if cc == 0x0000:
            pos += 2
        elif cc <= 0x0002:
            # Inline control chars: 2 + 14 = 16 bytes
            pos += 16
        elif 0x0003 <= cc <= 0x0008 or 0x000B <= cc <= 0x000C or 0x000E <= cc <= 0x001F:
            # Extended control chars: 2 + 14 = 16 bytes
            pos += 16
        elif cc == 0x0009:
            chars.append('\t')
            pos += 2
        elif cc == 0x000A:
            chars.append('\n')
            pos += 2
        elif cc == 0x000D:
            chars.append('\n')
            pos += 2
        else:
            # Normal char - filter HWP internal fill codes
            if not (cc == 0x0100 or 0x0E00 <= cc <= 0x0FFF):
                chars.append(chr(cc))
            pos += 2
    return "".join(chars).strip()

def parse_company_info_from_text(text):
    """추출된 텍스트에서 회사정보 파싱 - 같은줄/다음줄(표) 형식 모두 지원"""
    info = {}
    lines = [l.strip() for l in text.split('\n')]

    # ── 1단계: "라벨 : 값" 같은줄 패턴 (값이 있는 것만) ──
    same_line = {
        '업체명': [r'업\s*체\s*명\s*[:：]\s*(.+)', r'회\s*사\s*명\s*[:：]\s*(.+)', r'상\s*호\s*[:：]\s*(.+)'],
        '대표자': [r'대\s*표\s*자\s*[:：]\s*(.+)', r'대\s*표\s*이\s*사\s*[:：]\s*(.+)'],
        '사업자번호': [r'사\s*업\s*자\s*등?\s*록?\s*번\s*호\s*[:：]\s*(.+)'],
        '주소': [r'주\s*소\s*[:：]\s*(.+)', r'소\s*재\s*지\s*[:：]\s*(.+)'],
        '전화번호': [r'전\s*화\s*번?\s*호?\s*[:：]\s*(.+)', r'TEL\s*[:：]\s*(.+)'],
        'FAX': [r'FAX\s*[:：]\s*(.+)', r'팩\s*스\s*[:：]\s*(.+)'],
        '설립일': [r'설\s*립\s*일\s*[:：]\s*(.+)'],
        '자본금': [r'자\s*본\s*금\s*[:：]\s*(.+)'],
        '전년도매출액': [r'전\s*년\s*도?\s*매\s*출\s*액?\s*[:：]\s*(.+)', r'매\s*출\s*액\s*[:：]\s*(.+)'],
    }
    for key, pats in same_line.items():
        for pat in pats:
            for line in lines:
                m = re.match(pat, line)
                if m:
                    val = m.group(1).strip()
                    # (인), (서명) 등 불필요 접미어 제거
                    val = re.sub(r'\s*[\(（]인[\)）]\s*$', '', val).strip()
                    val = re.sub(r'\s*[\(（]서명[\)）]\s*$', '', val).strip()
                    val = re.split(r'\s{3,}|\t', val)[0].strip()
                    if val and len(val) > 0 and len(val) < 150:
                        info[key] = val
                        break
            if key in info:
                break

    # ── 2단계: 표 형식 (라벨이 한 줄, 값이 다음 줄) ── 서식 제12호 등
    next_line = {
        '업체명': [r'^업\s*체\s*명$', r'^회\s*사\s*명$', r'^상\s*호$'],
        '대표자': [r'^대\s*표\s*자$', r'^대\s*표\s*이\s*사$'],
        '사업자번호': [r'^사\s*업\s*자\s*번\s*호$', r'^사\s*업\s*자\s*등\s*록\s*번\s*호$'],
        '주소': [r'^주\s*소$', r'^소\s*재\s*지$'],
        '전화번호': [r'^전\s*화\s*번?\s*호?$', r'^TEL$'],
        'FAX': [r'^FAX$', r'^팩\s*스$'],
        '설립일': [r'^설\s*립\s*일$', r'^설\s*립\s*년\s*도$'],
        '자본금': [r'^자\s*본\s*금$', r'^납\s*입\s*자\s*본\s*금$'],
        '전년도매출액': [r'^전\s*년\s*도?\s*매\s*출\s*액?$', r'^매\s*출\s*액$'],
    }
    # 라벨형 placeholder 패턴 (다음 줄에 값이 아니라 플레이스홀더가 있는 경우)
    placeholder_pat = re.compile(r'^[\(（][가-힣\s]+[\)）]$')

    for key, pats in next_line.items():
        if key in info:
            continue  # 이미 1단계에서 추출됨
        for pat in pats:
            for i, line in enumerate(lines):
                if re.match(pat, line, re.IGNORECASE) and i + 1 < len(lines):
                    val = lines[i + 1].strip()
                    # 다음 줄이 다른 라벨이거나 비어있으면 스킵
                    if not val or re.match(r'^[가-힣\s]{2,}[:：]', val) or len(val) > 150:
                        continue
                    # (휴대폰번호), (연락처) 등 placeholder 스킵 - 실제 값이 아님
                    if placeholder_pat.match(val):
                        continue
                    # 전화번호 특수 검증: 숫자/하이픈 포함 여부 확인
                    if key in ('전화번호', 'FAX') and not re.search(r'\d', val):
                        continue
                    # 사업자번호 검증: 숫자-숫자-숫자 형식이어야 함
                    if key == '사업자번호' and not re.search(r'\d{3}-\d{2}-\d{5}', val):
                        continue
                    # (인), (서명) 제거
                    val = re.sub(r'\s*[\(（]인[\)）]\s*$', '', val).strip()
                    val = re.sub(r'\s*[\(（]서명[\)）]\s*$', '', val).strip()
                    # 다중 공백을 단일 공백으로 정규화 (예: "2013년    9월    3일")
                    val = re.sub(r'\s+', ' ', val).strip()
                    if val and len(val) > 0:
                        info[key] = val
                        break
            if key in info:
                break

    # ── 3단계: 사업자번호 특수 패턴 (숫자-숫자-숫자) ──
    if '사업자번호' not in info:
        for line in lines:
            m = re.search(r'(\d{3}-\d{2}-\d{5})', line)
            if m:
                info['사업자번호'] = m.group(1)
                break

    # ── 4단계: 인력현황, 설립일 등 보조 추출 ──
    if '설립일' not in info:
        for line in lines:
            m = re.search(r'설\s*립\s*일\s*[:：]?\s*(\d{4}년\s*\d{1,2}월\s*\d{1,2}일)', line)
            if m:
                info['설립일'] = m.group(1).strip()
                break
        # 표 형식에서 다음 줄 확인
        if '설립일' not in info:
            for i, line in enumerate(lines):
                if re.match(r'^설\s*립\s*일$', line) and i + 1 < len(lines):
                    val = lines[i + 1].strip()
                    m2 = re.search(r'(\d{4}년.*)', val)
                    if m2:
                        info['설립일'] = m2.group(1).strip()
                        break

    return info

# ── 한컴 COM 찾아바꾸기 ──
def replace_all(hwp, find, replace):
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = find
    hwp.HParameterSet.HFindReplace.ReplaceString = replace
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.Direction = 0
    hwp.HParameterSet.HFindReplace.FindType = 0
    return hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

try:
    from autofill import build_patterns as _autofill_build_patterns
    def build_patterns(info, bid_info, extended=None):
        """autofill.py의 풍부한 패턴(인력/면허/실적/추가정보/일반데이터 포함)을 사용"""
        return _autofill_build_patterns(info, bid_info, extended)
except Exception as _e:
    print(f"[경고] autofill.py 로드 실패, 기본 패턴만 사용: {_e}")
    def build_patterns(info, bid_info, extended=None):
        업체명 = info.get('업체명', ''); 대표자 = info.get('대표자', '')
        주소 = info.get('주소', ''); 사번 = info.get('사업자번호', '')
        전화 = info.get('전화번호', ''); 팩스 = info.get('FAX', '')
        자본금 = info.get('자본금', ''); 매출액 = info.get('전년도매출액', '')
        설립일 = info.get('설립일', '')
        입찰명 = bid_info.get('입찰명', ''); 발주처 = bid_info.get('발주처', '')
        patterns = []
        for label in ["업 체 명 :", "업체명 :", "회사명 :", "업체명:"]:
            patterns.append(("업체명", label, f"{label} {업체명}"))
        for label in ["대 표 자 :", "대표자 :", "대표이사 :"]:
            patterns.append(("대표자", label, f"{label} {대표자}"))
        for label in ["주    소 :", "주소 :", "소재지 :"]:
            patterns.append(("주소", label, f"{label} {주소}"))
        for label in ["사업자등록번호 :", "사업자번호 :"]:
            patterns.append(("사업자번호", label, f"{label} {사번}"))
        for label in ["전화번호 :", "전 화 번 호 :", "전화 :"]:
            patterns.append(("전화번호", label, f"{label} {전화}"))
        for label in ["FAX :", "팩스 :"]:
            patterns.append(("FAX", label, f"{label} {팩스}"))
        for label in ["자본금 :"]:
            patterns.append(("자본금", label, f"{label} {자본금}"))
        for label in ["전년도매출액 :", "매출액 :"]:
            patterns.append(("매출액", label, f"{label} {매출액}"))
        for label in ["설립일 :", "설 립 일 :"]:
            patterns.append(("설립일", label, f"{label} {설립일}"))
        patterns.append(("소속", "소속 :", f"소속 : {업체명}"))
        patterns.append(("직위", "직위 :", f"직위 : 대표이사"))
        patterns.append(("성명", "성명 :", f"성명 : {대표자}"))
        if 입찰명:
            for label in ["용역명 :", "사업명 :", "계약건명 :"]:
                patterns.append(("입찰명", label, f"{label} {입찰명}"))
        if 발주처:
            for label in ["발주처 :", "발주기관명 :"]:
                patterns.append(("발주처", label, f"{label} {발주처}"))
        return patterns

def run_autofill(form_path, output_name, bid_info, demo_mode=False):
    global status
    status = {"running": True, "progress": 0, "total": 0, "current_task": "준비 중...", "log": [], "result": None}
    try:
        db = load_json(MASTER_DB)
        info = db.get("회사정보", {})
        # 확장 정보 로드 (인력/면허/실적/추가정보/연혁/강점/일반데이터)
        extended = {
            '인력': db.get('인력_전체', []),
            '면허': db.get('면허_허가_등록증', []),
            '실적': db.get('사업수행실적', []),
            '추가정보': db.get('추가정보', {}),
            '연혁': db.get('연혁', []),
            '강점': db.get('강점', []),
            '일반데이터': db.get('일반데이터', {})
        }
        out_dir = SCRIPT_DIR
        output_hwp = os.path.join(out_dir, f"{output_name}.hwp")
        output_pdf = os.path.join(out_dir, f"{output_name}.pdf")
        for f in [output_hwp, output_pdf]:
            if os.path.exists(f):
                try: os.remove(f)
                except: pass
        status["current_task"] = "양식 복사 중..."
        status["log"].append("📄 양식 파일 복사")
        shutil.copy2(form_path, output_hwp)
        time.sleep(0.3)
        status["current_task"] = "한컴오피스 실행 중..."
        status["log"].append("🖥️ 한컴오피스 실행")
        hwp = win32.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.XHwpWindows.Item(0).Visible = True
        if demo_mode: time.sleep(2)
        status["current_task"] = "파일 열기 중..."
        status["log"].append("📂 파일 열기")
        hwp.Open(output_hwp, "HWP", "forceopen:true")
        if demo_mode: time.sleep(2)
        patterns = build_patterns(info, bid_info, extended)
        year = str(datetime.now().year)
        제출일 = bid_info.get('제출일', f"{year}년  {datetime.now().month}월  {datetime.now().day}일")
        date_patterns = [
            (f"{year}년    월    일", 제출일), (f"{year}년   월    일", 제출일),
            (f"{year}년   월   일", 제출일), (f"{year}년  월  일", 제출일),
        ]
        total = len(patterns) + len(date_patterns)
        status["total"] = total
        ok_count = 0
        status["current_task"] = "자동입력 진행 중..."
        status["log"].append(f"✏️ 자동입력 시작 (패턴 {total}개)")
        for i, (category, find, rep) in enumerate(patterns):
            if demo_mode: hwp.HAction.Run("MoveDocBegin")
            try:
                replace_all(hwp, find, rep); ok_count += 1
                status["log"].append(f"  ✓ {category}: {find}")
            except: pass
            status["progress"] = i + 1
            if demo_mode: time.sleep(1.0)
        for i, (find, rep) in enumerate(date_patterns):
            if demo_mode: hwp.HAction.Run("MoveDocBegin")
            try:
                replace_all(hwp, find, rep); ok_count += 1
                status["log"].append(f"  ✓ 날짜: {rep}")
            except: pass
            status["progress"] = len(patterns) + i + 1
            if demo_mode: time.sleep(1.0)
        status["current_task"] = "저장 중..."
        status["log"].append("💾 HWP 저장")
        hwp.SaveAs(output_hwp, "HWP")
        pdf_ok = False
        try:
            hwp.SaveAs(output_pdf, "PDF"); pdf_ok = True
            status["log"].append("💾 PDF 변환 완료")
        except: status["log"].append("⚠️ PDF 변환 실패")
        if demo_mode: time.sleep(1)
        hwp.Clear(1); hwp.Quit()
        status["current_task"] = "완료!"
        status["log"].append(f"✅ 완료! (성공 {ok_count}개)")
        status["result"] = {
            "success": True, "ok_count": ok_count,
            "hwp_path": output_hwp, "pdf_path": output_pdf if pdf_ok else None,
            "hwp_exists": os.path.exists(output_hwp), "pdf_exists": os.path.exists(output_pdf)
        }
        status["progress"] = total
    except Exception as e:
        status["current_task"] = f"오류: {str(e)}"
        status["log"].append(f"❌ 오류: {str(e)}")
        status["result"] = {"success": False, "error": str(e)}
    finally:
        status["running"] = False

# ══════════════════════════════════════
#   API 라우트
# ══════════════════════════════════════

from flask import make_response

@app.route('/')
def index():
    """
    외부 index.html 우선 로드 (소스 단일화).
    파일이 없으면 내장 HTML_PAGE로 폴백.
    브라우저 캐시를 방지하기 위한 헤더 설정 포함.
    """
    ext_html = os.path.join(SCRIPT_DIR, 'index.html')
    html_content = HTML_PAGE
    if os.path.exists(ext_html):
        try:
            with open(ext_html, 'r', encoding='utf-8') as f:
                html_content = f.read()
        except Exception as e:
            print(f"[경고] index.html 읽기 실패, 내장 페이지 사용: {e}")

    resp = make_response(html_content)
    # 브라우저 캐시 방지 (코드 업데이트 시 즉시 반영되도록)
    resp.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp

# ── 회사정보 API ──
@app.route('/api/company', methods=['GET'])
def get_company():
    try: return jsonify(load_json(MASTER_DB))
    except: return jsonify({"error": "company_master.json 없음"}), 404

@app.route('/api/company', methods=['POST'])
def save_company():
    save_json(MASTER_DB, request.json)
    return jsonify({"ok": True})

@app.route('/api/company/import', methods=['POST'])
def import_company():
    """회사정보 HWP 파일 업로드 → 텍스트 추출 → 회사정보 파싱"""
    if 'file' not in request.files:
        return jsonify({"error": "파일이 없습니다"}), 400
    f = request.files['file']
    if not f.filename.lower().endswith('.hwp'):
        return jsonify({"error": "HWP 파일만 업로드 가능합니다"}), 400
    try:
        # 한글 파일명 안전 처리
        safe_name = f.filename.replace('/', '_').replace('\\', '_')
        os.makedirs(COMPANY_DIR, exist_ok=True)
        save_path = os.path.join(COMPANY_DIR, safe_name)
        f.save(save_path)
    except Exception as e:
        return jsonify({"error": f"파일 저장 실패: {e}"}), 500
    # 텍스트 추출
    try:
        import olefile
    except ImportError:
        return jsonify({"error": "olefile 패키지가 설치되지 않았습니다. pip install olefile 실행 후 재시도하세요."}), 500
    text = extract_hwp_text(save_path)
    if text.startswith("[추출 오류"):
        return jsonify({"error": text, "text": ""}), 500
    if not text.strip():
        return jsonify({"error": "파일에서 텍스트를 추출할 수 없습니다. HWP 파일이 맞는지 확인하세요."}), 500
    # 파싱
    parsed = parse_company_info_from_text(text)
    if not parsed:
        return jsonify({"error": "회사정보 항목을 찾을 수 없습니다. 파일 내용을 확인하세요.", "text_preview": text[:1000]}), 400
    return jsonify({
        "ok": True,
        "parsed": parsed,
        "text_preview": text[:3000],
        "filename": safe_name,
        "saved_path": save_path
    })

@app.route('/api/company/files', methods=['GET'])
def list_company_files():
    """업로드된 회사정보 파일 목록"""
    files = []
    for f in os.listdir(COMPANY_DIR):
        fp = os.path.join(COMPANY_DIR, f)
        if os.path.isfile(fp):
            files.append({
                "name": f,
                "path": fp,
                "size": os.path.getsize(fp),
                "modified": datetime.fromtimestamp(os.path.getmtime(fp)).strftime('%Y-%m-%d %H:%M')
            })
    files.sort(key=lambda x: x['modified'], reverse=True)
    return jsonify(files)

@app.route('/api/company/files/delete', methods=['POST'])
def delete_company_file():
    path = request.json.get("path", "")
    if os.path.exists(path) and COMPANY_DIR in path:
        os.remove(path)
        return jsonify({"ok": True})
    return jsonify({"error": "파일 없음"}), 404

# ── 양식 파일 API ──
@app.route('/api/forms', methods=['GET'])
def list_forms():
    """업로드된 양식 목록"""
    files = []
    for f in os.listdir(FORMS_DIR):
        fp = os.path.join(FORMS_DIR, f)
        if os.path.isfile(fp) and f.lower().endswith('.hwp'):
            files.append({
                "name": f,
                "path": fp,
                "size": os.path.getsize(fp),
                "size_str": f"{os.path.getsize(fp) / 1024:.0f}KB",
                "modified": datetime.fromtimestamp(os.path.getmtime(fp)).strftime('%Y-%m-%d %H:%M')
            })
    files.sort(key=lambda x: x['modified'], reverse=True)
    return jsonify(files)

@app.route('/api/forms/upload', methods=['POST'])
def upload_form():
    """양식 HWP 파일 업로드"""
    if 'file' not in request.files:
        return jsonify({"error": "파일이 없습니다"}), 400
    f = request.files['file']
    if not f.filename.lower().endswith('.hwp'):
        return jsonify({"error": "HWP 파일만 업로드 가능합니다"}), 400
    save_path = os.path.join(FORMS_DIR, f.filename)
    f.save(save_path)
    return jsonify({"ok": True, "name": f.filename, "path": save_path,
                     "size_str": f"{os.path.getsize(save_path)/1024:.0f}KB"})

@app.route('/api/forms/delete', methods=['POST'])
def delete_form():
    """양식 파일 삭제"""
    path = request.json.get("path", "")
    if os.path.exists(path) and FORMS_DIR in path:
        os.remove(path)
        return jsonify({"ok": True})
    return jsonify({"error": "파일 없음"}), 404

# ── 설정/실행 API ──
@app.route('/api/config', methods=['GET'])
def get_config():
    try: return jsonify(load_json(CONFIG_FILE))
    except: return jsonify({}), 404

@app.route('/api/config', methods=['POST'])
def save_config():
    save_json(CONFIG_FILE, request.json)
    return jsonify({"ok": True})

@app.route('/api/run', methods=['POST'])
def run():
    if status["running"]:
        return jsonify({"error": "이미 실행 중입니다"}), 400
    data = request.json
    form_path = data.get("form_path", "")
    if not os.path.exists(form_path):
        return jsonify({"error": f"양식 파일을 찾을 수 없습니다: {form_path}"}), 400
    t = threading.Thread(target=run_autofill, args=(
        form_path, data.get("output_name", "자동입력완료"),
        data.get("bid_info", {}), data.get("demo_mode", False)))
    t.daemon = True; t.start()
    return jsonify({"ok": True})

@app.route('/api/status', methods=['GET'])
def get_status():
    return jsonify(status)

@app.route('/api/open-file', methods=['POST'])
def open_file():
    path = request.json.get("path", "")
    if os.path.exists(path): os.startfile(path); return jsonify({"ok": True})
    return jsonify({"error": "파일 없음"}), 404

@app.route('/api/open-folder', methods=['POST'])
def open_folder():
    os.startfile(SCRIPT_DIR)
    return jsonify({"ok": True})

# ══════════════════════════════════════
#   HTML 대시보드
# ══════════════════════════════════════
HTML_PAGE = r'''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>입찰 정량평가 자동입력 시스템</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&display=swap');
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:'Noto Sans KR',sans-serif; background:linear-gradient(135deg,#0f172a 0%,#1e293b 50%,#0f172a 100%); min-height:100vh; color:#e2e8f0; }
  .header { background:rgba(30,41,59,.8); backdrop-filter:blur(20px); border-bottom:1px solid rgba(99,102,241,.3); padding:20px 40px; display:flex; align-items:center; justify-content:space-between; }
  .header h1 { font-size:22px; font-weight:700; background:linear-gradient(135deg,#818cf8,#a78bfa); -webkit-background-clip:text; -webkit-text-fill-color:transparent; }
  .header .company { font-size:14px; color:#94a3b8; }
  .header .company strong { color:#c4b5fd; }
  .main { max-width:1200px; margin:30px auto; padding:0 20px; }
  .tabs { display:flex; gap:4px; margin-bottom:20px; }
  .tab { padding:10px 24px; border-radius:10px 10px 0 0; background:rgba(30,41,59,.5); color:#94a3b8; cursor:pointer; font-size:14px; font-weight:500; border:1px solid transparent; border-bottom:none; transition:all .2s; }
  .tab:hover { color:#c4b5fd; }
  .tab.active { background:rgba(30,41,59,.9); color:#a78bfa; border-color:rgba(99,102,241,.3); }
  .card { background:rgba(30,41,59,.8); backdrop-filter:blur(20px); border:1px solid rgba(99,102,241,.2); border-radius:16px; padding:28px; margin-bottom:20px; }
  .card h2 { font-size:16px; font-weight:600; color:#a78bfa; margin-bottom:20px; padding-bottom:10px; border-bottom:1px solid rgba(99,102,241,.15); }
  .card h2 .icon { margin-right:8px; }
  .form-grid { display:grid; grid-template-columns:1fr 1fr; gap:16px; }
  .form-grid.three { grid-template-columns:1fr 1fr 1fr; }
  .form-group { display:flex; flex-direction:column; }
  .form-group.full { grid-column:1/-1; }
  .form-group label { font-size:12px; color:#94a3b8; margin-bottom:6px; font-weight:500; }
  .form-group input,.form-group textarea,.form-group select { background:rgba(15,23,42,.6); border:1px solid rgba(99,102,241,.2); border-radius:8px; padding:10px 14px; color:#e2e8f0; font-size:14px; font-family:inherit; transition:border-color .2s; }
  .form-group input:focus,.form-group textarea:focus,.form-group select:focus { outline:none; border-color:#818cf8; box-shadow:0 0 0 3px rgba(129,140,248,.1); }
  .form-group input::placeholder { color:#475569; }
  .form-group select option { background:#1e293b; color:#e2e8f0; }
  .btn-row { display:flex; gap:12px; margin-top:20px; flex-wrap:wrap; }
  .btn { padding:12px 28px; border-radius:10px; border:none; font-size:14px; font-weight:600; cursor:pointer; font-family:inherit; transition:all .2s; display:inline-flex; align-items:center; gap:8px; }
  .btn-sm { padding:6px 14px; font-size:12px; border-radius:8px; }
  .btn-primary { background:linear-gradient(135deg,#6366f1,#8b5cf6); color:white; box-shadow:0 4px 15px rgba(99,102,241,.4); }
  .btn-primary:hover { transform:translateY(-1px); box-shadow:0 6px 20px rgba(99,102,241,.5); }
  .btn-primary:disabled { opacity:.5; cursor:not-allowed; transform:none; }
  .btn-secondary { background:rgba(99,102,241,.15); color:#a78bfa; border:1px solid rgba(99,102,241,.3); }
  .btn-secondary:hover { background:rgba(99,102,241,.25); }
  .btn-success { background:linear-gradient(135deg,#059669,#10b981); color:white; }
  .btn-danger { background:rgba(239,68,68,.2); color:#f87171; border:1px solid rgba(239,68,68,.3); }
  .btn-danger:hover { background:rgba(239,68,68,.35); }
  .checkbox-row { display:flex; align-items:center; gap:8px; margin-top:10px; }
  .checkbox-row input[type="checkbox"] { width:18px; height:18px; accent-color:#8b5cf6; }
  .checkbox-row label { font-size:13px; color:#cbd5e1; margin:0; }

  /* 업로드 영역 */
  .upload-zone { border:2px dashed rgba(99,102,241,.3); border-radius:12px; padding:24px; text-align:center; cursor:pointer; transition:all .3s; margin-bottom:16px; position:relative; }
  .upload-zone:hover { border-color:#818cf8; background:rgba(99,102,241,.05); }
  .upload-zone.dragover { border-color:#a78bfa; background:rgba(99,102,241,.1); transform:scale(1.01); }
  .upload-zone input[type="file"] { position:absolute; inset:0; opacity:0; cursor:pointer; }
  .upload-zone .upload-icon { font-size:32px; margin-bottom:8px; }
  .upload-zone .upload-text { font-size:14px; color:#94a3b8; }
  .upload-zone .upload-hint { font-size:12px; color:#64748b; margin-top:4px; }

  /* 파일 목록 */
  .file-list { display:flex; flex-direction:column; gap:8px; }
  .file-item { display:flex; align-items:center; justify-content:space-between; padding:10px 16px; background:rgba(15,23,42,.5); border:1px solid rgba(99,102,241,.15); border-radius:10px; transition:all .2s; }
  .file-item:hover { border-color:rgba(99,102,241,.3); }
  .file-item.selected { border-color:#818cf8; background:rgba(99,102,241,.1); }
  .file-info { display:flex; align-items:center; gap:12px; flex:1; cursor:pointer; }
  .file-info .file-icon { font-size:20px; }
  .file-info .file-name { font-size:13px; font-weight:500; color:#e2e8f0; }
  .file-info .file-meta { font-size:11px; color:#64748b; }
  .file-actions { display:flex; gap:6px; }

  /* 진행 상태 */
  .progress-area { display:none; }
  .progress-area.show { display:block; }
  .progress-bar-wrap { background:rgba(15,23,42,.6); border-radius:10px; height:12px; overflow:hidden; margin:12px 0; }
  .progress-bar { height:100%; border-radius:10px; background:linear-gradient(90deg,#6366f1,#a78bfa,#c4b5fd); background-size:200% 100%; animation:shimmer 2s infinite; transition:width .5s ease; width:0%; }
  @keyframes shimmer { 0%{background-position:200% 0} 100%{background-position:-200% 0} }
  .progress-text { font-size:13px; color:#94a3b8; margin-bottom:4px; }
  .progress-percent { font-size:20px; font-weight:700; color:#a78bfa; }
  .log-box { background:rgba(15,23,42,.8); border:1px solid rgba(99,102,241,.15); border-radius:10px; padding:16px; max-height:250px; overflow-y:auto; font-size:13px; line-height:1.8; font-family:'Consolas',monospace; }
  .log-box .log-line { color:#94a3b8; } .log-box .log-line.success { color:#34d399; } .log-box .log-line.error { color:#f87171; }
  .result-card { display:none; margin-top:20px; background:rgba(16,185,129,.1); border:1px solid rgba(16,185,129,.3); border-radius:12px; padding:20px; }
  .result-card.show { display:block; }
  .result-card.error { background:rgba(248,113,113,.1); border-color:rgba(248,113,113,.3); }
  .result-files { margin:12px 0; }
  .result-file { display:flex; align-items:center; gap:10px; padding:8px 0; font-size:14px; }
  .result-file .path { color:#94a3b8; font-family:monospace; font-size:12px; }
  .tab-content { display:none; } .tab-content.active { display:block; }
  .toast { position:fixed; bottom:30px; right:30px; background:#059669; color:white; padding:12px 24px; border-radius:10px; font-size:14px; font-weight:500; box-shadow:0 8px 30px rgba(0,0,0,.3); transform:translateY(100px); opacity:0; transition:all .3s ease; z-index:1000; }
  .toast.show { transform:translateY(0); opacity:1; }
  .toast.error { background:#dc2626; }

  /* 추출 미리보기 */
  .preview-box { background:rgba(15,23,42,.8); border:1px solid rgba(99,102,241,.15); border-radius:10px; padding:16px; max-height:200px; overflow-y:auto; font-size:12px; line-height:1.6; color:#94a3b8; font-family:'Consolas',monospace; white-space:pre-wrap; margin-top:12px; }
  .parsed-info { display:grid; grid-template-columns:1fr 1fr; gap:8px; margin:12px 0; }
  .parsed-item { padding:8px 12px; background:rgba(99,102,241,.1); border-radius:8px; font-size:13px; }
  .parsed-item .pk { color:#94a3b8; font-size:11px; }
  .parsed-item .pv { color:#c4b5fd; font-weight:500; }
</style>
</head>
<body>

<div class="header">
  <div>
    <h1>입찰 정량평가 자동입력 시스템</h1>
    <div class="company"><strong id="headerCompany"></strong> <span id="headerSep" style="display:none">|</span> 정량평가 양식 자동화 대시보드</div>
  </div>
  <div style="font-size:12px;color:#64748b;">v4.0</div>
</div>

<div class="main">
  <div class="tabs">
    <div class="tab active" onclick="switchTab('run')">🚀 자동입력 실행</div>
    <div class="tab" onclick="switchTab('company')">🏢 회사정보 관리</div>
    <div class="tab" onclick="switchTab('forms')">📁 양식 관리</div>
    <div class="tab" onclick="switchTab('history')">📋 실행 이력</div>
  </div>

  <!-- ═══ 탭1: 자동입력 실행 ═══ -->
  <div id="tab-run" class="tab-content active">
    <div class="card">
      <h2><span class="icon">📝</span>입찰 정보 입력</h2>
      <div class="form-grid">
        <div class="form-group full"><label>입찰명 (용역명)</label><input type="text" id="bidName" placeholder="예: 춘천 2026 세계태권도품새선수권대회 행사운영 대행 용역"></div>
        <div class="form-group"><label>발주처</label><input type="text" id="bidOrg" placeholder="예: (재)춘천레저·태권도조직위원회"></div>
        <div class="form-group"><label>제출일</label><input type="text" id="bidDate" placeholder="예: 2026년  4월  15일"></div>
      </div>
    </div>

    <div class="card">
      <h2><span class="icon">📂</span>양식 파일 선택</h2>

      <!-- 양식 업로드 + 선택 영역 -->
      <div style="display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap;align-items:stretch;">
        <div class="upload-zone" id="runFormUploadZone" style="flex:1;min-width:250px;margin-bottom:0;padding:16px;">
          <input type="file" accept=".hwp" onchange="uploadFormFromRun(this)">
          <div class="upload-icon" style="font-size:24px;">📤</div>
          <div class="upload-text" style="font-size:13px;">새 양식 HWP 업로드</div>
          <div class="upload-hint">클릭 또는 드래그</div>
        </div>
        <div style="flex:2;min-width:300px;display:flex;flex-direction:column;gap:10px;">
          <div class="form-group" style="margin:0;">
            <label>등록된 양식에서 선택</label>
            <select id="formSelect" onchange="onFormSelect()">
              <option value="">-- 양식을 선택하세요 --</option>
            </select>
          </div>
          <div class="form-group" style="margin:0;">
            <label>또는 직접 파일 선택 / 경로 입력</label>
            <div style="display:flex;gap:8px;">
              <input type="text" id="formPath" placeholder="예: C:\Users\...\정량적평가-양식.hwp" style="flex:1;">
              <label class="btn btn-secondary btn-sm" style="cursor:pointer;white-space:nowrap;margin:0;display:inline-flex;align-items:center;">
                📂 찾아보기
                <input type="file" accept=".hwp" onchange="browseFormFile(this)" style="display:none;">
              </label>
            </div>
          </div>
        </div>
      </div>

      <div class="form-grid">
        <div class="form-group">
          <label>출력 파일명 (확장자 제외)</label>
          <input type="text" id="outputName" value="정량적평가-자동입력완료">
        </div>
        <div class="form-group">
          <label>&nbsp;</label>
          <div class="checkbox-row"><input type="checkbox" id="doPdf" checked><label for="doPdf">PDF도 함께 생성</label></div>
          <div class="checkbox-row" style="margin-top:6px"><input type="checkbox" id="demoMode"><label for="demoMode">데모 모드 (입력 과정을 천천히 보기)</label></div>
        </div>
      </div>
      <div class="btn-row">
        <button class="btn btn-primary" id="btnRun" onclick="startRun()">▶ 자동입력 실행</button>
        <button class="btn btn-secondary" onclick="saveConfig()">💾 설정 저장</button>
        <button class="btn btn-secondary" onclick="openFolder()">📁 출력 폴더 열기</button>
      </div>
    </div>

    <div class="card progress-area" id="progressArea">
      <h2><span class="icon">⚡</span>실행 상태</h2>
      <div style="display:flex;align-items:center;gap:20px;margin-bottom:8px;">
        <span class="progress-percent" id="progressPercent">0%</span>
        <span class="progress-text" id="progressText">준비 중...</span>
      </div>
      <div class="progress-bar-wrap"><div class="progress-bar" id="progressBar"></div></div>
      <div class="log-box" id="logBox"></div>
    </div>

    <div class="result-card" id="resultCard">
      <h2 style="margin-bottom:12px;" id="resultTitle"></h2>
      <div class="result-files" id="resultFiles"></div>
      <div class="btn-row" style="margin-top:16px">
        <button class="btn btn-success" onclick="openResult('hwp')">📄 HWP 열기</button>
        <button class="btn btn-success" id="btnOpenPdf" onclick="openResult('pdf')">📑 PDF 열기</button>
      </div>
    </div>
  </div>

  <!-- ═══ 탭2: 회사정보 관리 ═══ -->
  <div id="tab-company" class="tab-content">
    <div class="card">
      <h2><span class="icon">📤</span>회사정보 HWP 파일에서 자동 추출</h2>
      <p style="font-size:13px;color:#94a3b8;margin-bottom:16px;">기존에 작성된 회사정보 HWP 파일을 업로드하면 자동으로 정보를 추출합니다.</p>
      <div style="display:flex;gap:12px;align-items:stretch;flex-wrap:wrap;">
        <div class="upload-zone" id="companyUploadZone" style="flex:1;min-width:250px;margin-bottom:0;">
          <input type="file" accept=".hwp" onchange="uploadCompanyFile(this)">
          <div class="upload-icon">📄</div>
          <div class="upload-text">HWP 파일을 여기에 드래그하거나 클릭</div>
          <div class="upload-hint">.hwp 파일만 가능 | 최대 50MB</div>
        </div>
        <div style="display:flex;flex-direction:column;justify-content:center;gap:8px;">
          <label class="btn btn-primary" style="cursor:pointer;margin:0;">
            📂 파일 선택하기
            <input type="file" accept=".hwp" onchange="uploadCompanyFile(this)" style="display:none;">
          </label>
          <div style="font-size:11px;color:#64748b;text-align:center;">또는 좌측 영역에<br>드래그 & 드롭</div>
        </div>
      </div>
      <div id="companyImportResult" style="display:none;">
        <h3 style="font-size:14px;color:#34d399;margin-bottom:8px;">✅ 추출된 회사정보</h3>
        <div class="parsed-info" id="parsedInfo"></div>
        <div class="btn-row">
          <button class="btn btn-primary btn-sm" onclick="applyParsedInfo()">📥 이 정보를 회사정보에 반영</button>
          <button class="btn btn-secondary btn-sm" onclick="togglePreview()">🔍 원문 보기/숨기기</button>
        </div>
        <div class="preview-box" id="textPreview" style="display:none;"></div>
      </div>
      <!-- 업로드된 회사정보 파일 목록 -->
      <div id="companyFileList" style="margin-top:16px;"></div>
    </div>

    <div class="card">
      <h2><span class="icon">🏢</span>기본 회사정보</h2>
      <div class="form-grid three">
        <div class="form-group"><label>업체명</label><input type="text" id="c_업체명"></div>
        <div class="form-group"><label>대표자</label><input type="text" id="c_대표자"></div>
        <div class="form-group"><label>사업자번호</label><input type="text" id="c_사업자번호"></div>
        <div class="form-group full"><label>주소</label><input type="text" id="c_주소"></div>
        <div class="form-group"><label>전화번호</label><input type="text" id="c_전화번호"></div>
        <div class="form-group"><label>FAX</label><input type="text" id="c_FAX"></div>
        <div class="form-group"><label>설립일</label><input type="text" id="c_설립일"></div>
        <div class="form-group"><label>자본금</label><input type="text" id="c_자본금"></div>
        <div class="form-group"><label>전년도매출액</label><input type="text" id="c_전년도매출액"></div>
      </div>
      <div class="btn-row"><button class="btn btn-primary" onclick="saveCompany()">💾 회사정보 저장</button></div>
    </div>

    <div class="card">
      <h2><span class="icon">👥</span>인력 현황</h2>
      <div id="staffList" style="font-size:13px;color:#94a3b8;line-height:2;"></div>
    </div>
    <div class="card">
      <h2><span class="icon">📜</span>면허/등록증</h2>
      <div id="licenseList" style="font-size:13px;color:#94a3b8;line-height:2;"></div>
    </div>
    <div class="card">
      <h2><span class="icon">📊</span>사업수행실적</h2>
      <div id="projectList" style="font-size:13px;color:#94a3b8;line-height:2;"></div>
    </div>
  </div>

  <!-- ═══ 탭3: 양식 관리 ═══ -->
  <div id="tab-forms" class="tab-content">
    <div class="card">
      <h2><span class="icon">📤</span>양식 파일 업로드</h2>
      <div class="upload-zone" id="formUploadZone">
        <input type="file" accept=".hwp" multiple onchange="uploadFormFiles(this)">
        <div class="upload-icon">📋</div>
        <div class="upload-text">정량평가 양식 HWP 파일을 여기에 드래그하거나 클릭하세요</div>
        <div class="upload-hint">.hwp 파일만 가능 | 여러 파일 동시 업로드 가능 | 최대 50MB</div>
      </div>
    </div>
    <div class="card">
      <h2><span class="icon">📁</span>등록된 양식 목록</h2>
      <div class="file-list" id="formFileList">
        <div style="color:#64748b;text-align:center;padding:20px;">등록된 양식이 없습니다. 위에서 업로드하세요.</div>
      </div>
    </div>
  </div>

  <!-- ═══ 탭4: 이력 ═══ -->
  <div id="tab-history" class="tab-content">
    <div class="card">
      <h2><span class="icon">📋</span>최근 실행 이력</h2>
      <div id="historyList" style="font-size:14px;color:#94a3b8;line-height:2;">아직 실행 이력이 없습니다.</div>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
const TAB_NAMES = ['run','company','forms','history'];
let parsedData = {};

function switchTab(name) {
  document.querySelectorAll('.tab').forEach((t,i) => t.classList.toggle('active', TAB_NAMES[i]===name));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
  if (name==='forms') loadForms();
  if (name==='company') loadCompanyFiles();
}

function showToast(msg, isError) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast show' + (isError ? ' error' : '');
  setTimeout(() => t.classList.remove('show'), isError ? 5000 : 2500);
}

// ── 양식 파일 업로드 ──
async function uploadFormFiles(input) {
  for (const file of input.files) {
    const fd = new FormData();
    fd.append('file', file);
    const res = await fetch('/api/forms/upload', {method:'POST', body:fd});
    const data = await res.json();
    if (data.ok) showToast(`✅ ${file.name} 업로드 완료`);
    else showToast(`❌ ${data.error}`, true);
  }
  input.value = '';
  loadForms();
  loadFormSelect();
}

// 자동입력 실행 탭에서 양식 업로드 (업로드 후 바로 선택됨)
async function uploadFormFromRun(input) {
  const file = input.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  showToast('📤 양식 업로드 중...');
  const res = await fetch('/api/forms/upload', {method:'POST', body:fd});
  const data = await res.json();
  input.value = '';
  if (data.ok) {
    document.getElementById('formPath').value = data.path;
    showToast(`✅ "${file.name}" 업로드 및 선택 완료`);
    loadFormSelect();
  } else {
    showToast(`❌ ${data.error}`, true);
  }
}

// 파일 찾아보기 (로컬 파일 경로를 직접 입력 대신 선택)
function browseFormFile(input) {
  const file = input.files[0];
  if (!file) return;
  // 업로드하고 경로 설정
  const fd = new FormData();
  fd.append('file', file);
  fetch('/api/forms/upload', {method:'POST', body:fd})
    .then(r => r.json())
    .then(data => {
      if (data.ok) {
        document.getElementById('formPath').value = data.path;
        showToast(`✅ "${file.name}" 불러오기 완료`);
        loadFormSelect();
      }
    });
  input.value = '';
}

async function loadForms() {
  const res = await fetch('/api/forms');
  const files = await res.json();
  const el = document.getElementById('formFileList');
  if (files.length === 0) {
    el.innerHTML = '<div style="color:#64748b;text-align:center;padding:20px;">등록된 양식이 없습니다.</div>';
    return;
  }
  el.innerHTML = files.map(f => `
    <div class="file-item">
      <div class="file-info">
        <span class="file-icon">📄</span>
        <div>
          <div class="file-name">${f.name}</div>
          <div class="file-meta">${f.size_str} · ${f.modified}</div>
        </div>
      </div>
      <div class="file-actions">
        <button class="btn btn-secondary btn-sm" onclick="selectForm('${f.path.replace(/\\/g,'\\\\')}','${f.name}')">✅ 선택</button>
        <button class="btn btn-danger btn-sm" onclick="deleteForm('${f.path.replace(/\\/g,'\\\\')}')">🗑️ 삭제</button>
      </div>
    </div>
  `).join('');
}

async function deleteForm(path) {
  if (!confirm('이 양식 파일을 삭제하시겠습니까?')) return;
  await fetch('/api/forms/delete', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({path})});
  showToast('🗑️ 삭제 완료');
  loadForms();
  loadFormSelect();
}

function selectForm(path, name) {
  document.getElementById('formPath').value = path;
  switchTab('run');
  showToast(`📄 "${name}" 선택됨`);
}

async function loadFormSelect() {
  const res = await fetch('/api/forms');
  const files = await res.json();
  const sel = document.getElementById('formSelect');
  sel.innerHTML = '<option value="">-- 양식을 선택하세요 --</option>';
  files.forEach(f => {
    sel.innerHTML += `<option value="${f.path}">${f.name} (${f.size_str})</option>`;
  });
}

function onFormSelect() {
  const val = document.getElementById('formSelect').value;
  if (val) document.getElementById('formPath').value = val;
}

// ── 회사정보 파일 업로드 ──
async function uploadCompanyFile(input) {
  const file = input.files[0];
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.hwp')) {
    showToast('❌ HWP 파일만 업로드 가능합니다', true);
    input.value = '';
    return;
  }
  const fd = new FormData();
  fd.append('file', file);
  showToast('📄 파일 분석 중...');
  try {
    const res = await fetch('/api/company/import', {method:'POST', body:fd});
    const data = await res.json();
    input.value = '';
    if (data.ok) {
      parsedData = data.parsed;
      const el = document.getElementById('parsedInfo');
      el.innerHTML = Object.entries(data.parsed).map(([k,v]) =>
        `<div class="parsed-item"><div class="pk">${k}</div><div class="pv">${v}</div></div>`
      ).join('');
      document.getElementById('textPreview').textContent = data.text_preview || '';
      document.getElementById('companyImportResult').style.display = 'block';
      showToast(`✅ ${Object.keys(data.parsed).length}개 항목 추출 완료`);
      loadCompanyFiles();
    } else {
      showToast(`❌ ${data.error}`, true);
      if (data.text_preview) {
        document.getElementById('textPreview').textContent = data.text_preview;
        document.getElementById('textPreview').style.display = 'block';
      }
    }
  } catch(e) {
    showToast('❌ 서버 연결 실패: ' + e.message, true);
    input.value = '';
  }
}

function applyParsedInfo() {
  const fields = ['업체명','대표자','사업자번호','주소','전화번호','FAX','설립일','자본금','전년도매출액'];
  fields.forEach(k => {
    if (parsedData[k]) {
      const el = document.getElementById('c_' + k);
      if (el) el.value = parsedData[k];
    }
  });
  showToast('✅ 추출된 정보가 입력란에 반영되었습니다. 확인 후 저장해주세요.');
}

function togglePreview() {
  const el = document.getElementById('textPreview');
  el.style.display = el.style.display === 'none' ? 'block' : 'none';
}

async function loadCompanyFiles() {
  const res = await fetch('/api/company/files');
  const files = await res.json();
  const el = document.getElementById('companyFileList');
  if (files.length === 0) { el.innerHTML = ''; return; }
  el.innerHTML = '<h3 style="font-size:13px;color:#94a3b8;margin-bottom:8px;">📂 업로드된 회사정보 파일</h3>' +
    files.map(f => `
      <div class="file-item" style="margin-bottom:4px;">
        <div class="file-info">
          <span class="file-icon">📄</span>
          <div><div class="file-name">${f.name}</div><div class="file-meta">${(f.size/1024).toFixed(0)}KB · ${f.modified}</div></div>
        </div>
        <button class="btn btn-danger btn-sm" onclick="deleteCompanyFile('${f.path.replace(/\\/g,'\\\\')}')">🗑️</button>
      </div>
    `).join('');
}

async function deleteCompanyFile(path) {
  if (!confirm('이 파일을 삭제하시겠습니까?')) return;
  await fetch('/api/company/files/delete', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({path})});
  showToast('🗑️ 삭제 완료');
  loadCompanyFiles();
}

// ── 기존 기능 (회사정보, 설정, 실행 등) ──
async function loadData() {
  try {
    const res = await fetch('/api/company');
    const data = await res.json();
    const info = data['회사정보'] || {};
    ['업체명','대표자','사업자번호','주소','전화번호','FAX','설립일','자본금','전년도매출액'].forEach(k => {
      const el = document.getElementById('c_'+k); if(el) el.value = info[k]||'';
    });
    document.getElementById('headerCompany').textContent = info['업체명']||'';
    document.getElementById('headerSep').style.display = info['업체명'] ? 'inline' : 'none';
    const staff = data['인력현황']||{};
    let sh = '';
    for (const [d,ms] of Object.entries(staff)) { if(Array.isArray(ms)) ms.forEach(m => { sh += `<div>• <strong>${m['성명']}</strong> (${m['직위']}) - ${d} ${m['자격증']?'/ '+m['자격증']:''}</div>`; }); }
    document.getElementById('staffList').innerHTML = sh||'등록된 인력 없음';
    document.getElementById('licenseList').innerHTML = (data['면허_허가_등록증']||[]).map(l=>`<div>• ${l['명칭']} (${l['등록번호']}) - ${l['발급기관']}</div>`).join('')||'없음';
    document.getElementById('projectList').innerHTML = (data['사업수행실적']||[]).map(p=>`<div>• ${p['용역명']} / ${p['발주처']} / ${p['계약금액']}</div>`).join('')||'없음';
  } catch(e){}
  try {
    const res = await fetch('/api/config'); const cfg = await res.json();
    const bid=cfg['입찰정보']||{}, file=cfg['파일경로']||{};
    document.getElementById('bidName').value = bid['입찰명']||'';
    document.getElementById('bidOrg').value = bid['발주처']||'';
    document.getElementById('bidDate').value = bid['제출일']||'';
    document.getElementById('formPath').value = file['양식파일']||'';
    document.getElementById('outputName').value = file['출력파일명']||'자동입력완료';
  } catch(e){}
  loadFormSelect();
}

async function saveCompany() {
  const res = await fetch('/api/company'); const data = await res.json();
  const info = data['회사정보']||{};
  ['업체명','대표자','사업자번호','주소','전화번호','FAX','설립일','자본금','전년도매출액'].forEach(k => { info[k]=document.getElementById('c_'+k).value; });
  data['회사정보']=info;
  await fetch('/api/company',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  document.getElementById('headerCompany').textContent=info['업체명'];
  document.getElementById('headerSep').style.display=info['업체명']?'inline':'none';
  showToast('✅ 회사정보가 저장되었습니다');
}

async function saveConfig() {
  const cfg = {"입찰정보":{"입찰명":document.getElementById('bidName').value,"발주처":document.getElementById('bidOrg').value,"제출일":document.getElementById('bidDate').value},"파일경로":{"양식파일":document.getElementById('formPath').value,"출력폴더":"C:\\tmp","출력파일명":document.getElementById('outputName').value},"옵션":{"PDF변환":document.getElementById('doPdf').checked,"데모모드":false,"데모_대기시간_초":1.5}};
  await fetch('/api/config',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(cfg)});
  showToast('✅ 설정이 저장되었습니다');
}

let pollTimer=null;
async function startRun() {
  const fp=document.getElementById('formPath').value.trim();
  if(!fp){alert('양식 파일을 선택하거나 경로를 입력해주세요.');return;}
  const body={form_path:fp,output_name:document.getElementById('outputName').value||'자동입력완료',bid_info:{"입찰명":document.getElementById('bidName').value,"발주처":document.getElementById('bidOrg').value,"제출일":document.getElementById('bidDate').value},demo_mode:document.getElementById('demoMode').checked};
  document.getElementById('btnRun').disabled=true;
  document.getElementById('progressArea').classList.add('show');
  document.getElementById('resultCard').classList.remove('show');
  document.getElementById('logBox').innerHTML='';
  const res=await fetch('/api/run',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  const data=await res.json();
  if(data.error){alert(data.error);document.getElementById('btnRun').disabled=false;return;}
  pollTimer=setInterval(pollStatus,500);
}

async function pollStatus() {
  const res=await fetch('/api/status'); const s=await res.json();
  const pct=s.total>0?Math.round(s.progress/s.total*100):0;
  document.getElementById('progressPercent').textContent=pct+'%';
  document.getElementById('progressBar').style.width=pct+'%';
  document.getElementById('progressText').textContent=s.current_task;
  const lb=document.getElementById('logBox');
  lb.innerHTML=s.log.map(l=>{let c=l.includes('✓')?'success':(l.includes('❌')||l.includes('오류')?'error':'');return`<div class="log-line ${c}">${l}</div>`;}).join('');
  lb.scrollTop=lb.scrollHeight;
  if(!s.running&&s.result){
    clearInterval(pollTimer);document.getElementById('btnRun').disabled=false;
    const rc=document.getElementById('resultCard');rc.classList.add('show');
    if(s.result.success){rc.classList.remove('error');document.getElementById('resultTitle').textContent='✅ 자동입력 완료!';
      let fh='';if(s.result.hwp_exists)fh+=`<div class="result-file">📄 HWP: <span class="path">${s.result.hwp_path}</span></div>`;
      if(s.result.pdf_exists)fh+=`<div class="result-file">📑 PDF: <span class="path">${s.result.pdf_path}</span></div>`;
      document.getElementById('resultFiles').innerHTML=fh;document.getElementById('btnOpenPdf').style.display=s.result.pdf_exists?'':'none';addHistory(s.result);
    } else {rc.classList.add('error');document.getElementById('resultTitle').textContent='❌ 오류 발생';document.getElementById('resultFiles').innerHTML=`<div style="color:#f87171">${s.result.error}</div>`;}
  }
}

async function openResult(type){const res=await fetch('/api/status');const s=await res.json();if(s.result){const p=type==='hwp'?s.result.hwp_path:s.result.pdf_path;if(p)await fetch('/api/open-file',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({path:p})});}}
async function openFolder(){await fetch('/api/open-folder',{method:'POST'});}
const histories=[];
function addHistory(r){const n=new Date();const t=`${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,'0')}-${String(n.getDate()).padStart(2,'0')} ${String(n.getHours()).padStart(2,'0')}:${String(n.getMinutes()).padStart(2,'0')}`;histories.unshift({time:t,bid:document.getElementById('bidName').value,count:r.ok_count});document.getElementById('historyList').innerHTML=histories.map(h=>`<div style="padding:8px 0;border-bottom:1px solid rgba(99,102,241,.1);"><strong style="color:#a78bfa">${h.time}</strong> — ${h.bid||'(입찰명 없음)'} — 입력 ${h.count}개 ✅</div>`).join('');}

// 드래그앤드롭
['formUploadZone','companyUploadZone','runFormUploadZone'].forEach(id => {
  const z = document.getElementById(id);
  if (!z) return;
  z.addEventListener('dragover', e => { e.preventDefault(); z.classList.add('dragover'); });
  z.addEventListener('dragleave', () => z.classList.remove('dragover'));
  z.addEventListener('drop', e => { e.preventDefault(); z.classList.remove('dragover');
    const input = z.querySelector('input[type="file"]');
    input.files = e.dataTransfer.files;
    input.dispatchEvent(new Event('change'));
  });
});

loadData();
</script>
</body>
</html>
'''

def kill_process_on_port(port):
    """지정 포트를 점유 중인 프로세스를 종료 (기존 대시보드 잔존 방지)"""
    import subprocess
    try:
        result = subprocess.run(
            ['netstat', '-ano'], capture_output=True, text=True, encoding='cp949', errors='ignore'
        )
        pids_to_kill = set()
        for line in result.stdout.splitlines():
            if f':{port} ' in line and 'LISTENING' in line:
                parts = line.split()
                if parts:
                    pid = parts[-1]
                    if pid.isdigit() and pid != str(os.getpid()):
                        pids_to_kill.add(pid)
        for pid in pids_to_kill:
            print(f"  → 기존 대시보드 프로세스 종료 (PID {pid})")
            subprocess.run(['taskkill', '/F', '/PID', pid], capture_output=True)
        return len(pids_to_kill) > 0
    except Exception as e:
        print(f"  [경고] 포트 정리 실패: {e}")
        return False


if __name__ == "__main__":
    import webbrowser
    port = 5000
    print(f"\n{'='*50}")
    print(f"  입찰 정량평가 자동입력 대시보드 v4")
    print(f"  http://localhost:{port}")
    print(f"{'='*50}")

    # 기존 대시보드 프로세스가 실행 중이면 종료 (옛 코드 잔존 방지)
    if kill_process_on_port(port):
        print(f"  구버전 대시보드를 종료하고 최신 코드로 재시작합니다.")
        time.sleep(1)

    # 외부 index.html 존재 여부 표시 (디버깅 도움)
    ext_html = os.path.join(SCRIPT_DIR, 'index.html')
    if os.path.exists(ext_html):
        size_kb = os.path.getsize(ext_html) // 1024
        print(f"  [OK] 외부 index.html 사용 ({size_kb} KB)")
    else:
        print(f"  [WARN] 내장 HTML 사용 (index.html 없음)")

    print(f"\n  브라우저에서 자동으로 열립니다...")
    print(f"  종료: Ctrl+C\n")
    webbrowser.open(f"http://localhost:{port}")
    # debug=False 유지 (리로더 이중실행 방지), 하지만 서버 시작 로그는 표시
    app.run(host='0.0.0.0', port=port, debug=False, use_reloader=False)
