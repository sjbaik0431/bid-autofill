# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════╗
║   입찰 정량평가 양식 자동입력 시스템 v3          ║
║   (주)스페이스락 - 범용 재사용 버전              ║
╚══════════════════════════════════════════════════╝

사용법:
  1. company_master.json  → 회사정보 (한 번만 설정)
  2. autofill_config.json → 입찰별 설정 (매번 수정)
  3. python autofill.py   또는 autofill.bat 더블클릭

새로운 양식 적용 시:
  - autofill_config.json의 "양식파일" 경로만 변경
  - 양식에 새로운 라벨 패턴이 있으면 아래 EXTRA_PATTERNS에 추가
"""
import win32com.client as win32
import json, os, shutil, sys, time
from datetime import datetime

# ── 경로 설정 ──
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_DB = os.path.join(SCRIPT_DIR, "company_master.json")
CONFIG_FILE = os.path.join(SCRIPT_DIR, "autofill_config.json")

def load_json(path, label):
    """JSON 파일 로드"""
    if not os.path.exists(path):
        print(f"\n[오류] {label} 파일을 찾을 수 없습니다:")
        print(f"  → {path}")
        print(f"  이 파일을 먼저 만들어주세요.")
        sys.exit(1)
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def replace_all(hwp, find, replace):
    """한컴 찾아바꾸기 (전체)"""
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = find
    hwp.HParameterSet.HFindReplace.ReplaceString = replace
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.Direction = 0
    hwp.HParameterSet.HFindReplace.FindType = 0
    return hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

def build_patterns(info, bid_info):
    """
    회사정보 + 입찰정보로 찾아바꾸기 패턴 목록 생성
    ───────────────────────────────────────────────
    양식마다 라벨 형식이 조금씩 다르므로,
    같은 항목을 여러 패턴으로 등록합니다.
    새 양식에서 매칭 안 되는 라벨이 있으면 여기에 추가하세요.
    """
    업체명 = info.get('업체명', '')
    대표자 = info.get('대표자', '')
    주소   = info.get('주소', '')
    사번   = info.get('사업자번호', '')
    전화   = info.get('전화번호', '')
    팩스   = info.get('FAX', '')
    자본금 = info.get('자본금', '')
    매출액 = info.get('전년도매출액', '')
    설립일 = info.get('설립일', '')

    입찰명 = bid_info.get('입찰명', '')
    발주처 = bid_info.get('발주처', '')
    제출일 = bid_info.get('제출일', '')

    patterns = []

    # ── 업체명 (다양한 띄어쓰기 패턴) ──
    for label in ["업 체 명 :", "업체명 :", "회사명 :", "업체명:",
                   "업    체    명 :", "상 호 :", "상호 :", "상호명 :"]:
        patterns.append(("업체명", label, f"{label} {업체명}"))

    # ── 대표자 ──
    for label in ["대 표 자 :", "대표자 :", "대표자:", "성 명 :",
                   "대 표 이 사 :", "대표이사 :"]:
        patterns.append(("대표자", label, f"{label} {대표자}"))

    # ── 주소 ──
    for label in ["주    소 :", "주  소 :", "주소 :", "주소:",
                   "소 재 지 :", "소재지 :", "본사주소 :"]:
        patterns.append(("주소", label, f"{label} {주소}"))

    # ── 사업자등록번호 ──
    for label in ["사업자등록번호 :", "사업자등록번호:", "사 업 자 등 록 번 호 :",
                   "사업자번호 :", "등록번호 :"]:
        patterns.append(("사업자번호", label, f"{label} {사번}"))

    # ── 전화번호 ──
    for label in ["전화번호 :", "전화번호:", "전 화 번 호 :", "T E L :",
                   "TEL :", "전화 :", "연락처 :"]:
        patterns.append(("전화번호", label, f"{label} {전화}"))

    # ── FAX ──
    for label in ["FAX :", "FAX:", "팩스 :", "팩 스 :", "F A X :"]:
        patterns.append(("FAX", label, f"{label} {팩스}"))

    # ── 자본금 ──
    for label in ["자본금 :", "자본금:", "자 본 금 :"]:
        patterns.append(("자본금", label, f"{label} {자본금}"))

    # ── 매출액 ──
    for label in ["전년도매출액 :", "매출액 :", "매출액:", "전년도 매출액 :"]:
        patterns.append(("매출액", label, f"{label} {매출액}"))

    # ── 설립일 ──
    for label in ["설립일 :", "설립일:", "설 립 일 :", "설립년도 :"]:
        patterns.append(("설립일", label, f"{label} {설립일}"))

    # ── 소속/직위/성명 (서명란) ──
    patterns.append(("소속", "소속 :", f"소속 : {업체명}"))
    patterns.append(("직위", "직위 :", f"직위 : 대표이사"))
    patterns.append(("성명-서명", "성명 :", f"성명 : {대표자}"))

    # ── 입찰명/발주처 (있으면 치환) ──
    if 입찰명:
        patterns.append(("입찰명", "용역명 :", f"용역명 : {입찰명}"))
        patterns.append(("입찰명", "사업명 :", f"사업명 : {입찰명}"))
    if 발주처:
        patterns.append(("발주처", "발주처 :", f"발주처 : {발주처}"))
        patterns.append(("발주처", "발주기관 :", f"발주기관 : {발주처}"))

    # ──────────────────────────────────────────
    # ★ 새 양식에서 매칭 안 되는 라벨이 있으면 여기에 추가 ★
    # 예시: patterns.append(("항목명", "찾을텍스트", "바꿀텍스트"))
    # ──────────────────────────────────────────

    return patterns

def build_date_patterns(bid_info):
    """날짜 패턴 생성 - 제출일 기준"""
    제출일 = bid_info.get('제출일', '')
    if not 제출일:
        today = datetime.now()
        제출일 = f"{today.year}년  {today.month}월  {today.day}일"

    # 현재 연도 추출
    year = str(datetime.now().year)

    date_patterns = [
        (f"{year}년    월    일", 제출일),
        (f"{year}년   월    일", 제출일),
        (f"{year}년   월   일", 제출일),
        (f"{year}년  월  일", 제출일),
        (f"{year}년 월 일", 제출일),
    ]
    return date_patterns

def main():
    # ── 설정 로드 ──
    db = load_json(MASTER_DB, "회사정보 DB")
    config = load_json(CONFIG_FILE, "설정")

    info = db["회사정보"]
    bid_info = config["입찰정보"]
    file_cfg = config["파일경로"]
    options = config.get("옵션", {})

    form_file = file_cfg["양식파일"]
    out_dir = file_cfg.get("출력폴더", SCRIPT_DIR)
    out_name = file_cfg.get("출력파일명", "자동입력완료")
    output_hwp = os.path.join(out_dir, f"{out_name}.hwp")
    output_pdf = os.path.join(out_dir, f"{out_name}.pdf")

    do_pdf = options.get("PDF변환", True)
    demo_mode = options.get("데모모드", False)
    demo_delay = options.get("데모_대기시간_초", 1.5)

    print("=" * 55)
    print("  입찰 정량평가 양식 자동입력 시스템 v3")
    print("=" * 55)
    print(f"  업체명 : {info['업체명']}")
    print(f"  대표자 : {info['대표자']}")
    print(f"  입찰명 : {bid_info.get('입찰명', '-')}")
    print(f"  발주처 : {bid_info.get('발주처', '-')}")
    print(f"  양  식 : {os.path.basename(form_file)}")
    if demo_mode:
        print(f"  ★ 데모 모드 (입력 과정을 화면에서 볼 수 있습니다)")
    print("-" * 55)

    # ── 양식 파일 확인 ──
    if not os.path.exists(form_file):
        print(f"\n[오류] 양식 파일을 찾을 수 없습니다:")
        print(f"  → {form_file}")
        print(f"  autofill_config.json의 '양식파일' 경로를 확인하세요.")
        input("\n아무 키나 누르면 종료...")
        sys.exit(1)

    # ── 기존 출력 파일 정리 ──
    for f in [output_hwp, output_pdf]:
        if os.path.exists(f):
            try:
                os.remove(f)
            except PermissionError:
                print(f"\n[오류] 파일이 이미 열려있습니다. 닫고 다시 실행하세요:")
                print(f"  → {f}")
                input("\n아무 키나 누르면 종료...")
                sys.exit(1)

    # ── 양식 복사 ──
    os.makedirs(out_dir, exist_ok=True)
    shutil.copy2(form_file, output_hwp)
    print(f"\n[1/6] 양식 복사 완료")

    # ── 한글 실행 ──
    try:
        hwp = win32.Dispatch("HWPFrame.HwpObject")
    except Exception as e:
        print(f"\n[오류] 한컴오피스를 실행할 수 없습니다:")
        print(f"  {e}")
        print(f"  한컴오피스가 설치되어 있는지 확인하세요.")
        input("\n아무 키나 누르면 종료...")
        sys.exit(1)

    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    hwp.XHwpWindows.Item(0).Visible = True
    print("[2/6] 한글 실행 완료")

    if demo_mode:
        time.sleep(2)

    # ── 파일 열기 ──
    hwp.Open(output_hwp, "HWP", "forceopen:true")
    print("[3/6] 파일 열기 완료")

    if demo_mode:
        time.sleep(2)

    # ── 패턴 생성 ──
    patterns = build_patterns(info, bid_info)
    date_patterns = build_date_patterns(bid_info)

    # ── 자동입력 실행 ──
    print(f"\n[4/6] 자동입력 시작... (패턴 {len(patterns) + len(date_patterns)}개)")
    ok_count = 0
    skip_count = 0

    for category, find, replace in patterns:
        if demo_mode:
            hwp.HAction.Run("MoveDocBegin")
        try:
            replace_all(hwp, find, replace)
            ok_count += 1
            print(f"  ✓ {category:12s} │ '{find}'")
            if demo_mode:
                time.sleep(demo_delay)
        except Exception as e:
            skip_count += 1

    print(f"\n  --- 날짜 입력 ---")
    for find, replace in date_patterns:
        if demo_mode:
            hwp.HAction.Run("MoveDocBegin")
        try:
            replace_all(hwp, find, replace)
            ok_count += 1
            print(f"  ✓ 날짜          │ '{find}' → '{replace}'")
            if demo_mode:
                time.sleep(demo_delay)
        except:
            skip_count += 1

    print(f"\n[5/6] 입력 결과: 성공 {ok_count}개 / 스킵 {skip_count}개")

    # ── 저장 ──
    print(f"\n[6/6] 저장 중...")
    hwp.SaveAs(output_hwp, "HWP")
    print(f"  ✓ HWP: {output_hwp}")

    if do_pdf:
        try:
            hwp.SaveAs(output_pdf, "PDF")
            print(f"  ✓ PDF: {output_pdf}")
        except Exception as e:
            print(f"  △ PDF 변환 실패: {e}")

    # ── 종료 ──
    if demo_mode:
        time.sleep(2)
    hwp.Clear(1)
    hwp.Quit()

    print(f"\n{'=' * 55}")
    print(f"  ✅ 자동입력 완료!")
    print(f"  HWP: {output_hwp}")
    if os.path.exists(output_pdf):
        print(f"  PDF: {output_pdf}")
    print(f"{'=' * 55}")

    input("\n아무 키나 누르면 종료...")

if __name__ == "__main__":
    main()
