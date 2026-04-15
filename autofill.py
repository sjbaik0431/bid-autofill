# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════╗
║   입찰 정량평가 양식 자동입력 시스템 v3          ║
║   범용 재사용 버전                               ║
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


# ═══════════════════════════════════════════════════════════
# 표(테이블) 셀 직접 접근 함수들
# ═══════════════════════════════════════════════════════════

def _find_text(hwp, text):
    """텍스트를 찾아 커서 이동 (찾기 성공 시 True)"""
    hwp.HAction.Run("MoveDocBegin")
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = text
    hwp.HParameterSet.HFindReplace.Direction = 0
    hwp.HParameterSet.HFindReplace.FindType = 0
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    return hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)


def _insert_text(hwp, text):
    """현재 위치에 텍스트 삽입 (기존 선택 영역 대체)"""
    act = hwp.CreateAction("InsertText")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("Text", text)
    act.Execute(pset)


def fill_table_cell(hwp, label_text, value, direction='right', nth=1):
    """
    표에서 label_text를 찾아 direction 방향의 셀에 value를 입력.
    nth: 동일 라벨이 여러 번 나올 때 n번째를 지정 (기본 1=첫번째)
    direction: 'right' (옆 셀), 'below' (아래 셀)
    """
    if not value:
        return False
    try:
        hwp.HAction.Run("MoveDocBegin")
        found_count = 0
        while True:
            hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            hwp.HParameterSet.HFindReplace.FindString = label_text
            hwp.HParameterSet.HFindReplace.Direction = 0
            hwp.HParameterSet.HFindReplace.FindType = 0
            hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
            ok = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            if not ok:
                return False
            found_count += 1
            if found_count >= nth:
                break
        # 인접 셀 이동
        if direction == 'right':
            hwp.HAction.Run("TableRightCell")
        elif direction == 'below':
            hwp.HAction.Run("TableLowerCell")
        # 셀 전체 선택 후 값 입력
        hwp.HAction.Run("TableCellBlock")
        _insert_text(hwp, str(value))
        return True
    except Exception as e:
        print(f"    [표 입력 실패] {label_text} → {e}")
        return False


def fill_table_rows(hwp, header_text, data_rows, columns):
    """
    표 헤더를 찾은 후 아래 행들에 데이터를 순서대로 채움.
    header_text: 헤더 셀 텍스트 (예: "성  명")
    data_rows: [[col1, col2, ...], ...]
    columns: 열 수
    """
    if not data_rows:
        return 0
    try:
        if not _find_text(hwp, header_text):
            return 0
        # 다음 행 첫 셀로 이동
        hwp.HAction.Run("TableLowerCell")
        filled = 0
        for row in data_rows:
            for col_idx in range(columns):
                val = str(row[col_idx]) if col_idx < len(row) and row[col_idx] else ''
                if val:
                    hwp.HAction.Run("TableCellBlock")
                    _insert_text(hwp, val)
                if col_idx < columns - 1:
                    hwp.HAction.Run("TableRightCell")
            # 다음 행으로 이동
            hwp.HAction.Run("TableLowerCell")
            # 첫 열로 되돌리기
            for _ in range(columns - 1):
                hwp.HAction.Run("TableLeftCell")
            filled += 1
            if filled >= 20:  # 안전장치: 최대 20행
                break
        return filled
    except Exception as e:
        print(f"    [표 행 입력 실패] {header_text} → {e}")
        return 0


def fill_extended_data(hwp, info, extended, status_log=None):
    """
    확장 회사정보(인력/면허/실적/연혁/강점)를 양식 표에 직접 입력.
    info: 기본 회사정보 dict
    extended: 확장 정보 dict
    status_log: 진행 로그 리스트 (append용)
    """
    if not extended:
        return 0
    log = status_log if status_log is not None else []
    count = 0

    # ═══ 1. 서식 제12호 일반현황 표: 라벨 → 옆 셀 입력 ═══
    log.append("📋 서식 제12호 일반현황 표 입력...")
    table_fields = [
        ("업 체 명", info.get('업체명', ''), 'right'),
        ("대  표  자", info.get('대표자', ''), 'right'),
        ("사업자번호", info.get('사업자번호', ''), 'right'),
        ("주       소", info.get('주소', ''), 'right'),
        ("전 화 번 호", info.get('전화번호', ''), 'right'),
        ("FAX", info.get('FAX', ''), 'right'),
        ("해당부문", (extended.get('추가정보', {}) or {}).get('해당부문', ''), 'right'),
    ]
    for label, val, direction in table_fields:
        if val and fill_table_cell(hwp, label, val, direction):
            log.append(f"  ✓ 표셀: {label} → {val[:20]}")
            count += 1

    # 플레이스홀더 치환 (표 내 빈칸 형태)
    placeholder_fills = [
        ("명(상근__/비상근__)", (extended.get('추가정보', {}) or {}).get('임직원수', '')),
    ]
    for find, val in placeholder_fills:
        if val:
            try:
                replace_all(hwp, find, val)
                log.append(f"  ✓ 플레이스홀더: {find} → {val}")
                count += 1
            except:
                pass

    # 설립일 플레이스홀더 (다양한 공백 패턴)
    설립일 = info.get('설립일', '')
    if 설립일:
        for pat in ["년    월    일", "년   월   일", "년  월  일"]:
            try:
                # 설립일 표시 (서식 제12호의 설립일 행)
                replace_all(hwp, pat, 설립일)
            except:
                pass

    # 사업기간 플레이스홀더
    사업기간 = (extended.get('추가정보', {}) or {}).get('사업기간', '')
    if 사업기간:
        for pat in ["년    월 ~     년    월  (    년   개월)",
                     "년   월 ~    년   월  (   년   개월)"]:
            try:
                replace_all(hwp, pat, 사업기간)
                log.append(f"  ✓ 사업기간: {사업기간}")
                count += 1
                break
            except:
                pass

    # 자본금/매출액 "백만원" 플레이스홀더 (표에서 "자본금" 다음의 "백만원")
    자본금 = info.get('자본금', '')
    매출액 = info.get('전년도매출액', '')
    if 자본금:
        if fill_table_cell(hwp, "자본금", 자본금, 'right'):
            log.append(f"  ✓ 자본금: {자본금}")
            count += 1
    if 매출액:
        if fill_table_cell(hwp, "전년도매출액", 매출액, 'right'):
            log.append(f"  ✓ 매출액: {매출액}")
            count += 1

    # ═══ 2. 면허/등록증 표 채우기 ═══
    면허 = extended.get('면허', [])
    if 면허:
        log.append(f"📜 면허/등록증 {len(면허)}건 입력...")
        rows = [[l.get('명칭',''), l.get('등록번호',''), l.get('취득일',''), l.get('발급기관','')]
                for l in 면허]
        filled = fill_table_rows(hwp, "면허·허가·등록증명", rows, 4)
        if filled == 0:
            # 대안 헤더
            filled = fill_table_rows(hwp, "면허/허가", rows, 4)
        if filled > 0:
            log.append(f"  ✓ 면허 {filled}건 입력")
            count += filled

    # ═══ 3. 인력 총괄표 채우기 ═══
    인력 = extended.get('인력', [])
    if 인력:
        log.append(f"👥 인력 {len(인력)}명 입력...")
        rows = [[p.get('성명',''), p.get('직위',''), p.get('근무경력',''),
                 p.get('자격증',''), p.get('취득일',''), p.get('담당업무','')]
                for p in 인력]
        filled = fill_table_rows(hwp, "성  명", rows, 6)
        if filled > 0:
            log.append(f"  ✓ 인력 {filled}명 입력")
            count += filled

        # 사업총괄책임자 / ○○ 분야 책임자 치환
        총괄 = next((p for p in 인력 if '대표' in p.get('직위','') or '총괄' in p.get('담당업무','')), None)
        if 총괄:
            fill_table_cell(hwp, "사업총괄책임자", 총괄.get('성명',''), 'below')

        # ○○ 분야 → 실제 분야명
        seen_duties = set()
        duty_map = {}
        for p in 인력:
            duty = p.get('담당업무', '')
            if duty and duty not in seen_duties and '총괄' not in duty:
                seen_duties.add(duty)
                duty_map[duty] = p
        for duty, person in list(duty_map.items())[:4]:
            try:
                replace_all(hwp, "○○ 분야 책임자", f"{duty} 분야 책임자")
                count += 1
            except:
                break

    # ═══ 4. 사업수행실적 표 채우기 ═══
    실적 = extended.get('실적', [])
    if 실적:
        log.append(f"📊 사업수행실적 {len(실적)}건 입력...")
        rows = [[str(p.get('연번','')), p.get('용역명',''), p.get('용역개요',''),
                 p.get('용역기간',''), p.get('계약금액',''), p.get('발주처','')]
                for p in 실적]
        filled = fill_table_rows(hwp, "용역명", rows, 6)
        if filled > 0:
            log.append(f"  ✓ 실적 {filled}건 입력")
            count += filled

    # ═══ 5. 주요연혁 치환 ═══
    연혁 = extended.get('연혁', [])
    if 연혁:
        log.append(f"📅 주요연혁 {len(연혁)}건 입력...")
        # □ 주요연혁 밑의 ○ 를 실제 내용으로 치환
        연혁_text = '\n'.join([f"○ {h.get('연도','')} {h.get('내용','')}" for h in 연혁])
        try:
            replace_all(hwp, "□ 주요연혁\n○", f"□ 주요연혁\n{연혁_text}")
            count += 1
        except:
            # 줄바꿈 못 찾으면 한 줄씩 시도
            for h in 연혁:
                try:
                    if replace_all(hwp, "○", f"○ {h.get('연도','')} {h.get('내용','')[:60]}"):
                        count += 1
                        break
                except:
                    pass

    # ═══ 6. 회사 강점 치환 ═══
    강점 = extended.get('강점', [])
    if 강점:
        log.append(f"⭐ 회사 강점 입력...")
        강점_text = '\n'.join([f"○ {s.get('제목','')}: {'; '.join(s.get('내용',[])[:3])}" for s in 강점[:3]])
        try:
            replace_all(hwp, "○ (기술력, 지식능력 등)", 강점_text[:300])
            log.append(f"  ✓ 강점 {len(강점)}개 입력")
            count += 1
        except:
            pass

    log.append(f"📋 확장 데이터 입력 완료: {count}건")
    return count

def _space_variants(base_chars, colon=True):
    """
    한글 라벨의 공백 변형 자동 생성
    예: "업체명" → ["업체명 :", "업 체 명 :", "업  체  명 :", "업    체    명 :", "업     체    명 :"]
    """
    chars = list(base_chars)
    suffix = " :" if colon else ""
    variants = set()
    variants.add(base_chars + suffix)       # 붙여쓰기: "업체명 :"
    variants.add(base_chars + ":" if colon else base_chars)  # "업체명:"
    for sp in [1, 2, 3, 4, 5]:
        spaced = (" " * sp).join(chars)
        variants.add(spaced + suffix)       # "업 체 명 :", "업  체  명 :" 등
    # 비대칭 공백 (자주 나오는 패턴)
    if len(chars) >= 3:
        variants.add(("    ".join(chars[:2]) + "    " + "    ".join(chars[2:])) + suffix)
        variants.add(("     ".join(chars[:2]) + "    " + "    ".join(chars[2:])) + suffix)
    return list(variants)


def scan_form_for_patterns(form_text, info, bid_info):
    """
    양식 텍스트를 스캔하여 회사정보 필드와 매칭되는 라벨을 자동 발견.
    공백 패턴에 관계없이, 핵심 한글 키워드로 매칭.
    """
    import re
    extra_patterns = []
    # 키워드 → (카테고리, 값) 매핑
    keyword_map = {
        '업체명': ('업체명', info.get('업체명','')),
        '회사명': ('업체명', info.get('업체명','')),
        '상호': ('업체명', info.get('업체명','')),
        '법인명': ('업체명', info.get('업체명','')),
        '대표자': ('대표자', info.get('대표자','')),
        '대표이사': ('대표자', info.get('대표자','')),
        '담당자': ('대표자', info.get('대표자','')),
        '주소': ('주소', info.get('주소','')),
        '소재지': ('주소', info.get('주소','')),
        '주사무소': ('주소', info.get('주소','')),
        '사업자': ('사업자번호', info.get('사업자번호','')),
        '등록번호': ('사업자번호', info.get('사업자번호','')),
        '전화': ('전화번호', info.get('전화번호','')),
        '연락처': ('전화번호', info.get('전화번호','')),
        '문의처': ('전화번호', info.get('전화번호','')),
        '설립일': ('설립일', info.get('설립일','')),
        '자본금': ('자본금', info.get('자본금','')),
        '매출액': ('매출액', info.get('전년도매출액','')),
        '법인등록번호': ('법인등록번호', info.get('법인등록번호','')),
        '입찰명': ('입찰명', bid_info.get('입찰명','')),
        '용역명': ('입찰명', bid_info.get('입찰명','')),
        '사업명': ('입찰명', bid_info.get('입찰명','')),
        '발주처': ('발주처', bid_info.get('발주처','')),
        '발주기관': ('발주처', bid_info.get('발주처','')),
        '기관명': ('발주처', bid_info.get('발주처','')),
        '소속': ('소속', info.get('업체명','')),
        '직위': ('직위', '대표이사'),
        '성명': ('성명', info.get('대표자','')),
    }
    # 양식에서 "라벨 :" 패턴 추출 (공백 무관)
    seen = set()
    for line in form_text.split('\n'):
        line = line.strip()
        # "한글(공백포함) :" 형태
        m = re.match(r'^([가-힣A-Za-z()（）/·\s]{2,30})\s*[:：]', line)
        if not m:
            continue
        raw_label = m.group(0)  # 콜론 포함
        # 핵심 키워드 추출 (공백 제거하여 비교)
        compressed = re.sub(r'\s+', '', m.group(1))
        for keyword, (cat, val) in keyword_map.items():
            kw_comp = keyword.replace(' ', '')
            if kw_comp in compressed and val and raw_label not in seen:
                extra_patterns.append((f"동적-{cat}", raw_label, f"{raw_label} {val}"))
                seen.add(raw_label)
                break
    return extra_patterns


def build_patterns(info, bid_info, extended=None):
    """
    회사정보 + 입찰정보로 찾아바꾸기 패턴 목록 생성
    extended: dict with 인력/면허/실적/추가정보/연혁/강점/일반데이터
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
    법인번호 = info.get('법인등록번호', '')

    # 확장 추가정보 우선 사용
    ext_info = (extended or {}).get('추가정보', {}) if extended else {}
    if not 법인번호:
        법인번호 = ext_info.get('법인등록번호', '')
    임직원수 = ext_info.get('임직원수', '')
    상근 = ext_info.get('상근인원', '')
    비상근 = ext_info.get('비상근인원', '')
    사업기간 = ext_info.get('사업기간', '')
    해당부문 = ext_info.get('해당부문', '')

    입찰명 = bid_info.get('입찰명', '')
    발주처 = bid_info.get('발주처', '')
    제출일 = bid_info.get('제출일', '')

    patterns = []

    # ═══ 기본 회사정보 (공백 변형 자동 생성) ═══

    # ── 업체명 ──
    for label in _space_variants("업체명") + _space_variants("회사명") + _space_variants("상호") + ["상호명 :", "상호(법인명칭)"]:
        patterns.append(("업체명", label, f"{label} {업체명}"))
    # 인라인 업체명 특수 패턴
    if 업체명:
        patterns.append(("업체명-인라인", "총괄표(업체명:", f"총괄표(업체명:{업체명})"))
        patterns.append(("업체명-인라인", "현황(업체명:", f"현황(업체명:{업체명})"))

    # ── 대표자 ──
    for label in _space_variants("대표자") + _space_variants("대표이사") + ["성 명 :", "성명 :", "담 당 자 :"]:
        patterns.append(("대표자", label, f"{label} {대표자}"))

    # ── 주소 ──
    for label in _space_variants("주소") + _space_variants("소재지") + [
        "본사주소 :", "주사무소 소재지 :", "주사무소소재지 :",
        "주 사 무 소 소 재 지 :", "주 사 무 소  소 재 지 :",
        "회사 :", "회사:"]:
        patterns.append(("주소", label, f"{label} {주소}"))

    # ── 사업자등록번호 ──
    for label in _space_variants("사업자등록번호") + _space_variants("사업자번호") + [
        "사업자 등록번호 :", "등록번호 :"]:
        patterns.append(("사업자번호", label, f"{label} {사번}"))

    # ── 전화번호 ──
    for label in _space_variants("전화번호") + _space_variants("전화") + [
        "T E L :", "TEL :", "연락처 :", "문의처 :", "담당자연락처 :"]:
        patterns.append(("전화번호", label, f"{label} {전화}"))

    # ── FAX ──
    for label in ["FAX :", "FAX:", "팩스 :", "팩 스 :", "F A X :", "F.A.X :", "FAX번호 :"]:
        patterns.append(("FAX", label, f"{label} {팩스}"))

    # ── 자본금 ──
    for label in _space_variants("자본금"):
        patterns.append(("자본금", label, f"{label} {자본금}"))

    # ── 매출액 ──
    for label in _space_variants("전년도매출액") + _space_variants("매출액") + ["전년도 매출액 :"]:
        patterns.append(("매출액", label, f"{label} {매출액}"))

    # ── 설립일 ──
    for label in _space_variants("설립일") + ["설립년도 :", "설립연월일 :"]:
        patterns.append(("설립일", label, f"{label} {설립일}"))

    # ── 소속/직위/성명 (서명란) ──
    for label in _space_variants("소속"):
        patterns.append(("소속", label, f"{label} {업체명}"))
    for label in _space_variants("직위"):
        patterns.append(("직위", label, f"{label} 대표이사"))
    for label in _space_variants("성명"):
        patterns.append(("성명-서명", label, f"{label} {대표자}"))

    # ── 입찰명/발주처 ──
    if 입찰명:
        for label in _space_variants("입찰명") + _space_variants("용역명") + _space_variants("사업명") + [
            "계약건명 :", "과 업 명 :", "과업명 :"]:
            patterns.append(("입찰명", label, f"{label} {입찰명}"))
    if 발주처:
        for label in _space_variants("발주처") + _space_variants("발주기관") + [
            "발주기관명 :", "기 관 명 :", "계약상대자 :"]:
            patterns.append(("발주처", label, f"{label} {발주처}"))

    # ── 법인등록번호 ──
    if 법인번호:
        for label in _space_variants("법인등록번호") + _space_variants("법인번호") + ["법인 등록번호 :"]:
            patterns.append(("법인등록번호", label, f"{label} {법인번호}"))

    # ═══════════════════════════════════════════════
    # 확장 필드 (인력/면허/실적/추가정보) 기반 자동 입력
    # ═══════════════════════════════════════════════
    # ── 임직원수/상근/비상근 ──
    if 임직원수:
        for label in ["인력현황 :", "임직원수 :", "직원수 :", "인 력 현 황 :"]:
            patterns.append(("임직원수", label, f"{label} {임직원수}"))
    if 상근:
        for label in ["상근 :", "상근인원 :", "상근직원 :"]:
            patterns.append(("상근", label, f"{label} {상근}명"))
    if 비상근:
        for label in ["비상근 :", "비상근인원 :"]:
            patterns.append(("비상근", label, f"{label} {비상근}명"))

    # ── 사업기간 ──
    if 사업기간:
        for label in ["사업기간 :", "사 업 기 간 :", "영업기간 :"]:
            patterns.append(("사업기간", label, f"{label} {사업기간}"))

    # ── 해당부문 ──
    if 해당부문:
        for label in ["해당부문 :", "해 당 부 문 :", "업종 :"]:
            patterns.append(("해당부문", label, f"{label} {해당부문}"))

    if extended:
        인력 = extended.get('인력', []) or []
        면허 = extended.get('면허', []) or []
        실적 = extended.get('실적', []) or []
        일반 = extended.get('일반데이터', {}) or {}

        # ── 범용 "라벨:값" 자동입력 (일반데이터) ──
        for lbl, value in 일반.items():
            if not value or len(str(value)) > 100:
                continue
            compact = lbl.replace(' ', '')
            for lbl_pattern in [f"{lbl} :", f"{lbl}:", f"{compact} :"]:
                patterns.append((f"기타-{lbl}", lbl_pattern, f"{lbl_pattern} {value}"))

        # ── 대표이사 성명 (첫 번째 인력 또는 회사정보의 대표자) ──
        if 인력:
            총괄 = next((p for p in 인력 if '대표' in (p.get('직위','') or '') or p.get('담당업무','') == '총괄업무'), None)
            if 총괄:
                patterns.append(("총괄책임자성명", "총괄 책임자 :", f"총괄 책임자 : {총괄.get('성명','')}"))
                patterns.append(("총괄책임자성명", "총괄책임자 :", f"총괄책임자 : {총괄.get('성명','')}"))

        # ── 면허/등록증 주요 항목 (첫 번째 면허의 등록번호/발급기관) ──
        if 면허:
            주면허 = 면허[0]
            patterns.append(("면허등록번호", "등 록 번 호 :", f"등 록 번 호 : {주면허.get('등록번호','')}"))
            patterns.append(("면허발급기관", "발 급 기 관 :", f"발 급 기 관 : {주면허.get('발급기관','')}"))

        # ── 사업수행실적 요약 ──
        if 실적:
            최근실적 = 실적[-1]  # 가장 최근
            patterns.append(("최근실적명", "최근 용역명 :", f"최근 용역명 : {최근실적.get('용역명','')}"))
            patterns.append(("실적건수", "실적 건수 :", f"실적 건수 : {len(실적)}건"))

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
    # 확장 정보 (신규 구조 지원, 없어도 OK)
    extended = {
        '인력': db.get('인력_전체', []),
        '면허': db.get('면허_허가_등록증', []),
        '실적': db.get('사업수행실적', []),
        '추가정보': db.get('추가정보', {}),
        '연혁': db.get('연혁', []),
        '강점': db.get('강점', []),
        '일반데이터': db.get('일반데이터', {})
    }
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
    patterns = build_patterns(info, bid_info, extended)
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
