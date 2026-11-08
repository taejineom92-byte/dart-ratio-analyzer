"""
dart_ratios.py — DART(전자공시) 재무제표로 재무비율을 자동 계산하는 학습용 도구

전체 흐름 (위에서 아래로 읽으면 됩니다)
  1) 회사 찾기   : 회사명/종목코드 → DART "고유번호(corp_code, 8자리)"로 바꾼다
  2) 데이터 받기 : OpenDART API로 사업보고서의 전체 재무제표를 JSON으로 받는다
  3) 계정 뽑기   : 수백 줄의 재무제표 중 비율 계산에 필요한 10개 계정만 골라낸다
  4) 비율 계산   : 수익성·안정성·활동성·성장성 비율을 계산한다
  5) 출력       : 표로 보여주고, 원하면 CSV 파일로도 저장한다

사용 예
  python dart_ratios.py --sample                 # API 키 없이 가상 데이터로 연습
  python dart_ratios.py 삼성전자 --year 2025      # 실제 데이터 (DART_API_KEY 환경변수 필요)
  python dart_ratios.py 005930 --csv result.csv  # 종목코드로 조회 + CSV 저장
"""

# ─────────────────────────────────────────────────────────────
# 0. 라이브러리 불러오기
# ─────────────────────────────────────────────────────────────
import argparse                      # 명령줄 옵션(--year, --csv 등)을 해석해 주는 표준 라이브러리
import csv                           # CSV 파일 쓰기 (엑셀에서 열 수 있는 표 형식)
import datetime                      # 오늘 날짜를 구해 기본 조회 연도를 정할 때 사용
import io                            # 메모리 속 바이트를 '파일처럼' 다루게 해 줌 (zip을 디스크에 안 쓰고 풀기)
import json                          # JSON 문자열 ↔ 파이썬 dict 변환
import os                            # 환경변수(API 키)를 읽기 위해 사용
import sys                           # 프로그램 종료(sys.exit) 등에 사용
import unicodedata                   # 한글이 화면에서 2칸을 차지하는지 판별 (표 정렬용)
import xml.etree.ElementTree as ET   # XML 파싱 — 회사 고유번호 목록이 XML로 제공됨
import zipfile                       # zip 압축 풀기 — 회사 목록 XML이 zip으로 묶여서 옴
from pathlib import Path             # 파일 경로를 문자열 대신 객체로 다루는 현대적인 방식

import requests                      # 인터넷에 HTTP 요청을 보내는 외부 라이브러리 (pip install requests)


# ─────────────────────────────────────────────────────────────
# 1. 설정값(상수) — 대문자 이름은 "바꾸지 않는 값"이라는 파이썬 관례
# ─────────────────────────────────────────────────────────────
BASE_URL = "https://opendart.fss.or.kr/api"              # OpenDART API의 공통 주소
HERE = Path(__file__).resolve().parent                    # 이 .py 파일이 있는 폴더
CORP_CODE_CACHE = HERE / "corp_codes.xml"                 # 회사 목록을 한 번 받으면 여기에 저장(캐시)해 재사용
SAMPLE_FILE = HERE / "sample_response.json"               # --sample 모드에서 읽을 가상 응답 파일
ANNUAL_REPORT = "11011"                                   # 보고서 코드: 11011=사업보고서(연간). 분기/반기는 다른 코드

# 필요한 계정과 "어떻게 찾을지"를 정리한 표.
#   sj    : 어느 재무제표에서 찾을지 (BS=재무상태표, IS=손익계산서, CIS=포괄손익계산서)
#   ids   : 1순위 — 표준 계정 ID(XBRL 태그). 회사가 달라도 같은 의미면 같은 ID를 씀
#   names : 2순위 — 회사가 표준 ID를 안 쓴 경우를 대비한 계정 '이름' 후보(공백 제거 후 비교)
ACCOUNTS = {
    "자산총계":   {"sj": ("BS",),        "ids": ("ifrs-full_Assets",),             "names": ("자산총계",)},
    "유동자산":   {"sj": ("BS",),        "ids": ("ifrs-full_CurrentAssets",),      "names": ("유동자산",)},
    "재고자산":   {"sj": ("BS",),        "ids": ("ifrs-full_Inventories",),        "names": ("재고자산",)},
    "부채총계":   {"sj": ("BS",),        "ids": ("ifrs-full_Liabilities",),        "names": ("부채총계",)},
    "유동부채":   {"sj": ("BS",),        "ids": ("ifrs-full_CurrentLiabilities",), "names": ("유동부채",)},
    "자본총계":   {"sj": ("BS",),        "ids": ("ifrs-full_Equity",),             "names": ("자본총계",)},
    "매출액":     {"sj": ("IS", "CIS"),  "ids": ("ifrs-full_Revenue",),            "names": ("매출액", "매출", "수익(매출액)", "영업수익")},
    "매출총이익": {"sj": ("IS", "CIS"),  "ids": ("ifrs-full_GrossProfit",),        "names": ("매출총이익", "매출총이익(손실)")},
    "영업이익":   {"sj": ("IS", "CIS"),  "ids": ("dart_OperatingIncomeLoss",),     "names": ("영업이익", "영업이익(손실)")},
    "당기순이익": {"sj": ("IS", "CIS"),  "ids": ("ifrs-full_ProfitLoss",),         "names": ("당기순이익", "당기순이익(손실)")},
}
# 왜 IS와 CIS 둘 다 보나? → 어떤 회사는 손익계산서를 따로 내고(IS),
# 어떤 회사는 '포괄손익계산서' 하나에 매출부터 다 넣는다(CIS). 그래서 IS → CIS 순서로 찾는다.


# ─────────────────────────────────────────────────────────────
# 2. 회사 찾기: 회사명/종목코드 → DART 고유번호(corp_code)
# ─────────────────────────────────────────────────────────────
def get_api_key():
    """환경변수 DART_API_KEY에서 API 키를 읽는다. 코드에 키를 직접 적지 않는 것이 보안 기본기."""
    key = os.environ.get("DART_API_KEY")                  # 환경변수가 없으면 None이 돌아옴
    if not key:                                           # None 또는 빈 문자열이면
        sys.exit("환경변수 DART_API_KEY가 없습니다. README의 'API 키 설정'을 참고하세요.")  # 안내 후 종료
    return key                                            # 키 문자열 반환


def load_corp_codes(api_key):
    """DART에 등록된 전체 회사 목록을 [{corp_code, corp_name, stock_code}, ...] 형태로 돌려준다."""
    if not CORP_CODE_CACHE.exists():                      # 캐시 파일이 없을 때만 다운로드 (수 MB라 매번 받으면 느림)
        print("회사 목록을 처음 내려받는 중입니다... (한 번만 실행됨)")
        resp = requests.get(f"{BASE_URL}/corpCode.xml",   # 회사 고유번호 API 호출
                            params={"crtfc_key": api_key}, # 쿼리 파라미터: ?crtfc_key=키
                            timeout=60)                    # 60초 안에 응답 없으면 포기 (무한 대기 방지)
        resp.raise_for_status()                           # HTTP 상태가 4xx/5xx면 예외를 일으킴
        try:
            zf = zipfile.ZipFile(io.BytesIO(resp.content))  # 받은 바이트를 메모리에서 바로 zip으로 연다
        except zipfile.BadZipFile:                        # 키가 틀리면 zip 대신 오류 메시지(XML)가 옴
            sys.exit(f"회사 목록 다운로드 실패 — API 키를 확인하세요.\n응답: {resp.text[:300]}")
        with zf:                                          # with: 블록이 끝나면 zip을 자동으로 닫아 줌
            xml_bytes = zf.read(zf.namelist()[0])         # zip 안의 첫 번째(유일한) 파일 = CORPCODE.xml
        CORP_CODE_CACHE.write_bytes(xml_bytes)            # 다음 실행을 위해 디스크에 저장

    root = ET.parse(CORP_CODE_CACHE).getroot()            # XML 파일을 읽어 트리 구조로 만들고 최상위 요소를 얻음
    corps = []                                            # 결과를 담을 빈 리스트
    for item in root.iter("list"):                        # <list> 태그 하나 = 회사 하나
        corps.append({
            "corp_code": item.findtext("corp_code"),                 # 8자리 고유번호 (API 호출에 필요)
            "corp_name": (item.findtext("corp_name") or "").strip(), # 회사명 (앞뒤 공백 제거)
            "stock_code": (item.findtext("stock_code") or "").strip(),  # 6자리 종목코드 (비상장사는 빈 값)
        })
    return corps


def find_corp(corps, query):
    """사용자가 입력한 회사명 또는 종목코드로 회사 하나를 찾는다."""
    query = query.strip()                                 # 입력값 앞뒤 공백 제거
    if query.isdigit() and len(query) == 6:               # 숫자 6자리면 종목코드로 판단
        matches = [c for c in corps if c["stock_code"] == query]   # 리스트 컴프리헨션: 조건에 맞는 것만 모음
    else:                                                 # 그 외에는 회사명으로 판단
        matches = [c for c in corps if c["corp_name"] == query]    # 이름이 '정확히' 같은 회사만

    listed = [c for c in matches if c["stock_code"]]      # 같은 이름이 여럿이면 상장사(종목코드 있음)를 우선
    if listed:
        return listed[0]
    if matches:                                           # 상장사가 없으면 비상장사라도 반환
        return matches[0]

    # 정확히 일치하는 회사가 없으면, 이름에 검색어가 '포함된' 상장사를 힌트로 보여 주고 종료
    hints = [c["corp_name"] for c in corps if query in c["corp_name"] and c["stock_code"]][:10]
    msg = f"'{query}' 회사를 찾지 못했습니다."
    if hints:
        msg += " 혹시 이 중 하나인가요? " + ", ".join(hints)
    sys.exit(msg)


# ─────────────────────────────────────────────────────────────
# 3. 재무제표 받기: OpenDART '단일회사 전체 재무제표' API
# ─────────────────────────────────────────────────────────────
def fetch_statements(api_key, corp_code, year, fs_div):
    """재무제표 행(row) 리스트를 돌려준다. 해당 데이터가 없으면 None."""
    params = {
        "crtfc_key": api_key,        # 인증키
        "corp_code": corp_code,      # 8자리 고유번호
        "bsns_year": str(year),      # 사업연도 (예: "2025")
        "reprt_code": ANNUAL_REPORT, # 사업보고서(연간)
        "fs_div": fs_div,            # CFS=연결재무제표, OFS=별도(개별)재무제표
    }
    resp = requests.get(f"{BASE_URL}/fnlttSinglAcntAll.json", params=params, timeout=30)  # API 호출
    resp.raise_for_status()                               # HTTP 오류 확인
    data = resp.json()                                    # 응답 본문(JSON 문자열)을 dict로 변환

    status = data.get("status")                           # DART는 HTTP 200이어도 내부 status로 성공/실패를 알려 줌
    if status == "013":                                   # "013" = 조회된 데이터 없음
        return None
    if status != "000":                                   # "000" = 정상. 그 외는 키 오류, 호출한도 초과 등
        sys.exit(f"DART 오류 [{status}]: {data.get('message')}")
    return data["list"]                                   # 실제 재무제표 행들이 들어 있는 리스트


# ─────────────────────────────────────────────────────────────
# 4. 필요한 계정만 뽑기
# ─────────────────────────────────────────────────────────────
def to_number(text):
    """'1,234,567' 이나 '-500' 같은 문자열을 정수로 바꾼다. 비어 있으면 None."""
    if text is None:                                      # 값 자체가 없으면
        return None
    text = str(text).replace(",", "").strip()             # 쉼표·공백 제거
    if text in ("", "-"):                                 # DART는 값이 없을 때 빈 문자열이나 '-'를 주기도 함
        return None
    return int(text)                                      # 금액은 '원' 단위 정수


def normalize(name):
    """계정 이름 비교용: 공백을 모두 없앤다. ('매출 총이익' == '매출총이익')"""
    return name.replace(" ", "")


def pick_account(rows, spec, column):
    """rows(재무제표 전체)에서 spec(찾는 법)에 맞는 계정 하나를 골라 column 열의 금액을 돌려준다."""
    for sj in spec["sj"]:                                 # 재무제표 종류를 우선순위대로 (예: IS 먼저, 다음 CIS)
        same_sheet = [r for r in rows if r.get("sj_div") == sj]   # 해당 재무제표의 행만 추림
        for r in same_sheet:                              # 1순위: 표준 계정 ID가 일치하는 행
            if r.get("account_id") in spec["ids"]:
                return to_number(r.get(column))
        for r in same_sheet:                              # 2순위: 계정 이름이 일치하는 행
            if normalize(r.get("account_nm", "")) in spec["names"]:
                return to_number(r.get(column))
    return None                                           # 끝까지 못 찾으면 None (예: 금융회사엔 매출총이익이 없음)


def extract_accounts(rows):
    """당기(올해)와 전기(작년) 금액을 각각 dict로 만든다."""
    # thstrm_amount = 당기 금액, frmtrm_amount = 전기 금액 (DART 응답의 열 이름)
    current = {name: pick_account(rows, spec, "thstrm_amount") for name, spec in ACCOUNTS.items()}
    previous = {name: pick_account(rows, spec, "frmtrm_amount") for name, spec in ACCOUNTS.items()}
    return current, previous                              # 튜플로 두 개를 한 번에 반환

def correct_check(cur):
    """자산 = 부채 + 자본 여부를 검증"""
    자산, 부채, 자본 = cur["자산총계"], cur["부채총계"], cur["자본총계"]
    if 자산 is None or 부채 is None or 자본 is None:
        return
    if 자산 != 부채 + 자본:
        print("⚠ 자산총계가 부채총계 + 자본총계와 일치하지 않습니다. 계정 추출을 확인하세요.")

# ─────────────────────────────────────────────────────────────
# 5. 비율 계산
# ─────────────────────────────────────────────────────────────
def safe_div(a, b):
    """a / b. 값이 없거나 분모가 0이면 계산하지 않고 None (0으로 나누기 오류 방지)."""
    if a is None or b is None or b == 0:
        return None
    return a / b


def average(cur, prev):
    """기초(작년 말)와 기말(올해 말)의 평균. 작년 값이 없으면 기말 값을 그대로 사용."""
    if cur is None:
        return None
    if prev is None:
        return cur
    return (cur + prev) / 2


def growth(cur, prev):
    """증가율 = (올해 - 작년) / |작년|. 작년이 적자(음수)여도 부호가 뒤집히지 않도록 절댓값 사용."""
    if cur is None or prev is None or prev == 0:
        return None
    return (cur - prev) / abs(prev)


def pct(x):
    """비율(0.125)을 퍼센트(12.5)로 바꾼다. None이면 그대로 None."""
    return None if x is None else x * 100

def debt_ratio(debt, equity):
    """부채비율, 자본잠식(자본 0 이하)이면 의미가 없으므로 None."""
    if equity is not None and equity <= 0:
        return None
    return pct(safe_div(debt, equity))

def roe(profit, avg_equity):
    """ROE를 구하는 함수"""
    if avg_equity is not None and avg_equity <= 0:
        return None
    return pct(safe_div(profit, avg_equity))

def calc_ratios(cur, prev):
    """비율 목록을 (분류, 이름, 값, 단위, 공식) 튜플들의 리스트로 돌려준다."""
    # ROE·ROA·회전율의 분모는 '1년 동안의 평균 잔액'을 쓰는 것이 정석
    # (손익은 1년 내내 쌓인 값인데, 재무상태표는 연말 하루의 사진이기 때문)
    avg_assets = average(cur["자산총계"], prev["자산총계"])
    avg_equity = average(cur["자본총계"], prev["자본총계"])
    # 당좌자산 = 유동자산 중 재고자산을 뺀 것 (재고는 현금화가 느리다고 보기 때문)
    quick_assets = None if cur["유동자산"] is None else cur["유동자산"] - (cur["재고자산"] or 0)

    return [
        # ── 수익성: 얼마나 잘 버는가
        ("수익성", "매출총이익률", pct(safe_div(cur["매출총이익"], cur["매출액"])), "%", "매출총이익 / 매출액"),
        ("수익성", "영업이익률",   pct(safe_div(cur["영업이익"], cur["매출액"])),   "%", "영업이익 / 매출액"),
        ("수익성", "순이익률",     pct(safe_div(cur["당기순이익"], cur["매출액"])), "%", "당기순이익 / 매출액"),
        ("수익성", "ROE",          roe(cur["당기순이익"], avg_equity),    "%", "당기순이익 / 평균 자본총계"),
        ("수익성", "ROA",          pct(safe_div(cur["당기순이익"], avg_assets)),    "%", "당기순이익 / 평균 자산총계"),
        # ── 안정성: 망하지 않을 체력이 있는가
        ("안정성", "부채비율",     debt_ratio(cur["부채총계"], cur["자본총계"]), "%", "부채총계 / 자본총계"),
        ("안정성", "유동비율",     pct(safe_div(cur["유동자산"], cur["유동부채"])), "%", "유동자산 / 유동부채"),
        ("안정성", "당좌비율",     pct(safe_div(quick_assets, cur["유동부채"])),    "%", "(유동자산 - 재고자산) / 유동부채"),
        ("안정성", "자기자본비율", pct(safe_div(cur["자본총계"], cur["자산총계"])), "%", "자본총계 / 자산총계"),
        # ── 활동성: 자산을 얼마나 바쁘게 굴리는가
        ("활동성", "총자산회전율", safe_div(cur["매출액"], avg_assets),             "회", "매출액 / 평균 자산총계"),
        # ── 성장성: 작년보다 얼마나 컸는가
        ("성장성", "매출액증가율",   pct(growth(cur["매출액"], prev["매출액"])),         "%", "(당기 - 전기) / |전기|"),
        ("성장성", "영업이익증가율", pct(growth(cur["영업이익"], prev["영업이익"])),     "%", "(당기 - 전기) / |전기|"),
        ("성장성", "순이익증가율",   pct(growth(cur["당기순이익"], prev["당기순이익"])), "%", "(당기 - 전기) / |전기|"),
    ]


# ─────────────────────────────────────────────────────────────
# 6. 출력
# ─────────────────────────────────────────────────────────────
def display_width(text):
    """화면에서 차지하는 칸 수. 한글·한자 같은 '넓은 글자'는 2칸, 영문·숫자는 1칸."""
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in text)


def pad(text, width):
    """text 뒤에 공백을 붙여 화면상 width칸으로 맞춘다. (f-string 정렬은 한글을 1칸으로 세서 줄이 어긋남)"""
    return text + " " * max(width - display_width(text), 0)


def fmt_won(value):
    """원 단위 금액을 '억원' 단위 문자열로. 1억 = 100,000,000원."""
    return "N/A" if value is None else f"{value / 1e8:,.0f}억"   # :,.0f → 천 단위 쉼표, 소수점 없음


def fmt_ratio(value, unit):
    """비율 값을 '12.34%' 또는 '1.26회'처럼 표시."""
    return "N/A" if value is None else f"{value:,.2f}{unit}"     # :,.2f → 소수점 둘째 자리까지


def print_report(corp_name, year, fs_div, cur, prev, ratios):
    fs_label = "연결" if fs_div == "CFS" else "별도"              # 조건 표현식(삼항 연산자)
    print()
    print(f"=== {corp_name} | {year}년 사업보고서 | {fs_label}재무제표 ===")
    if cur["자본총계"] is not None and cur["자본총계"] < 0:
        print("자본잠식 상태입니다. 부채비율 등 일부 비율은 해석에 주의하세요")        
    correct_check(cur)
    print()
    print("[주요 계정]")
    print(pad("계정", 14) + pad("당기", 16) + "전기")              # 표 머리글
    for name in ACCOUNTS:                                         # dict를 for로 돌리면 키(계정 이름)가 나옴
        print(pad(name, 14) + pad(fmt_won(cur[name]), 16) + fmt_won(prev[name]))
    print()
    print("[재무비율]")
    last_group = None                                             # 분류가 바뀔 때만 분류명을 찍기 위한 변수
    for group, name, value, unit, formula in ratios:              # 튜플 언패킹: 5개 값을 각 변수에 나눠 담음
        if group != last_group:                                   # 새 분류가 시작되면
            print(f"- {group}")
            last_group = group
        print("  " + pad(name, 16) + pad(fmt_ratio(value, unit), 12) + formula)
    print()
    print("※ N/A = 해당 계정이 공시에 없거나, 분모가 0이거나, 자본잠식이라 계산 불가")


def save_csv(path, corp_name, year, ratios):
    # encoding="utf-8-sig": 맨 앞에 BOM을 넣어 엑셀이 한글을 깨뜨리지 않고 열게 함
    # newline="": csv 모듈이 줄바꿈을 직접 처리하므로 빈 줄이 끼지 않게 함 (윈도우에서 특히 중요)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)                                    # CSV 작성기 생성
        writer.writerow(["회사", "연도", "분류", "비율", "값", "단위", "공식"])  # 머리글 행
        for group, name, value, unit, formula in ratios:
            shown = "" if value is None else round(value, 4)       # None은 빈 칸으로
            writer.writerow([corp_name, year, group, name, shown, unit, formula])
    print(f"CSV 저장 완료: {path}")


# ─────────────────────────────────────────────────────────────
# 7. 프로그램 시작점
# ─────────────────────────────────────────────────────────────
def build_parser():
    parser = argparse.ArgumentParser(description="DART 공시 재무제표로 재무비율을 계산합니다.")
    parser.add_argument("company", nargs="?",                     # nargs="?": 생략 가능한 위치 인자
                        help="회사명(예: 삼성전자) 또는 6자리 종목코드(예: 005930)")
    parser.add_argument("--year", type=int,
                        default=datetime.date.today().year - 1,   # 기본값: 작년 (사업보고서는 이듬해 3월쯤 공시됨)
                        help="사업연도 (기본: 작년)")
    parser.add_argument("--fs", choices=["CFS", "OFS"], default="CFS",
                        help="CFS=연결(기본), OFS=별도")
    parser.add_argument("--csv", metavar="파일경로", help="결과를 CSV로도 저장")
    parser.add_argument("--sample", action="store_true",          # store_true: 옵션을 쓰면 True, 안 쓰면 False
                        help="API 키 없이 가상 예시 데이터로 실행")
    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()                                    # 실제 명령줄 입력을 해석

    if args.sample:                                               # ── 연습 모드: 파일에서 읽기
        data = json.loads(SAMPLE_FILE.read_text(encoding="utf-8"))
        rows = data["list"]
        corp_name, year, fs_used = "예시전자(가상 데이터)", rows[0]["bsns_year"], "CFS"
    else:                                                         # ── 실전 모드: DART API 호출
        if not args.company:                                      # 회사명을 안 넣었으면 사용법과 함께 오류
            parser.error("회사명 또는 종목코드를 입력하세요. (연습은 --sample)")
        api_key = get_api_key()
        corp = find_corp(load_corp_codes(api_key), args.company)
        corp_name, year, fs_used = corp["corp_name"], args.year, args.fs
        print(f"{corp_name} (고유번호 {corp['corp_code']}) {year}년 재무제표 조회 중...")
        rows = fetch_statements(api_key, corp["corp_code"], year, fs_used)
        if rows is None and fs_used == "CFS":                     # 자회사가 없는 회사는 연결재무제표가 없음
            print("연결재무제표가 없어 별도재무제표로 다시 조회합니다.")
            fs_used = "OFS"
            rows = fetch_statements(api_key, corp["corp_code"], year, fs_used)
        if rows is None:
            sys.exit(f"{year}년 사업보고서 재무제표가 없습니다. 연도를 확인하세요 (아직 공시 전일 수 있음).")

    cur, prev = extract_accounts(rows)                            # 4단계: 계정 뽑기
    ratios = calc_ratios(cur, prev)                               # 5단계: 비율 계산
    print_report(corp_name, year, fs_used, cur, prev, ratios)     # 6단계: 화면 출력
    if args.csv:
        save_csv(args.csv, corp_name, year, ratios)


# 이 파일을 'python dart_ratios.py'로 직접 실행했을 때만 main()을 부른다.
# 다른 파일에서 'import dart_ratios'로 가져다 쓸 때는 자동 실행되지 않는다.
if __name__ == "__main__":
    main()
