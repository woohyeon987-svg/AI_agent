# -*- coding: utf-8 -*-
"""
신한카드 데이터 기반 '사회적 고립 위험 가능성' 1차 탐색 도구
=====================================================================


필요한 라이브러리
-----------------
- pandas   : 엑셀 읽기/쓰기, 그룹화, 통계 계산
- numpy    : 수치 계산(엔트로피, 기울기 등)
- openpyxl : .xlsx 읽기/쓰기 엔진 및 결과 엑셀 서식 지정
- xlrd     : (.xls 파일이 있을 때만) 예전 형식 엑셀 읽기

설치:  pip install pandas numpy openpyxl xlrd
실행:  python risk_analysis.py

폴더 구조
---------
project/
├─ data/                  <- 분석할 엑셀 파일(.xlsx, .xls)을 여기에 넣는다
│   ├─ data_01.xlsx
│   └─ ...
├─ risk_analysis.py
└─ risk_analysis_result.xlsx   (실행 후 생성)
"""

import os #파일 불러오는 헤더
import re
import sys
import warnings #경고메시지 무시

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# =====================================================================
# [설정 영역]  
# =====================================================================

# ---- 입출력 ----------------------------------------------------------
DATA_FOLDER = "./data"                    # 엑셀 파일이 들어 있는 폴더
OUTPUT_FILE = "risk_analysis_result.xlsx" # 결과 엑셀 파일 이름
EXTENSIONS = (".xlsx",".csv")            # 읽을 확장자
SEARCH_SUBFOLDERS = False                 # True면 하위 폴더도 검색

# ---- 시트 ------------------------------------------------------------
# None            : 모든 시트를 읽고, 필요한 컬럼이 모두 있는 시트만 사용
# "Sheet1"        : 해당 이름의 시트만 읽음
# ["Sheet1","S2"] : 여러 시트 지정
SHEET_NAME = None

# ---- 컬럼 이름 (실제 데이터와 다르면 여기서 수정) --------------------
COL_DATE = "TA_YMD"      # 날짜
COL_TIME = "TIME_GB"     # 시간대
COL_REGION = "MCT_SGG_CD"  # 가맹점 지역(시군구)  ※ 거주지가 아니라 '사용한 가맹점'의 지역
COL_INDUSTRY = "MCT_RY_CD" # 가맹점 업종
COL_SEX = "SEX_CCD"      # 성별
COL_AGE = "AGE_CCD"      # 연령대
COL_AMOUNT = "TS_AT"     # 거래/매출 지표
COL_COUNT = "USE_CNT"    # 이용 횟수
REQUIRED_COLS = [COL_DATE, COL_TIME, COL_REGION, COL_INDUSTRY,
                 COL_SEX, COL_AGE, COL_AMOUNT, COL_COUNT]

# ---- 분석 단위 -------------------------------------------------------
# 이 데이터에는 개인 ID가 없는 '집계 데이터'이므로 개인이 아니라
# "분석 단위(세그먼트)"별로 위험지수를 계산합니다.
# 기본값: 연령대 x 성별 x 가맹점 지역
# - 연령대(AGE_CCD)는 반드시 포함해야 합니다(연령대 내 상대 비교에 사용).
# - 여기에 포함된 컬럼은 '활동 다양성/편중' 계산에서 자동 제외됩니다.
#   (예: 지역을 분석 단위에 넣으면 '지역 다양성'은 계산하지 않음)
# - 연령대x성별만 쓰려면 [COL_AGE, COL_SEX] 로 바꾸세요.
#   (이 경우 연령대 내 비교 대상이 2개뿐이라 자동으로 '전체 기준'으로 대체)
SEGMENT_COLS = [COL_AGE, COL_SEX, COL_REGION]
PEER_COL = COL_AGE                         # 상대 비교를 하는 기준 집단
DIVERSITY_DIMS = [COL_REGION, COL_INDUSTRY, COL_TIME]  # 다양성/편중을 볼 차원

# ---- 활동량/기간 설정 ------------------------------------------------
ACTIVITY_METRIC = COL_COUNT   # 활동량 기준: USE_CNT(이용횟수) 권장. TS_AT도 가능
RECENT_MONTHS = 3             # '최근' 구간 개월 수
BASELINE_MONTHS = 6           # '과거' 구간 개월 수(최근 직전부터). None이면 이전 전체
MIN_BASELINE_MONTHS = 2       # 과거 구간 최소 개월 수(미만이면 '변화 판단불가')
MIN_DAYS_PER_MONTH = 7        # 데이터 일수가 이보다 적은 달은 분석에서 제외
MIN_SEGMENT_TOTAL = 30        # 분석 단위의 전체 이용건수가 이보다 적으면 '표본부족'
MIN_PEERS = 5                 # 연령대 안에 분석 단위가 이보다 적으면 전체 기준으로 대체
FILL_MISSING_MONTH_AS_ZERO = True  # 어떤 달에 기록이 없으면 0으로 간주할지

# ---- 구조(다양성/편중) 계산 구간 -------------------------------------
STRUCTURE_WINDOW = "recent"   # "recent"=최근 구간만, "all"=전체 기간

# ---- 활동 변화 설정 --------------------------------------------------
CHANGE_BASIS = "raw"          # "raw"=원래 활동량 변화, "relative"=같은 달 동료 중앙값 대비(계절성 제거)
PERSIST_DROP_THRESHOLD = 0.1 # 과거 평균보다 이 비율 이상 낮은 달을 '감소한 달'로 센다
BASELINE_FLOOR = 0.5          # 과거 활동량이 낮았던 집단의 변화 점수 축소 하한 (0~1)
#   변화 점수 x (BASELINE_FLOOR + (1-BASELINE_FLOOR) x 과거활동량 백분위)
#   -> 원래 활동이 낮았던 집단의 감소는 덜, 원래 활발했던 집단의 감소는 더 반영
DROP_FULL = "auto"            # 감소율이 이 값이면 변화 위험 100%. "auto"=데이터 분포(P90)로 결정
AUTO_FULL_QUANTILE = 0.90
FALLBACK_DROP_FULL = 0.50     # auto인데 감소 사례가 5개 미만일 때 쓰는 임시값

# ---- 가중치 ----------------------------------------------------------
# "auto"  : 사용자가 정한 원칙(변화 중시) + 데이터에서 본 지표 간 중복도로 '제안'된 값
# "manual": 아래 MANUAL_WEIGHTS 를 그대로 사용(합이 100이 아니어도 자동 환산)
WEIGHT_MODE = "auto"
PRIOR_IMPORTANCE = {            # auto 모드의 출발점(임시): 변화를 상대적으로 더 중시
    "activity": 1.0, "diversity": 1.0, "concentration": 1.0, "change": 1.5,
}
MIN_WEIGHT_FACTOR = 0.25        # 중복 보정 시 가중치가 줄어드는 최소 배율
MIN_SEGMENTS_FOR_AUTO = 10      # 상관 계산에 필요한 최소 분석 단위 수
MANUAL_WEIGHTS = {              # manual 모드 전용(예시값일 뿐 확정값 아님)
    "activity": 25, "diversity": 25, "concentration": 25, "change": 25,
}
CHANGE_SPLIT = {"drop": 0.50, "persistence": 0.25}  # 변화 가중치 내부 배분

# ---- 외부 데이터(동시 감소) -----------------------------------------
# 없으면 None. 있으면 엑셀/CSV 경로. 필요한 컬럼:
#   TA_YMD, EXTERNAL_VALUE(활동/유동 지표), 그리고 AGE_CCD(필수 권장)
#   (SEX_CCD, MCT_SGG_CD가 있으면 그 단위까지 맞춰서 결합)
EXTERNAL_FILE = None

# ---- SK 유동인구 데이터 ----------------------------------------------
# data 폴더에 SK 파일(STD_YM, BLOCK_CD, MAN_/WMAN_FLOW_POP_CNT_xxG 컬럼)이 있으면
# 자동으로 인식해 '동시 감소' 계산에 사용합니다. (신한카드 파일과 같은 폴더에 두면 됩니다)
USE_SK_DATA = True
# 성별 값 대응: 신한 SEX_CCD 값이 아래 목록 중 무엇인지 적어 주세요. (대소문자 무시)
SK_SEX_VALUES = {"MAN": ["M", "1", "남", "남성", "MALE"],
                 "WMAN": ["F", "2", "여", "여성", "FEMALE"]}
# 연령대 대응: 기본은 AGE_CCD 안의 숫자로 자동 매칭 (예: "20" -> SK 20대, 60 이상 -> SK 60대이상)
# 신한 코드가 숫자가 아니면 직접 지정: {"신한코드": SK연령(10,20,30,40,50,60)}
SK_AGE_MAP = None
# BLOCK_CD 앞 n자리를 지역으로 보고 MCT_SGG_CD와 맞추려면 자릿수 지정 (예: 5). None이면 전 지역 합산
SK_REGION_PREFIX_LEN = None

# ---- 위험군 분류 -----------------------------------------------------
CLASSIFY_MODE = "fixed"         # "fixed" | "zscore" | "percentile"
FIXED_CUTS = (35, 50)           # fixed : 관심군 시작점, 위험군 시작점 (예시값)
Z_CUTS = (1.0, 2.0)             # zscore: 평균 + 1.0σ, 평균 + 2.0σ
PCT_CUTS = (0.75, 0.90)         # percentile: 상위 25%, 상위 10%
LABELS = ("정상", "관심군", "위험군")

# ---- 데이터 점검 -----------------------------------------------------
DROP_EXACT_DUPLICATES = True    # 파일 간 완전히 같은 행이 있으면 제거
OUTLIER_IQR_K = 3.0             # Q3 + K*IQR 초과를 이상치 후보로 '표시'(제거하지 않음)

# =====================================================================
# [설정 끝] 이 아래는 일반적으로 수정할 필요가 없음
# =====================================================================

ACT = ACTIVITY_METRIC
SK_FLOW_RE = re.compile(r"^(MAN|WMAN)_FLOW_POP_CNT_(\d+)G(U?)$")
SEG = list(SEGMENT_COLS)
DIM_LABEL = {COL_TIME: "시간대", COL_REGION: "지역", COL_INDUSTRY: "업종"}
BLOCKS = ["activity", "diversity", "concentration", "change"]
BLOCK_KR = {"activity": "활동량", "diversity": "활동 다양성",
            "concentration": "활동 편중", "change": "활동 변화"}


# ---------------------------------------------------------------------
# 1. 파일 읽기
# ---------------------------------------------------------------------
def find_excel_files(folder):
    """data 폴더에서 엑셀 파일을 자동으로 찾는다."""
    if not os.path.isdir(folder):
        return []
    found = []
    if SEARCH_SUBFOLDERS:
        walker = os.walk(folder)
    else:
        walker = [(folder, [], os.listdir(folder))]
    out_abs = os.path.abspath(OUTPUT_FILE)
    for root, _, names in walker:
        for n in names:
            p = os.path.join(root, n)
            if n.startswith("~$") or os.path.abspath(p) == out_abs:
                continue
            if n.lower().endswith(tuple(e.lower() for e in EXTENSIONS)):
                found.append(p)
    return sorted(found)


def load_all_excels(folder):
    """모든 엑셀 파일/시트를 읽어 하나의 DataFrame으로 합친다. 실패한 파일은 건너뛴다."""
    files = find_excel_files(folder)
    log, frames, sk_frames = [], [], []
    code_dtype = {c: str for c in [COL_TIME, COL_REGION, COL_INDUSTRY, COL_SEX, COL_AGE,
                                   "STD_YM", "BLOCK_CD"]}
    for path in files:
        name = os.path.basename(path)
        try:
            if name.lower().endswith(".csv"):
                # CSV 파일은 read_csv로 읽기 (인코딩 자동 시도)
                try:
                    raw = pd.read_csv(path, dtype=code_dtype, encoding="utf-8")
                except UnicodeDecodeError:
                    raw = pd.read_csv(path, dtype=code_dtype, encoding="cp949")
            else:
                # 실제 엑셀 파일만 read_excel 사용
                raw = pd.read_excel(path, sheet_name=SHEET_NAME, dtype=code_dtype)
        except Exception as e:  # 파일 하나가 깨져도 계속 진행 # 파일 하나가 깨져도 계속 진행
            print(f"  [오류] {name}: 읽지 못함 -> {type(e).__name__}: {e}")
            log.append({"파일": name, "시트": "-", "상태": "읽기 실패", "행수": 0,
                        "비고": f"{type(e).__name__}: {e}"})
            continue
        if isinstance(raw, pd.DataFrame):       # 시트 1개를 지정한 경우
            raw = {str(SHEET_NAME): raw}
        for sheet, d in raw.items():
            d = d.copy()
            d.columns = [str(c).strip().upper() for c in d.columns]
            # SK 유동인구 파일(컬럼 구조가 다름)인지 먼저 확인
            flow_cols = [c for c in d.columns if SK_FLOW_RE.search(c)]
            if (USE_SK_DATA and "STD_YM" in d.columns and flow_cols
                    and not all(c in d.columns for c in REQUIRED_COLS)):
                keep = ["STD_YM"] + (["BLOCK_CD"] if "BLOCK_CD" in d.columns else []) + flow_cols
                sk = d[keep].copy()
                if "BLOCK_CD" not in sk.columns:
                    sk["BLOCK_CD"] = np.nan
                sk["SOURCE_FILE"] = name
                sk_frames.append(sk)
                log.append({"파일": name, "시트": sheet, "상태": "사용(SK 유동인구)", "행수": len(d),
                            "비고": f"유동인구 컬럼 {len(flow_cols)}개"})
                print(f"  [읽음-SK] {name} / {sheet}: {len(d):,}행 (유동인구 컬럼 {len(flow_cols)}개)")
                continue
            miss = [c for c in REQUIRED_COLS if c not in d.columns]
            if miss:
                print(f"  [건너뜀] {name} / {sheet}: 필요한 컬럼 없음 {miss}")
                log.append({"파일": name, "시트": sheet, "상태": "건너뜀(컬럼 부족)",
                            "행수": len(d), "비고": f"없는 컬럼: {miss}"})
                continue
            d = d[REQUIRED_COLS].copy()
            d["SOURCE_FILE"] = name
            d["SOURCE_SHEET"] = str(sheet)
            frames.append(d)
            log.append({"파일": name, "시트": sheet, "상태": "사용", "행수": len(d), "비고": ""})
            print(f"  [읽음] {name} / {sheet}: {len(d):,}행")
    sk_df = pd.concat(sk_frames, ignore_index=True) if sk_frames else None
    if not frames:
        return None, log, files, sk_df
    return pd.concat(frames, ignore_index=True), log, files, sk_df



# ---------------------------------------------------------------------
# 2. 형식 정리 + 결측/이상치 점검
# ---------------------------------------------------------------------
def parse_dates(s):
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    t = s.astype("string").str.strip().str.replace(r"\.0$", "", regex=True)
    t = t.str.replace(r"[-/.\s]", "", regex=True)
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    is_ym = t.str.len() == 6                    # YYYYMM 형태(SK STD_YM 등)는 그 달 1일로
    out[is_ym] = pd.to_datetime(t[is_ym], format="%Y%m", errors="coerce")
    is_ymd = t.str.len() == 8                   # YYYYMMDD 형태(신한 TA_YMD)
    out[is_ymd] = pd.to_datetime(t[is_ymd], format="%Y%m%d", errors="coerce")
    left = out.isna() & s.notna()               # 그래도 안 되면 일반 변환 시도
    if left.any():
        out[left] = pd.to_datetime(s[left], errors="coerce")
    return out


def clean_data(df):
    """형식 정리 및 점검. (정리된 df, 점검 리포트 DataFrame) 반환"""
    rep = []

    def add(item, count, note=""):
        rep.append({"점검 항목": item, "건수": count, "비고": note})

    add("전체 행 수(합친 후)", len(df))

    df[COL_DATE] = parse_dates(df[COL_DATE])
    for c in [COL_AMOUNT, COL_COUNT]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in [COL_TIME, COL_REGION, COL_INDUSTRY, COL_SEX, COL_AGE]:
        df[c] = df[c].astype("string").str.strip().str.replace(r"\.0$", "", regex=True)

    # 결측 현황
    for c in REQUIRED_COLS:
        n = int(df[c].isna().sum())
        if n:
            add(f"결측: {c}", n)

    # 제거 대상: 날짜/연령대/이용횟수 없음
    bad = df[COL_DATE].isna() | df[COL_AGE].isna() | df[COL_COUNT].isna() | df[ACT].isna()
    add("제거: 날짜·연령대·활동량 결측 행", int(bad.sum()), "분석에서 제외")
    df = df[~bad].copy()    

    # 그 외 범주 결측은 '미상'으로 (활동량 자체는 살리기 위해)
    for c in [COL_TIME, COL_REGION, COL_INDUSTRY, COL_SEX]:
        n = int(df[c].isna().sum())
        if n:
            add(f"'미상' 처리: {c}", n)
        df[c] = df[c].fillna("미상")
    df[COL_AMOUNT] = df[COL_AMOUNT].fillna(0)

    # 중복
    key_cols = [COL_DATE, COL_TIME, COL_REGION, COL_INDUSTRY, COL_SEX, COL_AGE]
    add("키(날짜·시간대·지역·업종·성별·연령대) 중복 행", int(df.duplicated(key_cols).sum()),
        "값이 달라도 키가 같은 경우 포함. 합산하여 사용")
    if DROP_EXACT_DUPLICATES:
        dup = df.duplicated(REQUIRED_COLS)
        add("제거: 완전히 동일한 행", int(dup.sum()), "파일 간 중복 가능성")
        df = df[~dup].copy()

    # 음수/0/이상치 후보 (제거하지 않고 표시만) ?
    for c in [COL_COUNT, COL_AMOUNT]:
        s = df[c]
        add(f"{c} 음수 값", int((s < 0).sum()), "취소/환불 가능성. 그대로 합산")
        add(f"{c} 0 값", int((s == 0).sum()))
        q1, q3 = s.quantile(0.25), s.quantile(0.75)
        hi = q3 + OUTLIER_IQR_K * (q3 - q1)
        add(f"{c} 이상치 후보(>Q3+{OUTLIER_IQR_K}xIQR)", int((s > hi).sum()),
            f"기준값 {hi:,.1f} (제거하지 않음)")

    add("분석 대상 행 수", len(df))
    add("분석 기간", f"{df[COL_DATE].min():%Y-%m-%d} ~ {df[COL_DATE].max():%Y-%m-%d}")
    return df, pd.DataFrame(rep)


# ---------------------------------------------------------------------
# 3. 월별 활동 행렬
# ---------------------------------------------------------------------
def build_monthly(df):
    """(분석단위 x 월) 일평균 활동량 행렬을 만든다."""
    df = df.copy()
    df["MONTH"] = df[COL_DATE].dt.to_period("M")
    days = df.groupby("MONTH")[COL_DATE].nunique()
    eff_min = min(MIN_DAYS_PER_MONTH, int(days.max()))   # 월 단위 데이터면 자동 완화
    ok_months = days[days >= eff_min].index
    dropped = days[days < eff_min]
    df = df[df["MONTH"].isin(ok_months)]
    months = sorted(ok_months)
    days = days.loc[months]

    monthly = df.groupby(SEG + ["MONTH"])[[COL_COUNT, COL_AMOUNT]].sum().reset_index()

    def pivot(col, divide):
        p = monthly.pivot_table(index=SEG, columns="MONTH", values=col, aggfunc="sum")
        p = p.reindex(columns=months)
        if FILL_MISSING_MONTH_AS_ZERO:
            p = p.fillna(0)
        if divide:   # 달마다 데이터 일수가 다를 수 있으므로 '일평균'으로 환산
            p = p.div(days.values, axis=1)
        return p

    total_use = pivot(COL_COUNT, False).sum(axis=1)
    p_act = pivot(ACT, True)
    p_cnt = pivot(COL_COUNT, True)
    p_amt = pivot(COL_AMOUNT, True)
    keys = p_act.index.to_frame(index=False)
    reset = lambda p: p.reset_index(drop=True)
    info = {"days": days, "dropped": dropped, "eff_min": eff_min}
    return df, months, keys, reset(p_act), reset(p_cnt), reset(p_amt), total_use.values, info


# ---------------------------------------------------------------------
# 4. 공통 도구: 연령대 내 백분위, 변화 지표
# ---------------------------------------------------------------------
def peer_percentile(values, peers, eligible):
    """같은 연령대 안에서 0(최저)~1(최고) 백분위. 비교 대상이 적으면 전체 기준으로 대체."""
    v = values.where(eligible)
    n_all = v.notna().sum()
    glob = (v.rank(method="average") - 1) / (n_all - 1) if n_all > 1 else v * np.nan
    cnt = v.groupby(peers).transform("count")
    rk = v.groupby(peers).rank(method="average")
    pct_g = (rk - 1) / (cnt - 1)
    use_g = (cnt >= MIN_PEERS) & v.notna()
    pct = pct_g.where(use_g, glob)
    scope = pd.Series(np.where(v.isna(), "-", np.where(use_g, "연령대 내", "전체 기준(표본부족)")),
                      index=values.index)
    return pct, scope
#================================================================================================
#최근 활동량 비교 // 과거에는 활동량이 어느 정도였고, 최근에는 얼마나 줄었으며, 그 감소가 몇 개월 동안 이어졌는지
#=================================================================================================
def change_metrics(M):
    """과거 대비 최근 활동 변화 지표. M: 행=분석단위, 열=월(일평균)"""
    n = M.shape[1] # 월 데이터가 몇개인지 나타냄
    idx = M.index # 행을 구분하는 이름
    cols = ["base", "recent", "change_pct", "drop", "persist_months", "persist_ratio"]
    if n < RECENT_MONTHS + MIN_BASELINE_MONTHS:
        return pd.DataFrame(np.nan, index=idx, columns=cols)
    rec_cols = list(M.columns[-RECENT_MONTHS:])
    base_cols = list(M.columns[:-RECENT_MONTHS])
    if BASELINE_MONTHS:
        base_cols = base_cols[-BASELINE_MONTHS:]
    base = M[base_cols].mean(axis=1)
    rec = M[rec_cols].mean(axis=1)
    valid = (base > 0) & base.notna()
    ratio = (rec / base).where(valid)
    out = pd.DataFrame(index=idx)
    out["base"], out["recent"] = base, rec
    out["change_pct"] = ratio - 1
    out["drop"] = (1 - ratio).clip(lower=0) # 최근 활동량이 과거 평균보다 얼마나 떨어졌는지(0~1) / 활동량이 증가하면 0으로 표시 감소율만 봄

    # 지속성: 가장 최근 달부터 거꾸로, 연속으로 '과거 평균보다 많이 낮은 달'이 몇 개월인지
    thr = (base * (1 - PERSIST_DROP_THRESHOLD)).values[:, None]
    below = M[rec_cols].values <= thr # 최근 분석 기간의 월별 활동량이 기준선 이하인지 확인
    run = np.cumprod(below[:, ::-1], axis=1).sum(axis=1).astype(float)
    run[~valid.values] = np.nan
    out["persist_months"] = run
    out["persist_ratio"] = run / RECENT_MONTHS  
    return out


def auto_full(values, fallback, fixed):
    """감소율/속도의 '위험 100%' 기준값 결정. 'auto'면 데이터 분포(P90) 사용."""
    if fixed != "auto":
        return float(fixed), "사용자 지정"
    pos = values[values > 0].dropna()
    if len(pos) >= 5:
        return max(float(pos.quantile(AUTO_FULL_QUANTILE)), 0.02), f"데이터 분포 P{int(AUTO_FULL_QUANTILE*100)}"
    return fallback, "사례 부족으로 임시값 사용"


def peer_relative(M, peers):
    """같은 달 동료 중앙값으로 나눠 계절성 같은 공통 변동을 제거."""
    med = M.groupby(peers).transform("median")
    cnt = M.iloc[:, 0].groupby(peers).transform("count")
    small = (cnt < MIN_PEERS).values              # 동료가 부족한 행은 전체 중앙값 사용
    arr = med.values.copy()
    arr[small, :] = M.median(axis=0).values[None, :]
    med = pd.DataFrame(arr, index=M.index, columns=M.columns)
    return M / med.where(med > 0)


def build_sk(sk_raw, keys, months):
    """SK 유동인구를 월별/성별/연령대별로 정리하고, 각 분석 단위(keys)에 대응시킨다.
    반환: (SK 요약표, 분석단위별 월별 유동인구 행렬, 안내 문구 목록). 사용 불가면 (None, None, 안내)"""
    notes = []
    d = sk_raw.copy()
    d["MONTH"] = parse_dates(d["STD_YM"]).dt.to_period("M")
    d = d.dropna(subset=["MONTH"])
    flow_cols = [c for c in d.columns if SK_FLOW_RE.match(c)]
    for c in flow_cols:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    if d["BLOCK_CD"].notna().any():             # 파일이 겹쳐 같은 (월, 소지역)이 두 번 나오면 한 번만
        d["BLOCK_CD"] = d["BLOCK_CD"].astype("string").str.strip()
        nd = int(d.duplicated(["MONTH", "BLOCK_CD"]).sum())
        if nd:
            notes.append(f"SK 중복 행(월+소지역) {nd:,}건 제거")
            d = d.drop_duplicates(["MONTH", "BLOCK_CD"])
    if SK_REGION_PREFIX_LEN and d["BLOCK_CD"].notna().any():
        d["REGION_KEY"] = d["BLOCK_CD"].str[:SK_REGION_PREFIX_LEN]
    else:
        d["REGION_KEY"] = "ALL"
    g = d.groupby(["MONTH", "REGION_KEY"])[flow_cols].sum(min_count=1).reset_index()
    long = g.melt(id_vars=["MONTH", "REGION_KEY"], value_vars=flow_cols, var_name="col", value_name="VALUE")
    mm = long["col"].str.extract(SK_FLOW_RE.pattern)
    long["SK_SEX"] = mm[0]
    long["SK_AGE"] = mm[1].astype(int)
    long["SK_OPEN"] = mm[2] == "U"
    piv = long.pivot_table(index=["REGION_KEY", "SK_SEX", "SK_AGE", "SK_OPEN"], columns="MONTH",
                           values="VALUE", aggfunc="sum").reset_index()
    meta = piv[["REGION_KEY", "SK_SEX", "SK_AGE", "SK_OPEN"]]
    mat = piv.drop(columns=["REGION_KEY", "SK_SEX", "SK_AGE", "SK_OPEN"]).reindex(columns=months)
    overlap = int(mat.notna().any(axis=0).sum())
    sk_months = sorted(long["MONTH"].unique())
    notes.append(f"SK 유동인구 기간: {sk_months[0]} ~ {sk_months[-1]} / 카드 분석 월과 겹치는 달 {overlap}개")
    if overlap == 0:
        notes.append("겹치는 달이 없어 SK 데이터는 사용하지 않음")
        return None, None, notes

    # 신한 분석 단위 -> SK 행 매칭
    rel = [PEER_COL] + ([COL_SEX] if COL_SEX in SEG else []) + \
          ([COL_REGION] if (COL_REGION in SEG and SK_REGION_PREFIX_LEN) else [])
    sex_lookup = {v.upper(): k for k, vs in SK_SEX_VALUES.items() for v in vs}
    cache, unmatched = {}, []
    for tup in keys[rel].drop_duplicates().itertuples(index=False):
        r = dict(zip(rel, tup))
        age = str(r[PEER_COL])
        if SK_AGE_MAP:
            n = SK_AGE_MAP.get(age, np.nan)
        else:
            f = re.search(r"\d+", age)
            n = int(f.group()) if f else np.nan
        mask = pd.Series(False, index=meta.index) if pd.isna(n) else \
            ((meta["SK_AGE"] == n) | (meta["SK_OPEN"] & (n >= meta["SK_AGE"])))
        if COL_SEX in rel:
            sx = sex_lookup.get(str(r[COL_SEX]).strip().upper())
            mask &= (meta["SK_SEX"] == sx) if sx else False
        if COL_REGION in rel:
            mask &= meta["REGION_KEY"] == str(r[COL_REGION])[:SK_REGION_PREFIX_LEN]
        if mask.any():
            cache[tup] = mat[mask].sum(axis=0, min_count=1).values
        else:
            unmatched.append(" / ".join(str(x) for x in tup))
    if unmatched:
        notes.append(f"SK와 매칭되지 않은 분석 단위 값({len(unmatched)}개): {', '.join(unmatched[:10])}"
                     " -> SK_AGE_MAP / SK_SEX_VALUES 설정을 확인하세요")
    if not cache:
        notes.append("매칭되는 분석 단위가 없어 SK 데이터는 사용하지 않음")
        return None, None, notes
    nan_row = np.full(len(months), np.nan)
    SKM = pd.DataFrame(np.vstack([cache.get(t, nan_row) for t in keys[rel].itertuples(index=False)]),
                       columns=months)
    cmeta = change_metrics(mat.reset_index(drop=True))
    table = meta.copy().reset_index(drop=True)
    table.columns = ["지역키", "SK 성별(MAN/WMAN)", "SK 연령(10~60)", "60대이상 여부"]
    for m in months:
        table[f"유동인구_{m}"] = mat[m].values
    table["과거 평균 유동인구"] = cmeta["base"].values
    table["최근 평균 유동인구"] = cmeta["recent"].values
    table["유동인구 변화율(%)"] = cmeta["change_pct"].values * 100
    table["감소 지속 개월"] = cmeta["persist_months"].values
    return table, SKM, notes



def load_external(keys, months, ext_log):
    """외부 활동 데이터를 불러와 분석 단위별 감소 정도(0~1)를 계산한다.
    선택적 컬럼이 없으면 제외하고, 필수 컬럼이 없으면 외부 데이터만 사용하지 않는다.
    """
    if not EXTERNAL_FILE:
        ext_log.append("외부 데이터 파일 미설정 -> 사용 안 함")
        return None

    try:
        # 1. 파일 읽기
        if EXTERNAL_FILE.lower().endswith(".csv"):
            ext = pd.read_csv(EXTERNAL_FILE, dtype=str)
        else:
            ext = pd.read_excel(EXTERNAL_FILE, dtype=str)

        # 2. 컬럼명 정리
        ext.columns = [str(c).strip().upper() for c in ext.columns]

        # 3. 필수 컬럼 확인
        required = [COL_DATE, "EXTERNAL_VALUE"]
        missing_required = [c for c in required if c not in ext.columns]

        if missing_required:
            ext_log.append(
                f"외부 데이터 필수 컬럼 누락: {missing_required} "
                "-> 외부 데이터만 사용 안 함"
            )
            return None

        # 4. 실제 파일에 존재하는 결합 키만 사용
        join = [c for c in SEG if c in ext.columns and c in keys.columns]
        missing_keys = [c for c in SEG if c not in ext.columns]

        if not join:
            ext_log.append(
                "외부 데이터와 연결할 공통 결합 키가 없음 -> 사용 안 함"
            )
            return None

        if missing_keys:
            ext_log.append(
                f"외부 데이터 결합 키 누락: {missing_keys} "
                f"-> 나머지 키 {join}만 사용"
            )

        # 5. 날짜 및 활동량 변환
        ext["MONTH"] = parse_dates(ext[COL_DATE]).dt.to_period("M")
        ext["EXTERNAL_VALUE"] = pd.to_numeric(
            ext["EXTERNAL_VALUE"], errors="coerce"
        )

        # 결합 키의 형식 정리
        for c in join:
            ext[c] = (
                ext[c]
                .astype("string")
                .str.strip()
                .str.replace(r"\.0$", "", regex=True)
            )

        # 사용할 수 있는 데이터만 남김
        ext = ext.dropna(
            subset=["MONTH", "EXTERNAL_VALUE"] + join
        )

        if ext.empty:
            ext_log.append(
                "외부 데이터에 유효한 날짜·활동량 데이터가 없음 -> 사용 안 함"
            )
            return None

        # 6. 분석 단위별 월간 활동량 계산
        p = ext.pivot_table(
            index=join,
            columns="MONTH",
            values="EXTERNAL_VALUE",
            aggfunc="mean"
        )

        p = p.reindex(columns=months)

        # 7. 활동 감소율 계산
        cm = change_metrics(p.reset_index(drop=True))

        full, reason = auto_full(
            cm["drop"],
            FALLBACK_DROP_FULL,
            DROP_FULL
        )

        # 감소 위험 점수: 0~1
        res = p.index.to_frame(index=False)
        res["EXT_DROP"] = (
            cm["drop"] / full
        ).clip(0, 1).values

        ext_log.append(
            f"외부 데이터 사용: 결합 키={join}, "
            f"기준={reason}, 데이터 행={len(ext)}"
        )

        # 8. 기존 분석 데이터와 결합
        merged = keys[SEG].merge(
            res,
            on=join,
            how="left"
        )

        return merged["EXT_DROP"].values

    except Exception as e:
        ext_log.append(
            f"외부 데이터 처리 실패 -> 사용 안 함 "
            f"({type(e).__name__}: {e})"
        )
        return None



# ---------------------------------------------------------------------
# 5. 구조 지표: 활동 다양성 / 활동 편중
# ---------------------------------------------------------------------
def structure_metrics(df, keys, struct_months):
    """지역/업종/시간대별 엔트로피(다양성)와 최다 비율(편중)을 계산한다."""
    d = df[df["MONTH"].isin(struct_months)]
    dims = [c for c in DIVERSITY_DIMS if c not in SEG and df[c].nunique() > 1]
    res = keys[SEG].copy()
    for dim in dims:
        K = df[dim].nunique()
        g = d.groupby(SEG + [dim])[ACT].sum().reset_index()
        g = g[g[ACT] > 0]
        g["p"] = g[ACT] / g.groupby(SEG)[ACT].transform("sum")
        g["plogp"] = -g["p"] * np.log(g["p"])
        g["p2"] = g["p"] ** 2
        agg = g.groupby(SEG).agg(H=("plogp", "sum"), TOP1=("p", "max"),
                                 HHI=("p2", "sum"), NUSED=("p", "size")).reset_index()
        top = g.sort_values("p", ascending=False).drop_duplicates(SEG)[SEG + [dim]]
        top = top.rename(columns={dim: "TOPCAT"})
        agg = agg.merge(top, on=SEG, how="left")
        agg["HN"] = agg["H"] / np.log(K)           # 0(한 곳에 몰림) ~ 1(완전히 고르게)
        agg = agg.rename(columns={c: f"{c}_{dim}" for c in ["H", "TOP1", "HHI", "NUSED", "TOPCAT", "HN"]})
        agg[f"K_{dim}"] = K
        res = res.merge(agg, on=SEG, how="left")
    if dims:
        res["div_raw"] = np.mean([1 - res[f"HN_{d}"] for d in dims], axis=0)   # 클수록 다양성 낮음
        res["conc_raw"] = np.mean([res[f"TOP1_{d}"] for d in dims], axis=0)    # 클수록 편중 높음
    else:
        res["div_raw"] = np.nan
        res["conc_raw"] = np.nan
    return res, dims


# ---------------------------------------------------------------------
# 6. 가중치 / 점수 / 분류
# ---------------------------------------------------------------------
def decide_weights(ind):
    """4개 지표의 가중치를 결정하고, 근거 표를 함께 반환한다."""
    avail = [b for b in BLOCKS if ind[b].notna().sum() > 0]
    rows = {}
    corr = None
    if WEIGHT_MODE == "auto" and len(avail) > 1 and ind[avail].dropna().shape[0] >= MIN_SEGMENTS_FOR_AUTO:
        corr = ind[avail].corr(method="spearman", min_periods=MIN_SEGMENTS_FOR_AUTO).abs()
    raw = {}
    for b in avail:
        mx, factor = np.nan, 1.0
        if WEIGHT_MODE == "auto":
            if corr is not None:
                mx = corr.loc[b].drop(b).max()
                factor = max(MIN_WEIGHT_FACTOR, 1 - (0 if pd.isna(mx) else mx))
            raw[b] = PRIOR_IMPORTANCE[b] * factor
        else:
            raw[b] = float(MANUAL_WEIGHTS.get(b, 0))
        rows[b] = {"지표": BLOCK_KR[b],
                   "출발점(임시 중요도)": PRIOR_IMPORTANCE[b] if WEIGHT_MODE == "auto" else "-",
                   "다른 지표와의 최대 |상관|": mx, "중복 보정 배율": factor if WEIGHT_MODE == "auto" else "-"}
    s = sum(raw.values())
    W = {b: raw[b] / s * 100 for b in avail}
    for b in avail:
        rows[b]["최종 가중치(합100)"] = W[b]
    note = ("자동 제안값(확정 아님): 변화를 상대적으로 중시하되, 서로 많이 겹치는 지표는 가중치를 줄임"
            if WEIGHT_MODE == "auto" else "사용자 수동 입력 가중치")
    return W, pd.DataFrame(rows.values()), corr, note


def level3(p, hi="높음", lo="낮음"):
    if pd.isna(p):
        return "판단불가"
    return lo if p <= 1 / 3 else (hi if p >= 2 / 3 else "보통")


def cutoffs(scores, mode):
    s = scores.dropna()
    if mode == "zscore":
        return (s.mean() + Z_CUTS[0] * s.std(), s.mean() + Z_CUTS[1] * s.std())
    if mode == "percentile":
        return (s.quantile(PCT_CUTS[0]), s.quantile(PCT_CUTS[1]))
    return FIXED_CUTS


def classify(score, cuts):
    out = pd.Series(np.where(score >= cuts[1], LABELS[2],
                    np.where(score >= cuts[0], LABELS[1], LABELS[0])), index=score.index)
    return out.where(score.notna(), "판단불가(표본부족)")


def age_sort_key(s):
    num = pd.to_numeric(s.astype(str).str.extract(r"(\d+)")[0], errors="coerce")
    return num.fillna(1e9)


# ---------------------------------------------------------------------
# 7. 엑셀 저장
# ---------------------------------------------------------------------
def style_workbook(path, block_sheets):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = load_workbook(path)
    hfill = PatternFill("solid", fgColor="DCE6F1")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                c.font = Font(name="Arial", size=10)
        if ws.title in block_sheets:       # 여러 표가 있는 시트: 제목/헤더 줄만 굵게
            for row in ws.iter_rows():
                first = row[0]
                if isinstance(first.value, str) and first.value.startswith("■"):
                    for c in row:
                        c.font = Font(name="Arial", size=11, bold=True)
        else:
            for c in ws[1]:
                c.font = Font(name="Arial", size=10, bold=True)
                c.fill = hfill
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.freeze_panes = "A2"
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            w = max((len(str(c.value)) if c.value is not None else 0) for c in col[:300])
            ws.column_dimensions[col[0].column_letter].width = min(max(10, w * 1.3 + 2), 60)
    wb.save(path)


def write_blocks(writer, sheet, blocks):
    """한 시트에 제목 + 표를 여러 개 이어서 쓴다."""
    r = 0
    for title, d in blocks:
        pd.DataFrame([[f"■ {title}"]]).to_excel(writer, sheet_name=sheet, startrow=r,
                                                  index=False, header=False)
        d.to_excel(writer, sheet_name=sheet, startrow=r + 1, index=False)
        r += len(d) + 4


# ---------------------------------------------------------------------
# 8. 메인
# ---------------------------------------------------------------------
def main():
    print("=" * 66)
    print(" 사회적 고립 위험 가능성 1차 탐색 (신한카드 데이터)")
    print("=" * 66)
    if PEER_COL not in SEG:
        sys.exit("[설정 오류] SEGMENT_COLS에 연령대 컬럼이 포함되어야 합니다.")

    # ① ~ ③ 파일 읽기
    print(f"\n[1] '{DATA_FOLDER}' 폴더에서 엑셀 파일 검색/읽기")
    df, load_log, files, sk_raw = load_all_excels(DATA_FOLDER)
    print(f"    발견한 파일 {len(files)}개")
    if df is None or df.empty:
        sys.exit("\n[중단] 분석에 쓸 수 있는 데이터를 읽지 못했습니다. "
                 "폴더 경로(DATA_FOLDER), 시트 이름(SHEET_NAME), 컬럼 이름을 확인하세요.")

    # ④ ⑤ 정리/점검
    print("\n[2] 데이터 형식 정리 및 점검")
    df, quality = clean_data(df)
    print(quality.to_string(index=False))
    if df.empty:
        sys.exit("[중단] 정리 후 남은 데이터가 없습니다.")

    # ⑥ ⑦ 월별 활동 행렬 + 활동량
    print("\n[3] 월별 활동량 계산")
    df, months, keys, M, Mcnt, Mamt, total_use, info = build_monthly(df)
    n_m = len(months)
    print(f"    분석 월 수: {n_m}개 ({months[0]} ~ {months[-1]}), 분석 단위: {len(keys):,}개")
    if len(info["dropped"]):
        print(f"    [참고] 데이터 일수가 {info['eff_min']}일 미만이라 제외한 달: "
              f"{', '.join(str(m) for m in info['dropped'].index)}")
    gaps = [str(p) for p in pd.period_range(months[0], months[-1], freq="M") if p not in months]
    if gaps:
        print(f"    [주의] 데이터에 없는 달: {', '.join(gaps)} (변화 계산에 영향 가능)")
    if n_m < RECENT_MONTHS + MIN_BASELINE_MONTHS:
        print(f"    [주의] 월 수가 부족해(필요 {RECENT_MONTHS + MIN_BASELINE_MONTHS}개) "
              "'활동 변화'를 계산하지 못합니다. 나머지 지표만으로 환산합니다.")

    S = keys.copy()
    S["eligible"] = total_use >= MIN_SEGMENT_TOTAL
    S["total_use"] = total_use
    elig = S["eligible"]
    peers = S[PEER_COL]
    rcols = list(M.columns[-min(RECENT_MONTHS, n_m):])
    S["recent_act"] = M[rcols].mean(axis=1)
    S["recent_cnt"] = Mcnt[rcols].mean(axis=1)
    S["recent_amt"] = Mamt[rcols].mean(axis=1)
    S["act_pct"], S["peer_scope"] = peer_percentile(S["recent_act"], peers, elig)
    S["act_risk"] = 1 - S["act_pct"]

    # ⑧ ⑨ 다양성 / 편중
    print("[4] 활동 다양성·편중 계산")
    struct_months = months[-RECENT_MONTHS:] if STRUCTURE_WINDOW == "recent" else months
    st, dims = structure_metrics(df, keys, struct_months)
    S = pd.concat([S, st.drop(columns=SEG)], axis=1)
    print(f"    사용한 차원: {[DIM_LABEL.get(d, d) for d in dims]} "
          f"(분석 단위에 포함되어 제외: {[DIM_LABEL.get(d, d) for d in DIVERSITY_DIMS if d in SEG]})")
    S["div_pct"], _ = peer_percentile(S["div_raw"], peers, elig)
    S["conc_pct"], _ = peer_percentile(S["conc_raw"], peers, elig)
    S["div_risk"], S["conc_risk"] = S["div_pct"], S["conc_pct"]

    # ⑩ 활동 변화
    print("[5] 활동 변화 계산")
    cm_raw = change_metrics(M)
    cm_rel = change_metrics(peer_relative(M, peers))
    cm = cm_rel if CHANGE_BASIS == "relative" else cm_raw
    S["base_act"], S["change_pct"] = cm_raw["base"], cm_raw["change_pct"]
    S["drop"], S["persist_months"]= cm["drop"], cm["persist_months"]
    S["persist_ratio"] = cm["persist_ratio"]
    S["change_pct_rel"] = cm_rel["change_pct"]
    S["base_pct"], _ = peer_percentile(S["base_act"], peers, elig)
    base_factor = BASELINE_FLOOR + (1 - BASELINE_FLOOR) * S["base_pct"]
    drop_full, drop_src = auto_full(S["drop"].where(elig), FALLBACK_DROP_FULL, DROP_FULL)
    S["drop_c"] = (S["drop"] / drop_full).clip(0, 1) * base_factor
    S["persist_c"] = S["persist_ratio"] * base_factor
    S["change_comp"] = (CHANGE_SPLIT["drop"] * S["drop_c"] + CHANGE_SPLIT["persistence"] * S["persist_c"])
    print(f"    감소율 100% 기준: {drop_full:.1%} ({drop_src}) / ")

    ext_log = []
    sk_table, ext_drop = None, None
    S["sk_change_pct"], S["sk_drop"] = np.nan, np.nan
    if sk_raw is not None:
        sk_table, SKM, sk_notes = build_sk(sk_raw, keys, months)
        ext_log += sk_notes
        if SKM is not None:
            cm_sk = change_metrics(SKM)
            S["sk_change_pct"], S["sk_drop"] = cm_sk["change_pct"], cm_sk["drop"]
            uniq_drop = cm_sk["drop"][~SKM.round(6).duplicated()]
            pos = uniq_drop[uniq_drop > 0].dropna()
            if len(pos) >= 5:
                sk_full = max(float(pos.quantile(AUTO_FULL_QUANTILE)), 0.02)
            else:
                sk_full = drop_full            # 사례가 적으면 카드와 같은 기준을 적용
            ext_drop = (cm_sk["drop"] / sk_full).clip(0, 1).values
            ext_log.append(f"SK 유동인구 사용: 유동인구 감소율 100% 기준 {sk_full:.1%}")
    elif EXTERNAL_FILE is None:
        ext_log.append("SK 유동인구 데이터 없음 -> '동시 감소'는 계산하지 않음")
    if ext_drop is None:
        ext_drop = load_external(keys, months, ext_log)
    for m in ext_log:
        print("    " + m)
    S["ext_c"] = np.nan
    if ext_drop is not None:       # 동시 감소: 카드 감소와 외부 감소 중 '작은 쪽'(둘 다 줄어야 인정)
        S["ext_c"] = np.minimum((S["drop"] / drop_full).clip(0, 1), ext_drop)

    # ⑪ 위험지수
    print("[6] 위험지수 계산")
    ind = S.loc[elig, ["act_risk", "div_risk", "conc_risk", "change_comp"]]
    ind.columns = BLOCKS
    W, wtable, corr, wnote = decide_weights(ind)
    use_ext = S["ext_c"].notna().any()
    scale = (1 - EXTERNAL_WEIGHT_SHARE) if use_ext else 1.0
    wt = {"활동량": W.get("activity", 0) * scale,
          "활동 다양성": W.get("diversity", 0) * scale,
          "활동 편중": W.get("concentration", 0) * scale,
          "활동 변화(감소율)": W.get("change", 0) * scale * CHANGE_SPLIT["drop"],
          "변화 지속성": W.get("change", 0) * scale * CHANGE_SPLIT["persistence"],
          "동시 감소": EXTERNAL_WEIGHT_SHARE * 100 if use_ext else 0}
    risks = {"활동량": S["act_risk"], "활동 다양성": S["div_risk"], "활동 편중": S["conc_risk"],
             "활동 변화(감소율)": S["drop_c"], "변화 지속성": S["persist_c"],"동시 감소": S["ext_c"]}
    pts = pd.DataFrame({k: risks[k] * wt[k] for k in wt})
    pts = pts.where(elig, np.nan)
    avail_w = sum(wt[k] * pts[k].notna() for k in wt)
    S["score"] = (pts.sum(axis=1, skipna=True) / avail_w.replace(0, np.nan) * 100).where(elig)
    for k in wt:
        S[f"pts_{k}"] = pts[k]
    S["main_driver"] = pts.fillna(-1).idxmax(axis=1).where(S["score"].notna(), "-")
    S["change_ok"] = S["drop"].notna()
    print("    가중치(%): " + ", ".join(f"{k} {v:.1f}" for k, v in wt.items() if v > 0))
    print(f"    ({wnote})")

    # ⑫ 분류
    cuts = cutoffs(S["score"], CLASSIFY_MODE)
    S["class"] = classify(S["score"], cuts)
    sc = S["score"].dropna()
    print(f"\n[7] 위험지수 분포: 평균 {sc.mean():.1f} / 중앙값 {sc.median():.1f} / 표준편차 {sc.std():.1f} "
          f"/ 최소 {sc.min():.1f} / 최대 {sc.max():.1f}")
    print(f"    분류 기준({CLASSIFY_MODE}): 관심군 {cuts[0]:.1f}점 이상, 위험군 {cuts[1]:.1f}점 이상")
    if not (S["class"] == LABELS[2]).any():
        print("    [참고] 현재 기준에서는 위험군이 한 곳도 없습니다. "
              "'설정_분포' 시트의 '기준 후보 비교'를 보고 CLASSIFY_MODE/기준값을 조정해 보세요.")

    # 해설 문장
    def reason(r):
        if pd.isna(r["score"]):
            return "표본 부족(전체 이용건수가 MIN_SEGMENT_TOTAL 미만)으로 판단하지 않음"
        parts = [f"활동량: {level3(r['act_pct'], '높음', '낮음')}"]
        if dims:
            parts.append(f"활동 다양성: {level3(1 - r['div_pct'], '높음', '낮음')}")
            top_dim = max(dims, key=lambda d: (r.get(f'TOP1_{d}') if pd.notna(r.get(f'TOP1_{d}')) else -1))
            parts.append(f"활동 편중: {level3(r['conc_pct'], '높음', '낮음')} "
                         f"(최다 {DIM_LABEL.get(top_dim, top_dim)} 비율 {r[f'TOP1_{top_dim}']:.0%})")
        if pd.notna(r["change_pct"]):
            way = "감소" if r["change_pct"] < 0 else "증가"
            parts.append(f"최근 {RECENT_MONTHS}개월 활동량 {abs(r['change_pct']):.0%} {way}(과거 평균 대비)")
            if r["persist_months"] and r["persist_months"] > 0:
                parts.append(f"감소 지속: {int(r['persist_months'])}개월")
            else:
                parts.append("감소 지속: 없음")
        else:
            parts.append("활동 변화: 판단불가(기간 부족/과거 활동 없음) -> 나머지 지표로 환산")
        if pd.notna(r["sk_change_pct"]):
            way = "감소" if r["sk_change_pct"] < 0 else "증가"
            both = " -> 카드·유동인구 동시 감소" if (r["sk_change_pct"] < 0 and pd.notna(r["change_pct"])
                                                  and r["change_pct"] < 0) else ""
            parts.append(f"유동인구(SK) {abs(r['sk_change_pct']):.0%} {way}{both}")
        parts.append(f"주요 기여 항목: {r['main_driver']}")
        return " | ".join(parts)

    S["reason"] = S.apply(reason, axis=1)
    S["age_order"] = age_sort_key(S[PEER_COL])

    # ⑬ 연령대별 / 성별 요약
    def summarize(by):
        t = S[S["score"].notna()].copy()
        t["is_risk"] = (t["class"] == LABELS[2]).astype(float)
        t["is_watch"] = (t["class"] == LABELS[1]).astype(float)
        t["w"] = t["recent_act"].clip(lower=0)
        g = t.groupby(by)
        out = g.agg(분석단위수=("score", "size"), 평균_위험지수=("score", "mean"),
                    중앙값=("score", "median"), 표준편차=("score", "std"),
                    위험군_비율=("is_risk", "mean"), 관심군_비율=("is_watch", "mean"),
                    활동량_점수=("pts_활동량", "mean"), 다양성_점수=("pts_활동 다양성", "mean"),
                    편중_점수=("pts_활동 편중", "mean"), 변화_점수=("pts_활동 변화(감소율)", "mean"),
                    지속성_점수=("pts_변화 지속성", "mean"), 
                    평균_활동변화율=("change_pct", "mean"),
                    평균_유동인구변화율=("sk_change_pct", "mean"),
                    동시감소_점수=("pts_동시 감소", "mean"))
        out["활동량가중_평균_위험지수"] = g.apply(
            lambda x: np.average(x["score"], weights=x["w"]) if x["w"].sum() > 0 else np.nan)
        out["관심군_이상_비율"] = out["위험군_비율"] + out["관심군_비율"]
        out["위험군_비율"] *= 100
        out["관심군_비율"] *= 100
        out["관심군_이상_비율"] *= 100
        out["평균_활동변화율"] *= 100
        out["평균_유동인구변화율"] *= 100
        if out["평균_유동인구변화율"].isna().all():      # SK 데이터가 없으면 열 제거
            out = out.drop(columns=["평균_유동인구변화율", "동시감소_점수"])
        out = out.reset_index()
        return out.sort_values(by, key=lambda s: age_sort_key(s) if s.name == PEER_COL else s)

    age_tbl = summarize([PEER_COL])
    sex_tbl = summarize([PEER_COL, COL_SEX]) if COL_SEX in S.columns else None
    print("\n[8] 연령대별 결과")
    show = age_tbl[[PEER_COL, "분석단위수", "평균_위험지수", "위험군_비율", "관심군_비율", "평균_활동변화율"]]
    print(show.round(1).to_string(index=False))

    # ⑭ 엑셀 저장
    print(f"\n[9] 결과 저장: {OUTPUT_FILE}")
    S_sorted = S.sort_values(["age_order", "score"], ascending=[True, False])
    seg_cols = SEG
    mlabels = [str(m) for m in months]

    t_all = S_sorted[seg_cols + ["score", "class", "pts_활동량", "pts_활동 다양성", "pts_활동 편중",
                                 "pts_활동 변화(감소율)", "pts_변화 지속성", "pts_동시 감소",
                                 "sk_change_pct", "main_driver", "reason"]].copy()
    t_all["sk_change_pct"] *= 100
    t_all.columns = seg_cols + ["위험지수", "위험군", "활동량 점수", "다양성 점수", "편중 점수", "변화 점수",
                                "지속성 점수","동시감소 점수", "유동인구 변화율(%)", "주요 기여 항목", "해석"]
    if t_all["유동인구 변화율(%)"].isna().all():
        t_all = t_all.drop(columns=["유동인구 변화율(%)"])
    t_all = t_all.sort_values(["위험지수"], ascending=False, na_position="last")

    t_act = S_sorted[seg_cols + ["total_use", "recent_cnt", "recent_amt", "recent_act", "act_pct",
                                 "peer_scope", "act_risk"]].copy()
    t_act.columns = seg_cols + ["전체 이용건수", "최근 일평균 이용건수", "최근 일평균 TS_AT",
                                f"최근 일평균 {ACT}(기준)", "연령대 내 백분위(0~1)", "비교 기준", "활동량 위험(0~1)"]

    div_cols, conc_cols = [], []
    for d in dims:
        div_cols += [f"HN_{d}", f"NUSED_{d}", f"K_{d}"]
        conc_cols += [f"TOP1_{d}", f"TOPCAT_{d}", f"HHI_{d}"]
    t_div = S_sorted[seg_cols + div_cols + ["div_raw", "div_risk"]].copy()
    t_div.columns = (seg_cols + sum([[f"{DIM_LABEL.get(d, d)} 다양성(정규화 엔트로피 0~1)",
                                      f"{DIM_LABEL.get(d, d)} 사용 개수", f"{DIM_LABEL.get(d, d)} 전체 개수"]
                                     for d in dims], [])
                     + ["다양성 낮음 정도(평균,0~1)", "다양성 위험(연령대 내 백분위)"])
    t_conc = S_sorted[seg_cols + conc_cols + ["conc_raw", "conc_risk"]].copy()
    t_conc.columns = (seg_cols + sum([[f"{DIM_LABEL.get(d, d)} 최다 비율",
                                       f"{DIM_LABEL.get(d, d)} 최다 항목",
                                       f"{DIM_LABEL.get(d, d)} 집중도(HHI)"] for d in dims], [])
                      + ["평균 최다 비율", "편중 위험(연령대 내 백분위)"])

    t_chg = S_sorted[seg_cols + ["base_act", "recent_act", "change_pct", "change_pct_rel", "persist_months",
                                 "base_pct", "drop_c", "persist_c","change_comp"]].copy()
    t_chg["change_pct"] *= 100
    t_chg["change_pct_rel"] *= 100
    t_chg.columns = seg_cols + ["과거 일평균 활동량", "최근 일평균 활동량", "활동 변화율(%)(+증가/-감소)",
                                "동료 대비 변화율(%)", "감소 지속 개월",
                                "과거 활동량 백분위", "감소 위험(0~1)", "지속성 위험(0~1)", "변화 위험 종합(0~1)"]
    mon = M.loc[S_sorted.index].copy()
    mon.columns = [f"일평균_{m}" for m in mlabels]
    t_chg = pd.concat([t_chg, mon], axis=1)

    ct = pd.crosstab(S.loc[S["score"].notna(), PEER_COL], S.loc[S["score"].notna(), "class"])
    ct = ct.reindex(columns=list(LABELS), fill_value=0)
    ct.insert(0, "합계", ct.sum(axis=1))
    ct = ct.reset_index().sort_values(PEER_COL, key=age_sort_key)
    t_cls = S_sorted[S_sorted["class"].isin([LABELS[1], LABELS[2]])][seg_cols + ["score", "class", "reason"]]
    t_cls.columns = seg_cols + ["위험지수", "위험군", "해석"]
    t_cls = t_cls.sort_values("위험지수", ascending=False)

    # 설정/분포 시트
    desc = S.loc[S["score"].notna(), ["score", "act_risk", "div_risk", "conc_risk", "change_comp",
                                      "change_pct", "drop"]].copy()
    desc.columns = ["위험지수", "활동량 위험", "다양성 위험", "편중 위험", "변화 위험 종합", "활동 변화율", "감소율"]
    dist = desc.describe(percentiles=[.1, .25, .5, .75, .9]).T.reset_index().rename(columns={"index": "지표"})
    by_age = S[S["score"].notna()].groupby(PEER_COL)["score"].describe(percentiles=[.25, .5, .75, .9]).reset_index()
    by_age = by_age.sort_values(PEER_COL, key=age_sort_key)

    cand = []
    for name, mode in [("현재 설정(" + CLASSIFY_MODE + ")", CLASSIFY_MODE), ("고정 예시(40/70)", "fixed"),
                       (f"평균+{Z_CUTS[0]}σ / +{Z_CUTS[1]}σ", "zscore"),
                       (f"분위수 P{int(PCT_CUTS[0]*100)} / P{int(PCT_CUTS[1]*100)}", "percentile")]:
        c = cutoffs(S["score"], mode)
        cl = classify(S["score"].dropna(), c)
        n = max(len(cl), 1)
        cand.append({"기준 후보": name, "관심군 시작점": c[0], "위험군 시작점": c[1],
                     "정상 비율(%)": (cl == LABELS[0]).sum() / n * 100,
                     "관심군 비율(%)": (cl == LABELS[1]).sum() / n * 100,
                     "위험군 비율(%)": (cl == LABELS[2]).sum() / n * 100})
    cand = pd.DataFrame(cand)

    setting = pd.DataFrame({"항목": [
        "분석 단위", "활동량 기준 컬럼", "최근 구간(개월)", "과거 구간(개월)", "연령대 내 비교 최소 단위 수",
        "변화 기준", "감소율 100% 기준","과거활동 보정 하한", "가중치 방식", "분류 방식",
        "분석한 월", "분석 단위 수(전체/판단 가능)", "SK 유동인구/동시 감소"],
        "값": [" x ".join(SEG), ACT, RECENT_MONTHS, BASELINE_MONTHS or "이전 전체", MIN_PEERS,
               CHANGE_BASIS, f"{drop_full:.1%} ({drop_src})",
               BASELINE_FLOOR, wnote, f"{CLASSIFY_MODE}: 관심군 {cuts[0]:.1f}, 위험군 {cuts[1]:.1f}",
               f"{months[0]} ~ {months[-1]} ({n_m}개월)", f"{len(S)} / {int(S['score'].notna().sum())}",
               ("사용 - " + "; ".join(ext_log)) if S["ext_c"].notna().any() else "사용 안 함"]})
    wfinal = pd.DataFrame([{"세부 항목": k, "가중치": v} for k, v in wt.items()])
    caveat = pd.DataFrame({"해석 시 유의사항": [
        "이 결과는 사회적 고립을 진단하는 것이 아니라, 위험 가능성이 높은 집단을 1차로 탐색하는 용도입니다.",
        "데이터에 개인 ID가 없어 개인이 아닌 '분석 단위(집계 집단)'의 위험도입니다.",
        "MCT_SGG_CD는 '가맹점 지역'이므로 거주지가 아니라 소비가 일어난 지역입니다.",
        "분석 단위의 규모(인구)를 알 수 없어, 활동량 '수준'은 규모 영향을 받습니다. 그래서 자기 자신의 과거 대비 '변화'를 함께 봅니다.",
        "가중치와 분류 기준은 확정값이 아니며, 분포를 확인한 뒤 조정해야 합니다."]})

    corr_df = corr.reset_index().rename(columns={"index": "지표"}) if corr is not None else \
        pd.DataFrame({"안내": ["분석 단위 수가 적어 지표 간 상관은 계산하지 않았습니다."]})
    if corr is not None:
        corr_df["지표"] = corr_df["지표"].map(BLOCK_KR)
        corr_df = corr_df.rename(columns=BLOCK_KR)

    for d_ in [t_all] + ([sk_table] if sk_table is not None else []) + [t_act, t_div, t_conc, t_chg, t_cls, age_tbl, desc, dist, by_age, cand, wtable, corr_df]:
        for c in d_.select_dtypes(include="number").columns:
            d_[c] = d_[c].round(3)
    if sex_tbl is not None:
        for c in sex_tbl.select_dtypes(include="number").columns:
            sex_tbl[c] = sex_tbl[c].round(2)
    for d_ in [age_tbl]:
        for c in d_.select_dtypes(include="number").columns:
            d_[c] = d_[c].round(2)

    block_sheets = {"위험군_분류", "설정_분포"}
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        t_all.to_excel(w, sheet_name="전체_위험도", index=False)
        age_tbl.to_excel(w, sheet_name="연령대별_분석", index=False)
        t_act.to_excel(w, sheet_name="활동량", index=False)
        t_div.to_excel(w, sheet_name="활동_다양성", index=False)
        t_conc.to_excel(w, sheet_name="활동_편중", index=False)
        t_chg.to_excel(w, sheet_name="활동_변화", index=False)
        write_blocks(w, "위험군_분류", [
            (f"연령대별 분류 건수 (기준: 관심군 {cuts[0]:.1f}점, 위험군 {cuts[1]:.1f}점 이상)", ct),
            ("관심군/위험군 분석 단위 목록", t_cls)])
        if sex_tbl is not None:
            sex_tbl.to_excel(w, sheet_name="연령대_성별", index=False)
        write_blocks(w, "설정_분포", [
            ("해석 시 유의사항", caveat), ("사용한 설정", setting),
            ("가중치 결정 근거(지표 단위)", wtable), ("최종 세부 가중치(합 100)", wfinal),
            ("지표 간 상관(Spearman 절대값)", corr_df),
            ("지표 분포(분위수)", dist), ("연령대별 위험지수 분포", by_age),
            ("분류 기준 후보 비교", cand)])
        if sk_table is not None:
            sk_table.to_excel(w, sheet_name="유동인구_변화", index=False)
        quality.to_excel(w, sheet_name="데이터_점검", index=False)
        pd.DataFrame(load_log).to_excel(w, sheet_name="파일_로그", index=False)
    style_workbook(OUTPUT_FILE, block_sheets)
    print("    완료. 시트: 전체_위험도, 연령대별_분석, 활동량, 활동_다양성, 활동_편중, 활동_변화, 위험군_분류, "
          "연령대_성별, 설정_분포, 데이터_점검, 파일_로그")


if __name__ == "__main__":
    main()
