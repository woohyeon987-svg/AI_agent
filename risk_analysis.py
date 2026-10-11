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
├─ data/                  <- 분석할 엑셀 파일(.xlsx, .xls)을 여기에 넣는다.
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
COL_REGION = "MCT_SGG_CD"  # 가맹점 지역(시군구)  ※ 거주지가 아니라 '사용한 가맹점'의 지역
COL_INDUSTRY = "MCT_RY_CD" # 가맹점 업종
COL_SEX = "SEX_CCD"      # 성별
COL_AGE = "AGE_CCD"      # 연령대
COL_AMOUNT = "TS_AT"     # 거래/매출 지표
COL_COUNT = "USE_CNT"    # 이용 횟수
# 시간대(TIME_GB)는 사용하지 않습니다. 기준일자는 TA_YMD 하나로 통일합니다.
REQUIRED_COLS = [COL_DATE, COL_REGION, COL_INDUSTRY,
                 COL_SEX, COL_AGE, COL_AMOUNT, COL_COUNT]

SK_COL_YM = "STD_YM"         # 기준년월
SK_COL_BLOCK = "BLOCK_CD"    # 블록 코드
SK_FLOW_RE = re.compile(r"FLOW_POP_CNT", re.IGNORECASE)

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
DIVERSITY_DIMS = [COL_REGION, COL_INDUSTRY]  # 다양성을 볼 차원(분석 단위에 포함된 컬럼은 자동 제외 -> 업종만 사용)

# ---- 활동량/기간 설정 ------------------------------------------------
ACTIVITY_METRIC = COL_COUNT   # 활동량 기준: USE_CNT(이용횟수) 권장. TS_AT도 가능
RECENT_MONTHS = 3             # '최근' 구간 개월 수 (활동량·다양성·변화·지속성 모두 동일하게 사용)
BASELINE_MONTHS = 3           # '과거' 구간 개월 수(최근 직전부터). None이면 이전 전체
MIN_BASELINE_MONTHS = 2       # 과거 구간 최소 개월 수(미만이면 '변화 판단불가')
MIN_DAYS_PER_MONTH = 7        # 데이터 일수가 이보다 적은 달은 분석에서 제외
MIN_SEGMENT_TOTAL = 30        # 분석 단위의 전체 이용건수가 이보다 적으면 '표본부족'
MIN_PEERS = 5                 # 연령대 안에 분석 단위가 이보다 적으면 전체 기준으로 대체
FILL_MISSING_MONTH_AS_ZERO = True  # 어떤 달에 기록이 없으면 0으로 간주할지

# ---- 구조(다양성/편중) 계산 구간 -------------------------------------
STRUCTURE_WINDOW = "recent"   # "recent"=최근 구간만, "all"=전체 기간

# ---- 활동 양식 변화 / 지속성 설정 ------------------------------------
# 활동 변화 : 최근 구간과 과거 구간의 '업종 분포' 차이 = 1 - 코사인 유사도 (차이가 클수록 높은 점수)
# 지속성    : 월별 업종 분포가 '과거 분포와 달랐던 달'이 최근 몇 개월 연속 이어졌는지 (길수록 높은 점수)
# '달랐던 달' 기준: "auto"면 과거 구간의 각 달이 과거 평균 분포에서 평소 벗어나는 정도(분위수)보다 더 벗어난 달.
PATTERN_DIFF_THRESHOLD = "auto"   # "auto" 또는 직접 지정한 거리값(0~1)
PATTERN_DIFF_QUANTILE = 0.75      # auto일 때 사용할 분위수
# 참고용(점수 미반영): 활동량(이용건수) 자체가 과거 평균보다 이 비율 이상 낮으면 '감소한 달'로 센다
PERSIST_DROP_THRESHOLD = 0.1

# ---- 점수 계산 방식 --------------------------------------------------
# 가중치는 사용하지 않습니다. 4개 항목(활동량 / 활동 다양성 / 활동 감소율 / 감소 지속)을
# 각각 0~100점(높을수록 활동적·안정)으로 만들고 단순 평균한 값이 '안정지수'입니다.
# 안정지수가 낮을수록 사회적 고립 위험 가능성이 큰 집단입니다.

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

# ---- 위험군 분류 (안정지수가 '낮을수록' 위험) -----------------------
CLASSIFY_MODE = "percentile"    # "percentile" | "zscore" | "fixed"
PCT_CUTS = (0.25, 0.10)         # percentile: 하위 25% 이하 관심군, 하위 10% 이하 위험군
Z_CUTS = (1.0, 2.0)             # zscore: 평균 - 1.0σ 이하 관심군, 평균 - 2.0σ 이하 위험군
FIXED_CUTS = (60, 45)           # fixed : 관심군 상한점, 위험군 상한점 (예시값 - 분포를 본 뒤 조정)
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
DIM_LABEL = {COL_REGION: "지역", COL_INDUSTRY: "업종"}


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

def _to_standard(d):
    """신한 데이터를 표준 컬럼만 남기고, 기준일자(TA_YMD)·지역·업종·성별·연령대 단위로 합산한다.
    시간대(TIME_GB)는 쓰지 않으므로 합산하며, 이때 시간대만 달라서 생기던 '가짜 완전중복 행'도 사라진다."""
    d = d[REQUIRED_COLS].copy()
    for c in [COL_AMOUNT, COL_COUNT]:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    keys = [c for c in REQUIRED_COLS if c not in (COL_AMOUNT, COL_COUNT)]
    return d.groupby(keys, dropna=False, as_index=False)[[COL_AMOUNT, COL_COUNT]].sum(min_count=1)


def load_all_excels(folder):
    """모든 엑셀/CSV 파일을 읽어 신한카드 데이터와 SK 유동인구 데이터로 분리·통합한다."""
    files = find_excel_files(folder)
    log, frames, sk_frames = [], [], []

    code_dtype = {c: str for c in [COL_DATE, COL_REGION, COL_INDUSTRY, COL_SEX, COL_AGE,
                                   "STD_YM", "BLOCK_CD", "CLN_SGG_CD"]}

    for path in files:
        name = os.path.basename(path)
        try:
            if name.lower().endswith(".csv"):
                try:
                    raw = pd.read_csv(path, sep=r'[,|]', engine="python", dtype=code_dtype, encoding="utf-8-sig")
                except UnicodeDecodeError:
                    raw = pd.read_csv(path, sep=r'[,|]', engine="python", dtype=code_dtype, encoding="cp949")
            else:
                raw = pd.read_excel(path, sheet_name=SHEET_NAME, dtype=code_dtype)
        except Exception as e:
            print(f"  [오류] {name}: 읽지 못함 -> {type(e).__name__}: {e}")
            log.append({"파일": name, "시트": "-", "상태": "읽기 실패", "행수": 0,
                        "비고": f"{type(e).__name__}: {e}"})
            continue

        if isinstance(raw, pd.DataFrame):
            raw = {str(SHEET_NAME): raw}

        for sheet, d in raw.items():
            d = d.copy()
            d.columns = [str(c).strip().upper() for c in d.columns]

            # ---- SK 유동인구 데이터 ----
            # 성별·연령대별 파일(flow_age_pop)만 원본 형태 그대로 사용하고, 시간대/요일 파일은 쓰지 않는다.
            if "STD_YM" in d.columns and any("FLOW_POP_CNT" in c or c.startswith("TMST_") for c in d.columns):
                flow_cols = [c for c in d.columns if SK_FLOW_RE.match(c)]
                if USE_SK_DATA and flow_cols and "BLOCK_CD" in d.columns:
                    sk_frames.append(d[["STD_YM", "BLOCK_CD"] + flow_cols].copy())
                    print(f"  [읽음-SK] {name} / {sheet}: {len(d):,}행")
                else:
                    print(f"  [건너뜀-SK] {name} / {sheet}: 성별·연령대 컬럼 없음(또는 SK 사용 안 함)")
                continue

            # ---- 신한카드 데이터 ----
            label = "신한"
            if COL_REGION not in d.columns and "CLN_SGG_CD" in d.columns:
                d[COL_REGION] = d["CLN_SGG_CD"]     # 가맹점 지역이 없고 고객 거주지만 있을 때만 대체
                label = "신한-거주지로 대체"
            miss = [c for c in REQUIRED_COLS if c not in d.columns]
            if miss:
                print(f"  [건너뜀] {name} / {sheet}: 필요한 컬럼 없음 -> {miss}")
                log.append({"파일": name, "시트": sheet, "상태": "건너뜀(구조 불일치)", "행수": len(d),
                            "비고": f"없는 컬럼: {miss}"})
                continue
            d_std = _to_standard(d)
            d_std["SOURCE_FILE"] = name
            d_std["SOURCE_SHEET"] = str(sheet)
            frames.append(d_std)
            print(f"  [읽음-{label}] {name} / {sheet}: 원본 {len(d):,}행 -> 합산 후 {len(d_std):,}행")
            log.append({"파일": name, "시트": sheet, "상태": f"읽음({label})", "행수": len(d),
                        "비고": f"합산 후 {len(d_std):,}행"})

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

    # 문자열로 변환하고 앞뒤 공백 제거
    t = s.astype("string").str.strip()

    # 엑셀 숫자형 연월에서 발생하는 .0 제거
    t = t.str.replace(r"\.0$", "", regex=True)

    # 날짜 구분 기호 제거
    t = t.str.replace(r"[-/.\s]", "", regex=True)

    out = pd.Series(
        pd.NaT,
        index=s.index,
        dtype="datetime64[ns]"
    )

    # YYYYMM 형태: 예) 202601
    is_ym = t.str.fullmatch(r"\d{6}", na=False)
    out.loc[is_ym] = pd.to_datetime(
        t.loc[is_ym],
        format="%Y%m",
        errors="coerce"
    )

    # YYYYMMDD 형태: 예) 20260131
    is_ymd = t.str.fullmatch(r"\d{8}", na=False)
    out.loc[is_ymd] = pd.to_datetime(
        t.loc[is_ymd],
        format="%Y%m%d",
        errors="coerce"
    )

    # 나머지 날짜 형식 처리
    left = out.isna() & s.notna()
    if left.any():
        out.loc[left] = pd.to_datetime(
            s.loc[left],
            errors="coerce"
        )

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
    for c in [COL_REGION, COL_INDUSTRY, COL_SEX, COL_AGE]:
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
    for c in [COL_REGION, COL_INDUSTRY, COL_SEX]:
        n = int(df[c].isna().sum())
        if n:
            add(f"'미상' 처리: {c}", n)
        df[c] = df[c].fillna("미상")
    df[COL_AMOUNT] = df[COL_AMOUNT].fillna(0)

    # 중복
    key_cols = [COL_DATE, COL_REGION, COL_INDUSTRY, COL_SEX, COL_AGE]
    add("키(날짜·지역·업종·성별·연령대) 중복 행", int(df.duplicated(key_cols).sum()),
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


def _cos_dist(a, b):
    """마지막 축(업종)을 벡터로 보고 1 - 코사인 유사도. 한쪽이 전부 0이면 NaN."""
    na, nb = np.linalg.norm(a, axis=-1), np.linalg.norm(b, axis=-1)
    with np.errstate(divide="ignore", invalid="ignore"):
        sim = (a * b).sum(axis=-1) / (na * nb)
    return np.where((na > 0) & (nb > 0), np.clip(1 - sim, 0, None), np.nan)


def pattern_metrics(df, keys, months):
    """업종 분포 기반 활동 양식 변화와 지속성.
    반환: (분석단위별 표, 사용한 '달랐던 달' 기준값). 계산 불가면 NaN 표."""
    idx = keys.index
    cols = ["pat_change", "pat_run", "pat_ratio"]
    rec_m = list(months[-RECENT_MONTHS:])
    out = pd.DataFrame(np.nan, index=idx, columns=cols + [f"dist_{m}" for m in rec_m])
    n = len(months)
    if n < RECENT_MONTHS + MIN_BASELINE_MONTHS or COL_INDUSTRY in SEG:
        return out, np.nan
    base_m = list(months[:-RECENT_MONTHS])
    if BASELINE_MONTHS:
        base_m = base_m[-BASELINE_MONTHS:]
    inds = sorted(df[COL_INDUSTRY].dropna().unique())
    g = (df.groupby(SEG + [COL_INDUSTRY, "MONTH"])[ACT].sum()
           .unstack(["MONTH", COL_INDUSTRY]))
    g = g.reindex(columns=pd.MultiIndex.from_product([months, inds]), fill_value=0).fillna(0)
    g = g.reindex(index=pd.MultiIndex.from_frame(keys[SEG]), fill_value=0)
    arr = np.clip(g.values.reshape(len(keys), n, len(inds)), 0, None)    # (분석단위, 월, 업종)
    pos = {m: i for i, m in enumerate(months)}
    past = arr[:, [pos[m] for m in base_m], :].sum(axis=1)
    recent = arr[:, [pos[m] for m in rec_m], :].sum(axis=1)
    change = _cos_dist(recent, past)                                      # 최근 vs 과거 분포 차이
    d_rec = np.column_stack([_cos_dist(arr[:, pos[m], :], past) for m in rec_m])
    if PATTERN_DIFF_THRESHOLD == "auto":
        # 과거의 각 달을 '그 달을 뺀 나머지 과거'와 비교해 평소 흔들림을 잰다(자기 자신이 섞여 거리가 작아지는 것을 방지)
        d_base = np.concatenate([_cos_dist(arr[:, pos[m], :], past - arr[:, pos[m], :]) for m in base_m])
        d_base = d_base[~np.isnan(d_base)]
        thr = float(np.quantile(d_base, PATTERN_DIFF_QUANTILE)) if len(d_base) else np.nan
    else:
        thr = float(PATTERN_DIFF_THRESHOLD)
    differ = np.nan_to_num(d_rec, nan=-1) > thr if not np.isnan(thr) else np.zeros_like(d_rec, dtype=bool)
    run = np.cumprod(differ[:, ::-1], axis=1).sum(axis=1).astype(float)   # 가장 최근 달부터 연속 개수
    run[np.isnan(change)] = np.nan
    out["pat_change"], out["pat_run"], out["pat_ratio"] = change, run, run / RECENT_MONTHS
    for j, m in enumerate(rec_m):
        out[f"dist_{m}"] = d_rec[:, j]
    return out, thr


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



# ---------------------------------------------------------------------
# 5. 구조 지표: 활동 다양성
# ---------------------------------------------------------------------
def structure_metrics(df, keys, struct_months):
    """업종별 활동 분포의 정규화 엔트로피(다양성)를 계산한다. (편중 지표는 사용하지 않음)"""
    d = df[df["MONTH"].isin(struct_months)]
    dims = [c for c in DIVERSITY_DIMS if c not in SEG and df[c].nunique() > 1]
    res = keys[SEG].copy()
    for dim in dims:
        K = df[dim].nunique()
        g = d.groupby(SEG + [dim])[ACT].sum().reset_index()
        g = g[g[ACT] > 0]
        g["p"] = g[ACT] / g.groupby(SEG)[ACT].transform("sum")
        g["plogp"] = -g["p"] * np.log(g["p"])
        agg = g.groupby(SEG).agg(H=("plogp", "sum"), NUSED=("p", "size")).reset_index()
        agg["HN"] = agg["H"] / np.log(K)           # 0(한 곳에 몰림) ~ 1(완전히 고르게)
        agg = agg.rename(columns={c: f"{c}_{dim}" for c in ["H", "NUSED", "HN"]})
        agg[f"K_{dim}"] = K
        res = res.merge(agg, on=SEG, how="left")
    if dims:
        res["div_raw"] = np.mean([res[f"HN_{d}"] for d in dims], axis=0)   # 클수록 다양함
    else:
        res["div_raw"] = np.nan
    return res, dims


# ---------------------------------------------------------------------
# 6. 점수 / 분류
# ---------------------------------------------------------------------
def level3(p, hi="높음", lo="낮음"):
    if pd.isna(p):
        return "판단불가"
    return lo if p <= 1 / 3 else (hi if p >= 2 / 3 else "보통")


def cutoffs(scores, mode):
    """(관심군 기준점, 위험군 기준점). 안정지수가 이 값 '이하'이면 해당 군. 관심군 기준점 >= 위험군 기준점."""
    s = scores.dropna()
    if mode == "zscore":
        return (s.mean() - Z_CUTS[0] * s.std(), s.mean() - Z_CUTS[1] * s.std())
    if mode == "percentile":
        return (s.quantile(PCT_CUTS[0]), s.quantile(PCT_CUTS[1]))
    return FIXED_CUTS


def classify(score, cuts):
    out = pd.Series(np.where(score <= cuts[1], LABELS[2],
                    np.where(score <= cuts[0], LABELS[1], LABELS[0])), index=score.index)
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

    # ⑧ 활동량: 최근 구간 일평균 이용건수의 '연령대 내 백분위' (높을수록 활동적 = 안정)
    S["recent_act"] = M[rcols].mean(axis=1)
    S["recent_cnt"] = Mcnt[rcols].mean(axis=1)
    S["recent_amt"] = Mamt[rcols].mean(axis=1)
    S["act_pct"], S["peer_scope"] = peer_percentile(S["recent_act"], peers, elig)
    S["act_score"] = S["act_pct"]

    # ⑨ 활동 다양성: 업종 분포 엔트로피의 '연령대 내 백분위' (높을수록 다양 = 안정)
    print("[4] 활동 다양성 계산")
    struct_months = months[-RECENT_MONTHS:] if STRUCTURE_WINDOW == "recent" else months
    st, dims = structure_metrics(df, keys, struct_months)
    S = pd.concat([S, st.drop(columns=SEG)], axis=1)
    print(f"    사용한 차원: {[DIM_LABEL.get(d, d) for d in dims]} "
          f"(분석 단위에 포함되어 제외: {[DIM_LABEL.get(d, d) for d in DIVERSITY_DIMS if d in SEG]})")
    S["div_score"], _ = peer_percentile(S["div_raw"], peers, elig)

    # ⑩ 활동 양식 변화 / 지속성 (업종 분포: 과거 구간 vs 최근 구간)
    print("[5] 활동 양식 변화·지속성 계산 (업종 분포 비교)")
    pat, pat_thr = pattern_metrics(df, keys, months)
    S["pat_change"], S["pat_run"], S["pat_ratio"] = pat["pat_change"], pat["pat_run"], pat["pat_ratio"]
    S["chg_score"], _ = peer_percentile(S["pat_change"], peers, elig)   # 차이가 클수록 높음(안정)
    S["persist_score"] = S["pat_ratio"].where(elig)                    # 달랐던 달이 길수록 높음(안정)
    # 참고용(점수 미반영): 활동량(이용건수) 자체의 증감
    cm_raw = change_metrics(M)
    S["base_act"], S["change_pct"] = cm_raw["base"], cm_raw["change_pct"]
    S["persist_months"] = cm_raw["persist_months"]
    print(f"    '달랐던 달' 기준(업종 분포 거리): {pat_thr:.4f}" if not pd.isna(pat_thr) else
          "    [주의] 월 수 부족 또는 업종이 분석 단위에 포함되어 활동 변화를 계산하지 못했습니다.")

    # SK 유동인구: 점수에는 넣지 않고, 참고용 변화율만 계산해 함께 보여준다
    sk_table, sk_notes = None, []
    S["sk_change_pct"], S["sk_drop"] = np.nan, np.nan
    if sk_raw is not None:
        sk_table, SKM, sk_notes = build_sk(sk_raw, keys, months)
        if SKM is not None:
            cm_sk = change_metrics(SKM)
            S["sk_change_pct"], S["sk_drop"] = cm_sk["change_pct"], cm_sk["drop"]
    else:
        sk_notes.append("SK 유동인구 데이터 없음")
    for m in sk_notes:
        print("    " + m)

    # ⑪ 안정지수 = 4개 항목(각 0~100점)의 단순 평균. 계산 가능한 항목만 평균
    print("[6] 안정지수 계산 (가중치 없음: 4개 항목 단순 평균, 높을수록 안정)")
    COMP = {"활동량": "act_score", "활동 다양성": "div_score",
            "활동 변화": "chg_score", "변화 지속": "persist_score"}
    comp = pd.DataFrame({k: S[v] * 100 for k, v in COMP.items()})
    comp.loc[~elig, :] = np.nan
    S["score"] = comp.mean(axis=1, skipna=True)
    for k in COMP:
        S[f"pts_{k}"] = comp[k]
    S["main_driver"] = comp.fillna(np.inf).idxmin(axis=1).where(S["score"].notna(), "-")   # 가장 낮은 항목
    S["change_ok"] = S["pat_change"].notna()

    # ⑫ 분류 (점수가 낮을수록 위험)
    cuts = cutoffs(S["score"], CLASSIFY_MODE)
    S["class"] = classify(S["score"], cuts)
    sc = S["score"].dropna()
    print(f"\n[7] 안정지수 분포: 평균 {sc.mean():.1f} / 중앙값 {sc.median():.1f} / 표준편차 {sc.std():.1f} "
          f"/ 최소 {sc.min():.1f} / 최대 {sc.max():.1f}")
    print(f"    분류 기준({CLASSIFY_MODE}): 관심군 {cuts[0]:.1f}점 이하, 위험군 {cuts[1]:.1f}점 이하")

    # 해설 문장
    def reason(r):
        if pd.isna(r["score"]):
            return "표본 부족(전체 이용건수가 MIN_SEGMENT_TOTAL 미만)으로 판단하지 않음"
        parts = [f"활동량: {level3(r['act_pct'])}"]
        if dims:
            parts.append(f"활동 다양성: {level3(r['div_score'])}")
        if pd.notna(r["pat_change"]):
            parts.append(f"활동 양식 변화: {level3(r['chg_score'])} (업종 분포 차이 {r['pat_change']:.3f})")
            parts.append(f"변화 지속: {int(r['pat_run'])}개월" if r["pat_run"] > 0 else "변화 지속: 없음")
        else:
            parts.append("활동 양식 변화: 판단불가(기간 부족/활동 없음) -> 나머지 지표로 환산")
        if pd.notna(r["change_pct"]):
            way = "감소" if r["change_pct"] < 0 else "증가"
            parts.append(f"(참고) 활동량 {abs(r['change_pct']):.0%} {way}")
        if pd.notna(r["sk_change_pct"]):
            way = "감소" if r["sk_change_pct"] < 0 else "증가"
            both = " -> 카드·유동인구 동시 감소" if (r["sk_change_pct"] < 0 and pd.notna(r["change_pct"])
                                                  and r["change_pct"] < 0) else ""
            parts.append(f"유동인구(SK) {abs(r['sk_change_pct']):.0%} {way}{both}")
        parts.append(f"가장 낮은 항목: {r['main_driver']}")
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
        out = g.agg(분석단위수=("score", "size"), 평균_안정지수=("score", "mean"),
                    중앙값=("score", "median"), 표준편차=("score", "std"),
                    위험군_비율=("is_risk", "mean"), 관심군_비율=("is_watch", "mean"),
                    활동량_점수=("pts_활동량", "mean"), 다양성_점수=("pts_활동 다양성", "mean"),
                    변화_점수=("pts_활동 변화", "mean"), 지속_점수=("pts_변화 지속", "mean"),
                    평균_활동변화율=("change_pct", "mean"),
                    평균_유동인구변화율=("sk_change_pct", "mean"))
        out["활동량가중_평균_안정지수"] = g.apply(
            lambda x: np.average(x["score"], weights=x["w"]) if x["w"].sum() > 0 else np.nan)
        out["관심군_이상_비율"] = out["위험군_비율"] + out["관심군_비율"]
        out["위험군_비율"] *= 100
        out["관심군_비율"] *= 100
        out["관심군_이상_비율"] *= 100
        out["평균_활동변화율"] *= 100
        out["평균_유동인구변화율"] *= 100
        if out["평균_유동인구변화율"].isna().all():      # SK 데이터가 없으면 열 제거
            out = out.drop(columns=["평균_유동인구변화율"])
        out = out.reset_index()
        return out.sort_values(by, key=lambda s: age_sort_key(s) if s.name == PEER_COL else s)

    age_tbl = summarize([PEER_COL])
    sex_tbl = summarize([PEER_COL, COL_SEX]) if COL_SEX in S.columns else None
    print("\n[8] 연령대별 결과")
    show = age_tbl[[PEER_COL, "분석단위수", "평균_안정지수", "위험군_비율", "관심군_비율", "평균_활동변화율"]]
    print(show.round(1).to_string(index=False))

    # ⑭ 엑셀 저장
    print(f"\n[9] 결과 저장: {OUTPUT_FILE}")
    S_sorted = S.sort_values(["age_order", "score"], ascending=[True, True])   # 점수 낮은(위험한) 순
    seg_cols = SEG
    mlabels = [str(m) for m in months]

    t_all = S_sorted[seg_cols + ["score", "class", "pts_활동량", "pts_활동 다양성", "pts_활동 변화",
                                 "pts_변화 지속", "sk_change_pct", "main_driver", "reason"]].copy()
    t_all["sk_change_pct"] *= 100
    t_all.columns = seg_cols + ["안정지수", "분류", "활동량 점수", "다양성 점수", "활동 변화 점수",
                                "변화 지속 점수", "유동인구 변화율(%)", "가장 낮은 항목", "해석"]
    if t_all["유동인구 변화율(%)"].isna().all():
        t_all = t_all.drop(columns=["유동인구 변화율(%)"])
    t_all = t_all.sort_values(["안정지수"], ascending=True, na_position="last")

    t_act = S_sorted[seg_cols + ["total_use", "recent_cnt", "recent_amt", "recent_act", "act_pct",
                                 "peer_scope", "act_score"]].copy()
    t_act.columns = seg_cols + ["전체 이용건수", "최근 일평균 이용건수", "최근 일평균 TS_AT",
                                f"최근 일평균 {ACT}(기준)", "연령대 내 백분위(0~1)", "비교 기준", "활동량 점수(0~1)"]

    div_cols = []
    for d in dims:
        div_cols += [f"HN_{d}", f"NUSED_{d}", f"K_{d}"]
    t_div = S_sorted[seg_cols + div_cols + ["div_raw", "div_score"]].copy()
    t_div.columns = (seg_cols + sum([[f"{DIM_LABEL.get(d, d)} 다양성(정규화 엔트로피 0~1)",
                                      f"{DIM_LABEL.get(d, d)} 사용 개수", f"{DIM_LABEL.get(d, d)} 전체 개수"]
                                     for d in dims], [])
                     + ["다양성(평균,0~1)", "다양성 점수(연령대 내 백분위, 0~1)"])

    dcols = [c for c in pat.columns if c.startswith("dist_")]
    t_chg = S_sorted[seg_cols + ["pat_change", "chg_score", "pat_run", "persist_score"]].copy()
    t_chg.columns = seg_cols + ["업종 분포 차이(1-코사인, 최근 vs 과거)", "활동 변화 점수(연령대 내 백분위,0~1)",
                                "분포가 달랐던 연속 개월", "변화 지속 점수(0~1)"]
    dm = pat.loc[S_sorted.index, dcols].copy()
    dm.columns = [f"월 분포 차이_{c[5:]}" for c in dcols]
    ref = S_sorted[["base_act", "recent_act", "change_pct", "persist_months"]].copy()
    ref["change_pct"] *= 100
    ref.columns = ["(참고)과거 일평균 활동량", "(참고)최근 일평균 활동량", "(참고)활동량 변화율(%)", "(참고)활동량 감소 지속 개월"]
    mon = M.loc[S_sorted.index].copy()
    mon.columns = [f"일평균_{m}" for m in mlabels]
    t_chg = pd.concat([t_chg, dm, ref, mon], axis=1)

    ct = pd.crosstab(S.loc[S["score"].notna(), PEER_COL], S.loc[S["score"].notna(), "class"])
    ct = ct.reindex(columns=list(LABELS), fill_value=0)
    ct.insert(0, "합계", ct.sum(axis=1))
    ct = ct.reset_index().sort_values(PEER_COL, key=age_sort_key)
    t_cls = S_sorted[S_sorted["class"].isin([LABELS[1], LABELS[2]])][seg_cols + ["score", "class", "reason"]]
    t_cls.columns = seg_cols + ["안정지수", "분류", "해석"]
    t_cls = t_cls.sort_values("안정지수", ascending=True)

    # 설정/분포 시트
    desc = S.loc[S["score"].notna(), ["score", "act_score", "div_score", "chg_score", "persist_score",
                                      "pat_change", "change_pct"]].copy()
    desc.columns = ["안정지수", "활동량 점수(0~1)", "다양성 점수(0~1)", "활동 변화 점수(0~1)",
                    "변화 지속 점수(0~1)", "업종 분포 차이", "(참고)활동량 변화율"]
    dist = desc.describe(percentiles=[.1, .25, .5, .75, .9]).T.reset_index().rename(columns={"index": "지표"})
    by_age = S[S["score"].notna()].groupby(PEER_COL)["score"].describe(percentiles=[.1, .25, .5, .75]).reset_index()
    by_age = by_age.sort_values(PEER_COL, key=age_sort_key)

    cand = []
    for name, mode in [("현재 설정(" + CLASSIFY_MODE + ")", CLASSIFY_MODE),
                       (f"고정({FIXED_CUTS[0]}/{FIXED_CUTS[1]})", "fixed"),
                       (f"평균-{Z_CUTS[0]}σ / -{Z_CUTS[1]}σ", "zscore"),
                       (f"하위 {int(PCT_CUTS[0]*100)}% / {int(PCT_CUTS[1]*100)}%", "percentile")]:
        c = cutoffs(S["score"], mode)
        cl = classify(S["score"].dropna(), c)
        n = max(len(cl), 1)
        cand.append({"기준 후보": name, "관심군 기준점(이하)": c[0], "위험군 기준점(이하)": c[1],
                     "정상 비율(%)": (cl == LABELS[0]).sum() / n * 100,
                     "관심군 비율(%)": (cl == LABELS[1]).sum() / n * 100,
                     "위험군 비율(%)": (cl == LABELS[2]).sum() / n * 100})
    cand = pd.DataFrame(cand)

    setting = pd.DataFrame({"항목": [
        "분석 단위", "활동량 기준 컬럼", "최근 구간(개월)", "과거 구간(개월)", "연령대 내 비교 최소 단위 수",
        "활동 변화 방식", "'달랐던 달' 기준(업종 분포 거리)", "(미사용)", "안정지수 계산", "분류 방식",
        "분석한 월", "분석 단위 수(전체/판단 가능)", "SK 유동인구(참고용, 점수 미반영)"],
        "값": [" x ".join(SEG), ACT, RECENT_MONTHS, BASELINE_MONTHS or "이전 전체", MIN_PEERS,
               "최근 vs 과거 업종 분포 차이(1-코사인 유사도)",
               ("-" if pd.isna(pat_thr) else f"{pat_thr:.4f} ({PATTERN_DIFF_THRESHOLD}, P{int(PATTERN_DIFF_QUANTILE*100)})"), "-",
               "활동량·다양성·활동 변화·변화 지속 4개 항목(각 0~100점) 단순 평균, 높을수록 안정",
               f"{CLASSIFY_MODE}: 관심군 {cuts[0]:.1f}점 이하, 위험군 {cuts[1]:.1f}점 이하",
               f"{months[0]} ~ {months[-1]} ({n_m}개월)", f"{len(S)} / {int(S['score'].notna().sum())}",
               ("사용 - " + "; ".join(sk_notes)) if sk_table is not None else "사용 안 함"]})
    caveat = pd.DataFrame({"해석 시 유의사항": [
        "이 결과는 사회적 고립을 진단하는 것이 아니라, 위험 가능성이 높은 집단을 1차로 탐색하는 용도입니다.",
        "데이터에 개인 ID가 없어 개인이 아닌 '분석 단위(집계 집단)'의 값입니다.",
        "MCT_SGG_CD는 '가맹점 지역'이므로 거주지가 아니라 소비가 일어난 지역입니다. (파일 로그에서 거주지로 대체된 파일이 있는지 확인)",
        "시간대(TIME_GB)는 사용하지 않았고, 기준일자는 TA_YMD를 사용했습니다.",
        "안정지수는 높을수록 활동적·안정, 낮을수록 위험 가능성이 큽니다. 분류 기준은 확정값이 아니며 분포를 확인한 뒤 조정해야 합니다.",
        "분석 단위의 규모(인구)를 알 수 없어 활동량 '수준'은 규모 영향을 받습니다. 그래서 자기 자신의 과거 대비 '변화'를 함께 봅니다."]})

    for d_ in [t_all] + ([sk_table] if sk_table is not None else []) + [t_act, t_div, t_chg, t_cls, age_tbl, desc, dist, by_age, cand]:
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
        t_all.to_excel(w, sheet_name="전체_안정지수", index=False)
        age_tbl.to_excel(w, sheet_name="연령대별_분석", index=False)
        t_act.to_excel(w, sheet_name="활동량", index=False)
        t_div.to_excel(w, sheet_name="활동_다양성", index=False)
        t_chg.to_excel(w, sheet_name="활동_변화", index=False)
        write_blocks(w, "위험군_분류", [
            (f"연령대별 분류 건수 (기준: 관심군 {cuts[0]:.1f}점 이하, 위험군 {cuts[1]:.1f}점 이하)", ct),
            ("관심군/위험군 분석 단위 목록", t_cls)])
        if sex_tbl is not None:
            sex_tbl.to_excel(w, sheet_name="연령대_성별", index=False)
        write_blocks(w, "설정_분포", [
            ("해석 시 유의사항", caveat), ("사용한 설정", setting),
            ("지표 분포(분위수)", dist), ("연령대별 안정지수 분포", by_age),
            ("분류 기준 후보 비교", cand)])
        if sk_table is not None:
            sk_table.to_excel(w, sheet_name="유동인구_변화", index=False)
        quality.to_excel(w, sheet_name="데이터_점검", index=False)
        pd.DataFrame(load_log).to_excel(w, sheet_name="파일_로그", index=False)
    style_workbook(OUTPUT_FILE, block_sheets)
    print("    완료. 시트: 전체_안정지수, 연령대별_분석, 활동량, 활동_다양성, 활동_변화, 위험군_분류, "
          "연령대_성별, 설정_분포, 데이터_점검, 파일_로그")


if __name__ == "__main__":
    main()
