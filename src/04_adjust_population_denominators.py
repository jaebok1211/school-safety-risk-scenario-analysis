"""위험 시나리오에 실제 학생 수·학교 수 모수를 적용해 신고부담을 보정한다."""

from pathlib import Path

import numpy as np
import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / "outputs"

STUDENT_FILE = DATA_DIR / "2025_연도별 학생수.xlsx"
SCHOOL_FILE = DATA_DIR / "2025_연도별 학교수.xlsx"
RULE_FILE = OUTPUT_DIR / "2단계_다차원_위험_시나리오_결과.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "전학교급_성별_실제모수보정_위험_시나리오.xlsx"

YEARS = [2021, 2022, 2023, 2024, 2025]
TOTAL_TRANSACTIONS = 812_363
MIN_SUPPORT_PERCENT = 0.50


def get_exact_denominators(condition_str: str) -> tuple[str, str]:
    """시나리오 조건에 맞는 학생 수·학교 수 분모 컬럼을 반환한다."""
    cond = str(condition_str)
    is_female = "여" in cond or "여학생" in cond
    is_male = "남" in cond or "남학생" in cond

    if "초등학교" in cond and "중학교" not in cond and "고등학교" not in cond:
        school_col = "학교_초등"
        student_col = "초등_여" if is_female else ("초등_남" if is_male else "초등_전체")
    elif "중학교" in cond:
        school_col = "학교_중학"
        student_col = "중학_여" if is_female else ("중학_남" if is_male else "중학_전체")
    elif "고등학교" in cond:
        school_col = "학교_고등"
        student_col = "고등_여" if is_female else ("고등_남" if is_male else "고등_전체")
    else:
        school_col = "학교_초중고_전체"
        student_col = "초중고_여" if is_female else ("초중고_남" if is_male else "초중고_전체")

    return student_col, school_col


def normalize_rule_key(row) -> str:
    """조건/결과 순서 차이로 생기는 중복 규칙을 비교하기 위한 키를 만든다."""
    condition = ", ".join(
        sorted(item.strip() for item in str(row["위험상황(조건)"]).split(","))
    )
    result = ", ".join(
        sorted(item.strip() for item in str(row["사고결과"]).split(","))
    )
    return f"{condition} ➡️ {result}"


def safe_float(value) -> float:
    value_str = str(value).replace("건", "").replace(",", "").strip()
    if value_str in {"-", "nan", "None", ""}:
        return 0.0
    try:
        return float(value_str)
    except (TypeError, ValueError):
        return 0.0


def main() -> None:
    required_files = [STUDENT_FILE, SCHOOL_FILE, RULE_FILE]
    missing = [path for path in required_files if not path.exists()]
    if missing:
        raise FileNotFoundError("필요한 입력 파일이 없습니다:\n" + "\n".join(map(str, missing)))

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    df_students = pd.read_excel(STUDENT_FILE, sheet_name="Sheet0")
    df_schools = pd.read_excel(SCHOOL_FILE, sheet_name="Sheet0")

    student_national = df_students[
        df_students["연도"].astype(str).isin([str(year) for year in YEARS])
        & (df_students["시도"] == "전국")
    ].copy()
    school_national = df_schools[
        df_schools["연도"].astype(str).isin([str(year) for year in YEARS])
        & (df_schools["시도"] == "전국")
    ].copy()

    pop_full = pd.DataFrame(
        {
            "연도": YEARS,
            "초등_전체": student_national["초등학교"].astype(float).values,
            "초등_여": student_national["초등학교.1"].astype(float).values,
            "중학_전체": student_national["중학교"].astype(float).values,
            "중학_여": student_national["중학교.1"].astype(float).values,
            "고등_전체": student_national["고등학교"].astype(float).values,
            "고등_여": student_national["고등학교.1"].astype(float).values,
            "학교_초등": school_national["초등학교"].astype(float).values,
            "학교_중학": school_national["중학교"].astype(float).values,
            "학교_고등": school_national["고등학교"].astype(float).values,
        }
    ).set_index("연도")

    pop_full["초등_남"] = pop_full["초등_전체"] - pop_full["초등_여"]
    pop_full["중학_남"] = pop_full["중학_전체"] - pop_full["중학_여"]
    pop_full["고등_남"] = pop_full["고등_전체"] - pop_full["고등_여"]

    pop_full["초중고_전체"] = (
        pop_full["초등_전체"] + pop_full["중학_전체"] + pop_full["고등_전체"]
    )
    pop_full["초중고_여"] = (
        pop_full["초등_여"] + pop_full["중학_여"] + pop_full["고등_여"]
    )
    pop_full["초중고_남"] = (
        pop_full["초등_남"] + pop_full["중학_남"] + pop_full["고등_남"]
    )
    pop_full["학교_초중고_전체"] = (
        pop_full["학교_초등"] + pop_full["학교_중학"] + pop_full["학교_고등"]
    )

    df_rules = pd.read_excel(RULE_FILE)
    df_rules["rule_key"] = df_rules.apply(normalize_rule_key, axis=1)
    df_rules = df_rules.drop_duplicates(subset=["rule_key"]).copy()

    support_col = "support" if "support" in df_rules.columns else "지지도(%)"
    multiplier = 1.0 if support_col == "support" else 0.01
    df_rules["원본_전체건수"] = (
        df_rules[support_col] * multiplier * TOTAL_TRANSACTIONS
    ).round().astype(int)

    df_rules["지지도(%)"] = np.round(
        (df_rules["원본_전체건수"] / TOTAL_TRANSACTIONS) * 100,
        2,
    )
    df_rules = df_rules[df_rules["지지도(%)"] >= MIN_SUPPORT_PERCENT].copy().reset_index(drop=True)

    has_yearly_cols = all(f"{year}년" in df_rules.columns for year in YEARS)

    student_rate_5yr_avg = []
    school_rate_5yr_avg = []

    for _, row in df_rules.iterrows():
        student_col, school_col = get_exact_denominators(row["위험상황(조건)"])

        if has_yearly_cols:
            yearly_counts = [safe_float(row[f"{year}년"]) for year in YEARS]
        else:
            yearly_counts = [safe_float(row["원본_전체건수"]) / 5.0] * 5

        yearly_student_rates = [
            (count / pop_full.loc[year, student_col]) * 10_000.0
            for year, count in zip(YEARS, yearly_counts)
        ]
        yearly_school_rates = [
            count / pop_full.loc[year, school_col]
            for year, count in zip(YEARS, yearly_counts)
        ]

        student_rate_5yr_avg.append(round(sum(yearly_student_rates) / 5.0, 2))
        school_rate_5yr_avg.append(round(sum(yearly_school_rates) / 5.0, 2))

    df_rules["학생_1만명당_연평균_발생건수"] = student_rate_5yr_avg
    df_rules["학교_1개교당_연평균_발생건수"] = school_rate_5yr_avg
    df_rules.drop(columns=["rule_key"], errors="ignore", inplace=True)

    df_rules.to_excel(OUTPUT_FILE, index=False)
    print(f"저장 완료: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
