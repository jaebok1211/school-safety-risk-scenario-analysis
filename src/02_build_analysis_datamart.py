"""2021~2025 학교안전사고 원자료를 분석용 데이터마트로 변환한다.

주요 파생변수:
- 계절
- 학기 맥락
- 학교급별 학년군
"""

import re
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / "outputs"

INPUT_FILE = DATA_DIR / "학교안전사고 데이터.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "초중고_최종_분석용_데이터마트.xlsx"

YEARS = ["2021", "2022", "2023", "2024", "2025"]
TARGET_COLUMNS = [
    "지역",
    "학교급",
    "사고자구분",
    "사고자학년",
    "사고자성별",
    "사고연월",
    "사고요일",
    "사고시간",
    "사고장소",
    "사고당시활동",
    "사고형태",
    "사고부위",
]


def parse_semester_context(date_val):
    """사고연월에서 계절과 학기 맥락을 파생한다."""
    if pd.isna(date_val):
        return "미분류", "미분류"

    date_str = str(date_val).strip()
    try:
        match = re.search(r"-(\d{2})", date_str)
        month = (
            int(match.group(1))
            if match
            else int(re.findall(r"\d+", date_str)[1])
        )

        if month in [3, 4, 5]:
            season = "봄"
        elif month in [6, 7, 8]:
            season = "여름"
        elif month in [9, 10, 11]:
            season = "가을"
        else:
            season = "겨울"

        if month in [3, 4]:
            context = "1학기초(적응·행사)"
        elif month in [5, 6]:
            context = "1학기말(체육·밀집)"
        elif month in [7, 8]:
            context = "여름방학기"
        elif month in [9, 10]:
            context = "2학기초(행사·수학여행)"
        elif month in [11, 12]:
            context = "2학기말(마무리)"
        else:
            context = "겨울방학기"

        return season, context
    except (IndexError, TypeError, ValueError):
        return "미분류", "미분류"


def parse_grade_group(row):
    """학교급과 학년을 결합해 분석용 학년군을 생성한다."""
    school = str(row["학교급"]).strip()
    grade = str(row["사고자학년"]).strip()

    if pd.isna(row["사고자학년"]) or grade in ["nan", "미분류", ""]:
        return "미분류"

    if "1학년" in grade:
        return f"{school}_1학년(신입적응기)"
    if school == "초등학교":
        if "2학년" in grade or "3학년" in grade:
            return "초등학교_중학년(2-3)"
        if "4학년" in grade or "5학년" in grade or "6학년" in grade:
            return "초등학교_고학년(4-6)"
        return "초등학교_기타"
    if school in ["중학교", "고등학교"]:
        if "2학년" in grade:
            return f"{school}_2학년(중간학년)"
        if "3학년" in grade:
            return f"{school}_3학년(졸업·입시학년)"
        return f"{school}_기타"
    return "기타학년"


def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {INPUT_FILE}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    all_data = []

    for year in YEARS:
        try:
            df_year = pd.read_excel(INPUT_FILE, sheet_name=year)
        except Exception as exc:
            print(f"[경고] {year} 시트를 읽지 못했습니다: {exc}")
            continue

        available_cols = [col for col in TARGET_COLUMNS if col in df_year.columns]
        df_filtered = df_year[available_cols].copy()
        df_filtered["연도"] = f"{year}년"
        all_data.append(df_filtered)

    if not all_data:
        raise ValueError("분석 가능한 연도별 데이터가 없습니다.")

    total_df = pd.concat(all_data, ignore_index=True)
    for col in total_df.columns:
        if total_df[col].dtype == "object":
            total_df[col] = total_df[col].astype(str).str.strip()

    main_schools = ["초등학교", "중학교", "고등학교"]
    mart_df = total_df[total_df["학교급"].isin(main_schools)].copy()

    seasons_contexts = mart_df["사고연월"].apply(parse_semester_context)
    mart_df["파생_계절"] = [value[0] for value in seasons_contexts]
    mart_df["파생_학기맥락"] = [value[1] for value in seasons_contexts]
    mart_df["파생_학년군"] = mart_df.apply(parse_grade_group, axis=1)

    mart_df.to_excel(OUTPUT_FILE, index=False)
    print(f"저장 완료: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
