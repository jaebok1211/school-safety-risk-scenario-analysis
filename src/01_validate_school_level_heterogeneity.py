"""학교급별 사고시간·사고형태 분포의 차이를 비교한다.

GitHub 공개용 정리본:
- 개인 PC 절대경로를 제거하고 프로젝트 루트 기준 상대경로를 사용한다.
- 분석 로직과 산출물 이름은 기존 코드와 동일하게 유지한다.
"""

from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / "outputs"

INPUT_FILE = DATA_DIR / "학교안전사고 데이터.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "학교급별_이질성_증명분석.xlsx"

YEARS = ["2021", "2022", "2023", "2024", "2025"]
REQUIRED_COLUMNS = ["학교급", "사고시간", "사고형태"]


def main() -> None:
    """연도별 원자료를 통합해 학교급별 분포 비교표를 생성한다."""
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

        if all(col in df_year.columns for col in REQUIRED_COLUMNS):
            all_data.append(df_year[REQUIRED_COLUMNS].copy())
        else:
            missing = [col for col in REQUIRED_COLUMNS if col not in df_year.columns]
            print(f"[경고] {year} 시트에서 필요한 컬럼이 없습니다: {missing}")

    if not all_data:
        raise ValueError("분석 가능한 연도별 데이터가 없습니다.")

    total_df = pd.concat(all_data, ignore_index=True)
    for col in REQUIRED_COLUMNS:
        total_df[col] = total_df[col].astype(str).str.strip()

    with pd.ExcelWriter(OUTPUT_FILE) as writer:
        time_pivot = total_df.pivot_table(
            index="사고시간",
            columns="학교급",
            aggfunc="size",
            fill_value=0,
        )
        time_pivot.to_excel(writer, sheet_name="사고시간_맥락_비교")

        type_pivot = total_df.pivot_table(
            index="사고형태",
            columns="학교급",
            aggfunc="size",
            fill_value=0,
        )
        type_ratio = type_pivot.div(type_pivot.sum(axis=0), axis=1) * 100
        type_ratio.round(2).to_excel(writer, sheet_name="사고형태_비율_비교")

    print(f"저장 완료: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
