"""최종 위험 시나리오의 연도 반복성과 지역 확산성을 검증한다."""

from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = PROJECT_ROOT / "outputs"

DATAMART_FILE = OUTPUT_DIR / "초중고_최종_분석용_데이터마트.xlsx"
SCENARIO_FILE = OUTPUT_DIR / "최종_위험_시나리오_선정보고서.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "시나리오_재현성_검증보고서.xlsx"


def main() -> None:
    required_files = [DATAMART_FILE, SCENARIO_FILE]
    missing = [path for path in required_files if not path.exists()]
    if missing:
        raise FileNotFoundError("필요한 입력 파일이 없습니다:\n" + "\n".join(map(str, missing)))

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df_mart = pd.read_excel(DATAMART_FILE)
    df_scenarios = pd.read_excel(SCENARIO_FILE)

    for col in df_mart.columns:
        if df_mart[col].dtype == "object":
            df_mart[col] = df_mart[col].astype(str).str.strip()

    verification_results = []
    column_value_map = {
        col: set(df_mart[col].dropna().unique())
        for col in df_mart.columns
    }

    for _, row in df_scenarios.iterrows():
        condition_str = row["위험상황(조건)"]
        result_str = row["사고결과"]

        condition_items = [item.strip() for item in str(condition_str).split(",")]
        result_items = [item.strip() for item in str(result_str).split(",")]

        query_df = df_mart.copy()

        for item in condition_items:
            for col, value_set in column_value_map.items():
                if item in value_set:
                    query_df = query_df[query_df[col] == item]
                    break

        for item in result_items:
            for col in ["사고형태", "사고부위"]:
                if col in column_value_map and item in column_value_map[col]:
                    query_df = query_df[query_df[col] == item]
                    break

        years_present = query_df["연도"].astype(str).unique()
        year_score = len(years_present)
        years_list = ", ".join(sorted(years_present))

        regions_present = query_df["지역"].unique()
        region_score = len(regions_present)

        if year_score == 5:
            risk_type = "구조적 위험 (5년 연속 발생)"
        elif year_score >= 3:
            risk_type = "지속적 위험 (3~4년 반복)"
        elif any("2025" in year for year in years_present) and year_score <= 2:
            risk_type = "최근 신흥 위험 (최근 집중)"
        else:
            risk_type = "일시적/우연적 노이즈"

        verification_results.append(
            {
                "위험상황(조건)": condition_str,
                "사고결과": result_str,
                "원본_전체건수": len(query_df),
                "학생_1만명당_연평균_발생건수": row.get("학생_1만명당_연평균_발생건수", None),
                "학교_1개교당_연평균_발생건수": row.get("학교_1개교당_연평균_발생건수", None),
                "지지도(%)": row.get("지지도(%)", None),
                "신뢰도(%)": row.get("신뢰도(%)", None),
                "향상도(Lift)": row.get("향상도(Lift)", None),
                "지속연도수(5점만점)": year_score,
                "발생연도종류": years_list,
                "확산지역수(시도단위)": region_score,
                "위험시나리오_판정": risk_type,
            }
        )

    df_verified = pd.DataFrame(verification_results)
    df_verified = df_verified.sort_values(
        by=["지속연도수(5점만점)", "원본_전체건수"],
        ascending=[False, False],
    ).reset_index(drop=True)

    df_verified.to_excel(OUTPUT_FILE, index=False)
    print(f"저장 완료: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
