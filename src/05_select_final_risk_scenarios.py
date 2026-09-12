"""지지도·신뢰도·향상도 기준을 적용해 최종 위험 시나리오를 선별한다."""

from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = PROJECT_ROOT / "outputs"

INPUT_FILE = OUTPUT_DIR / "전학교급_성별_실제모수보정_위험_시나리오.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "최종_위험_시나리오_선정보고서.xlsx"

MIN_CONFIDENCE = 0.30
MIN_LIFT = 1.50


def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {INPUT_FILE}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df_rules = pd.read_excel(INPUT_FILE)

    confidence_col = "confidence" if "confidence" in df_rules.columns else "신뢰도(%)"
    lift_col = "lift" if "lift" in df_rules.columns else "향상도(Lift)"

    confidence_threshold = (
        MIN_CONFIDENCE if df_rules[confidence_col].max() <= 1.0 else MIN_CONFIDENCE * 100
    )

    refined_rules = df_rules[
        (df_rules[confidence_col] >= confidence_threshold)
        & (df_rules[lift_col] >= MIN_LIFT)
    ].copy()

    if "support" in refined_rules.columns:
        refined_rules["지지도(%)"] = (refined_rules["support"] * 100).round(2)
    if "confidence" in refined_rules.columns:
        refined_rules["신뢰도(%)"] = (refined_rules["confidence"] * 100).round(2)
    if "lift" in refined_rules.columns:
        refined_rules["향상도(Lift)"] = refined_rules["lift"].round(2)

    refined_rules = refined_rules.sort_values(
        by=["향상도(Lift)", "지지도(%)"],
        ascending=[False, False],
    ).reset_index(drop=True)

    final_columns = [
        "위험상황(조건)",
        "사고결과",
        "원본_전체건수",
        "학생_1만명당_연평균_발생건수",
        "학교_1개교당_연평균_발생건수",
        "지지도(%)",
        "신뢰도(%)",
        "향상도(Lift)",
    ]
    available_columns = [col for col in final_columns if col in refined_rules.columns]
    final_report = refined_rules[available_columns]

    final_report.to_excel(OUTPUT_FILE, index=False)
    print(f"저장 완료: {OUTPUT_FILE} ({len(final_report):,}개 시나리오)")


if __name__ == "__main__":
    main()
