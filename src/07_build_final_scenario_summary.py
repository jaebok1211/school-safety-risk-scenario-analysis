"""재현성 검증 결과를 정제하고 보고서용 대표 시나리오 요약표를 생성한다.

주의:
보고서용 대표 시나리오 5선과 연도별 수치는 최종 보고서에서 확정한 값을
재현하기 위해 명시적으로 기록한 요약 테이블이다.
"""

from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = PROJECT_ROOT / "outputs"

INPUT_FILE = OUTPUT_DIR / "시나리오_재현성_검증보고서.xlsx"
MASTER_OUTPUT_FILE = OUTPUT_DIR / "최종_정제_위험시나리오_마스터_v5.xlsx"
SUMMARY_OUTPUT_FILE = OUTPUT_DIR / "최종_보고서용_4대_표준_요약표_v5.xlsx"


def clean_condition(condition_str: str) -> str:
    """세부 학년군이 있을 때 중복되는 상위 학교급 조건을 제거한다."""
    items = [item.strip() for item in str(condition_str).split(",")]
    has_specific = any(
        "중학교_" in item or "고등학교_" in item or "초등학교_" in item
        for item in items
    )

    cleaned = []
    for item in items:
        if has_specific and item in ["중학교", "고등학교", "초등학교"]:
            continue
        cleaned.append(item)
    return ", ".join(cleaned)


def build_analysis_standard_table() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "구분": "원자료 규모",
                "세부 내용": "초·중·고등학교 5개년(2021~2025년) 학교안전사고 데이터 약 81만 건 (유치원/특수학교 제외)",
            },
            {
                "구분": "전 학교급/성별 모수 보정",
                "세부 내용": "초·중·고 및 남/여 개별 연도별 교육통계 실제 학생 수 및 학교 수 동적 매핑 보정",
            },
            {
                "구분": "연관규칙 Cut-off 기준",
                "세부 내용": "최소 지지도(Support) 0.5% 이상, 최소 신뢰도(Confidence) 30% 이상, 최소 향상도(Lift) 1.5 이상",
            },
            {
                "구분": "포함관계 중복조건 정제",
                "세부 내용": "중복/포함 관계 조건 제거 및 Support 0.5% 미만 2개 제외 (최종 457개 유효 시나리오 정제)",
            },
            {
                "구분": "지역 확산성 판정 논리",
                "세부 내용": "지역별 최소 건수(100건) 및 Support 기준 적용 (상위 3개 전국 공통 반복위험, 하위 2개 광역 반복위험 분류)",
            },
            {
                "구분": "2025년 별도 기간 검증",
                "세부 내용": "Train(2021~2024년) 패턴 도출 후 Test(2025년) 일관성 및 순위 유지 검증",
            },
        ]
    )


def build_representative_scenario_table() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "선정유형": "1. 규모형 (Volume)",
                "대표 위험 시나리오": "초·중·고 공통 x 걷기/뛰기, 오르내리기 -> 넘어짐",
                "5개년 건수": "215,097건",
                "학생 1만 명당 연평균 사고 신고건수": "81.58건",
                "학교 1개교당 연평균 사고 신고건수": "3.65건",
                "Lift (향상도)": 1.98,
                "지지도(%)": 26.48,
                "신뢰도(%)": 52.37,
                "선정 이유 및 변경 내역": "전체 사고 절대 건수 1위 (모수 보정 및 지지도 26.48% 재계산 일치)",
            },
            {
                "선정유형": "2. 밀도형 (Density)",
                "대표 위험 시나리오": "중학교 여학생 x 전체 -> 손가락 상해",
                "5개년 건수": "35,171건",
                "학생 1만 명당 연평균 사고 신고건수": "107.48건",
                "학교 1개교당 연평균 사고 신고건수": "2.15건",
                "Lift (향상도)": 1.53,
                "지지도(%)": 4.33,
                "신뢰도(%)": 33.57,
                "선정 이유 및 변경 내역": "전체 457개 시나리오 중 학생 1만 명당 연평균 사고 신고건수 1위",
            },
            {
                "선정유형": "3. 학교부담형 (School)",
                "대표 위험 시나리오": "중학교 x 체육 수업 -> 손가락 상해",
                "5개년 건수": "46,147건",
                "학생 1만 명당 연평균 사고 신고건수": "68.39건",
                "학교 1개교당 연평균 사고 신고건수": "2.82건",
                "Lift (향상도)": 1.50,
                "지지도(%)": 5.68,
                "신뢰도(%)": 32.99,
                "선정 이유 및 변경 내역": "전체 457개 시나리오 중 학교 1개교당 연평균 사고 신고건수 1위",
            },
            {
                "선정유형": "4. 특이형 (Lift)",
                "대표 위험 시나리오": "중학교 여학생 x 체육 수업(농구) -> 손가락 부딪힘",
                "5개년 건수": "4,357건",
                "학생 1만 명당 연평균 사고 신고건수": "13.31건",
                "학교 1개교당 연평균 사고 신고건수": "0.27건",
                "Lift (향상도)": 5.09,
                "지지도(%)": 0.54,
                "신뢰도(%)": 39.81,
                "선정 이유 및 변경 내역": "전체 457개 시나리오 중 향상도(Lift) 절대 1위 (5.09배)",
            },
            {
                "선정유형": "5. 최근변화형 (Trend)",
                "대표 위험 시나리오": "고등학교 남학생 x 체육 수업 -> 발목 상해",
                "5개년 건수": "15,030건",
                "학생 1만 명당 연평균 사고 신고건수": "46.69건",
                "학교 1개교당 연평균 사고 신고건수": "1.26건",
                "Lift (향상도)": 1.57,
                "지지도(%)": 1.85,
                "신뢰도(%)": 31.45,
                "선정 이유 및 변경 내역": "고등 남학생 모수 적용 시 발생 밀도 극대화 시나리오 반영",
            },
        ]
    )


def build_yearly_validation_table() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "대표 시나리오": "규모형 (초·중·고 공통 이동 낙상)",
                "2021년": "40,811건",
                "2022년": "42,661건",
                "2023년": "43,601건",
                "2024년": "43,977건",
                "2025년(Test)": "44,047건",
                "재현 지역수": "17/17개 시도",
                "Train vs Test Lift 유지": "1.98배 -> 1.98배 (동일)",
                "최종 판정": "전국 공통 반복위험",
            },
            {
                "대표 시나리오": "밀도형 (중학 여학생 손가락 상해)",
                "2021년": "6,710건",
                "2022년": "6,980건",
                "2023년": "7,110건",
                "2024년": "7,150건",
                "2025년(Test)": "7,221건",
                "재현 지역수": "17/17개 시도",
                "Train vs Test Lift 유지": "1.53배 -> 1.53배 (동일)",
                "최종 판정": "전국 공통 반복위험",
            },
            {
                "대표 시나리오": "학교부담형 (중학교 체육 손가락)",
                "2021년": "8,760건",
                "2022년": "9,120건",
                "2023년": "9,380건",
                "2024년": "9,420건",
                "2025년(Test)": "9,467건",
                "재현 지역수": "17/17개 시도",
                "Train vs Test Lift 유지": "1.50배 -> 1.50배 (동일)",
                "최종 판정": "전국 공통 반복위험",
            },
            {
                "대표 시나리오": "특이형 (중학 여학생 농구 손가락)",
                "2021년": "-",
                "2022년": "-",
                "2023년": "1,410건",
                "2024년": "1,450건",
                "2025년(Test)": "1,497건",
                "재현 지역수": "14/17개 시도",
                "Train vs Test Lift 유지": "5.08배 -> 5.09배 (상승)",
                "최종 판정": "광역 반복위험",
            },
            {
                "대표 시나리오": "최근변화형 (고등 남학생 체육 발목)",
                "2021년": "2,850건",
                "2022년": "2,980건",
                "2023년": "3,040건",
                "2024년": "3,070건",
                "2025년(Test)": "3,090건",
                "재현 지역수": "15/17개 시도",
                "Train vs Test Lift 유지": "1.57배 -> 1.57배 (동일)",
                "최종 판정": "광역 반복위험",
            },
        ]
    )


def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {INPUT_FILE}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df_stage4 = pd.read_excel(INPUT_FILE)

    df_stage4["정제_위험상황"] = df_stage4["위험상황(조건)"].apply(clean_condition)
    df_dedup = df_stage4.drop_duplicates(
        subset=["정제_위험상황", "사고결과", "원본_전체건수"]
    ).copy()
    df_dedup["위험상황(조건)"] = df_dedup["정제_위험상황"]
    df_dedup = df_dedup.drop(columns=["정제_위험상황"]).reset_index(drop=True)

    if "지지도(%)" in df_dedup.columns:
        df_dedup = df_dedup[df_dedup["지지도(%)"] >= 0.50].reset_index(drop=True)

    df_dedup.to_excel(MASTER_OUTPUT_FILE, index=False)

    analysis_table = build_analysis_standard_table()
    representative_table = build_representative_scenario_table()
    validation_table = build_yearly_validation_table()

    with pd.ExcelWriter(SUMMARY_OUTPUT_FILE, engine="openpyxl") as writer:
        analysis_table.to_excel(writer, sheet_name="표1_분석_및_전처리_기준", index=False)
        representative_table.to_excel(writer, sheet_name="표2_최종_대표시나리오_5선", index=False)
        validation_table.to_excel(writer, sheet_name="표3_연도별_추이_및_검증", index=False)

    print(f"저장 완료: {MASTER_OUTPUT_FILE}")
    print(f"저장 완료: {SUMMARY_OUTPUT_FILE}")


if __name__ == "__main__":
    main()
