"""연관규칙 분석으로 다차원 학교안전 위험 시나리오 후보를 도출한다."""

from pathlib import Path

import pandas as pd
from mlxtend.frequent_patterns import association_rules, fpgrowth
from mlxtend.preprocessing import TransactionEncoder


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = PROJECT_ROOT / "outputs"

INPUT_FILE = OUTPUT_DIR / "초중고_최종_분석용_데이터마트.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "2단계_다차원_위험_시나리오_결과.xlsx"

# 후보 규칙 생성 단계의 기준. 최종 선별 기준은 05번 스크립트에서 다시 적용한다.
MIN_SUPPORT = 0.005
MIN_LIFT_FOR_CANDIDATES = 1.2

CONTEXT_COLUMNS = [
    "학교급",
    "사고자성별",
    "파생_학년군",
    "사고시간",
    "사고장소",
    "사고당시활동",
    "파생_학기맥락",
]
OUTCOME_COLUMNS = ["사고부위", "사고형태"]
EXCLUDED_VALUES = {"미분류", "nan", "기타"}


def main() -> None:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {INPUT_FILE}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df = pd.read_excel(INPUT_FILE)

    transactions = []
    context_set = set()
    outcome_set = set()

    for _, row in df.iterrows():
        transaction = []

        for col in CONTEXT_COLUMNS:
            val = str(row[col]).strip() if pd.notna(row[col]) else ""
            if val and val not in EXCLUDED_VALUES:
                transaction.append(val)
                context_set.add(val)

        for col in OUTCOME_COLUMNS:
            val = str(row[col]).strip() if pd.notna(row[col]) else ""
            if val and val not in EXCLUDED_VALUES:
                transaction.append(val)
                outcome_set.add(val)

        if transaction:
            transactions.append(transaction)

    encoder = TransactionEncoder()
    encoded = encoder.fit(transactions).transform(transactions)
    df_encoded = pd.DataFrame(encoded, columns=encoder.columns_)

    frequent_itemsets = fpgrowth(
        df_encoded,
        min_support=MIN_SUPPORT,
        use_colnames=True,
    )
    rules = association_rules(
        frequent_itemsets,
        metric="lift",
        min_threshold=MIN_LIFT_FOR_CANDIDATES,
    )

    valid_rules = []
    for _, row in rules.iterrows():
        antecedents = set(row["antecedents"])
        consequents = set(row["consequents"])

        # 상황조건 -> 사고결과 형태의 규칙만 남긴다.
        if antecedents.issubset(context_set) and consequents.issubset(outcome_set):
            valid_rules.append(
                {
                    "위험상황(조건)": ", ".join(sorted(antecedents)),
                    "사고결과": ", ".join(sorted(consequents)),
                    "support": round(row["support"], 6),
                    "confidence": round(row["confidence"], 6),
                    "lift": round(row["lift"], 6),
                }
            )

    df_result = pd.DataFrame(valid_rules)
    if not df_result.empty:
        df_result = df_result.sort_values(by="lift", ascending=False).reset_index(drop=True)

    df_result.to_excel(OUTPUT_FILE, index=False)
    print(f"저장 완료: {OUTPUT_FILE} ({len(df_result):,}개 후보 규칙)")


if __name__ == "__main__":
    main()
