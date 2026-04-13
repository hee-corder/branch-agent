import pandas as pd
from utils.scoring import piecewise_score, size_factor, bonus_by_step

STANDARD_BRANCHES = {
    "대전프라임", "대전한울", "군산", "청주제일", "상무", "전북", "익산", "광주",
    "충북", "충주", "광주중앙", "남악", "계룡",
    "익산중앙", "전남", "순천", "목포", "빛고을", "무등", "보령", "모악",
    "천안", "순천중앙", "논산", "서청주", "아산", "천안제일"
}

MBO_BRANCHES = {"충청", "서산"}


def classify_branch(branch_name: str) -> str:
    if branch_name in STANDARD_BRANCHES:
        return "표준"
    if branch_name in MBO_BRANCHES:
        return "영업소관리"
    return "TC"


def parse_effective_income(file):
    raw = pd.read_excel(file, header=None)

    # 3줄 헤더 결합
    col_names = []
    for col in range(raw.shape[1]):
        a = str(raw.iloc[0, col]).replace("\n", "").strip()
        b = str(raw.iloc[1, col]).replace("\n", "").strip()
        c = str(raw.iloc[2, col]).replace("\n", "").strip()

        if col <= 2:
            col_names.append(c if c != "nan" else a)
        else:
            col_names.append(f"{a}_{c}")

    df = raw.iloc[3:].copy()
    df.columns = col_names

    # 합계 제거
    df = df[df["조직명"] != "합계"]

    # 본부/지역단 제외, 지점만
    df["지점구분"] = df["조직명"].apply(classify_branch)
    df = df[~df["조직명"].isin(["충청호남본부"])]
    df = df[~df["조직명"].str.contains("지역단", na=False)]

    # 숫자 변환
    target_cols = [
        "유효100_목표", "유효100_진척", "유효100_달성율",
        "유효200_목표", "유효200_진척", "유효200_달성율"
    ]
    for c in target_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    xs = [96, 97, 98, 99, 100, 101, 102, 103, 104, 105]
    ys = [0.2, 0.4, 0.6, 0.8, 1.0, 1.2, 1.4, 1.6, 1.8, 2.0]

    results = []

    for item in ["유효100", "유효200"]:
        actual_col = f"{item}_진척"
        target_col = f"{item}_목표"
        rate_col = f"{item}_달성율"

        for group_name in ["표준", "TC"]:
            group_df = df[df["지점구분"] == group_name].copy()
            if group_df.empty:
                continue

            group_avg = group_df[actual_col].mean()

            for _, row in group_df.iterrows():
                rate = float(row[rate_col])
                actual = float(row[actual_col])
                target = float(row[target_col])

                base_score = 0.0 if rate < 96 else piecewise_score(rate, xs, ys)
                factor = size_factor(actual, group_avg, 0.9, 1.1)
                bonus = bonus_by_step(rate, start_rate=105, step_rate=5, step_point=0.2, max_bonus=2.0)

                final_score = round(base_score * factor + bonus, 2)

                results.append({
                    "평가항목": item,
                    "조직유형": "지점",
                    "조직명": row["조직명"],
                    "지점구분": row["지점구분"],
                    "목표": target,
                    "실적": actual,
                    "달성율": rate,
                    "부족분": max(target - actual, 0),
                    "달성현황": "달성" if rate >= 100 else "미달성",
                    "산출득점": final_score,
                    "최종득점": final_score,
                    "비고": ""
                })

    return pd.DataFrame(results)