import pandas as pd
from utils.scoring import piecewise_score, size_factor

STANDARD_BRANCHES = {
    "대전프라임", "대전한울", "군산", "청주제일", "상무", "전북", "익산", "광주",
    "충북", "충주", "광주중앙", "남악", "계룡",
    "익산중앙", "전남", "순천", "목포", "빛고을", "무등", "보령", "모악",
    "천안", "순천중앙", "논산", "서청주", "아산", "천안제일"
}

MBO_BRANCHES = {"충청", "서산"}


def normalize_branch_name(name: str) -> str:
    name = str(name).strip()
    name = name.replace("지점", "")
    name = name.replace("TC사업단", "")
    name = name.replace("TC지점", "")
    name = name.replace("사업단", "")
    return name.strip()


def classify_branch(branch_name: str) -> str:
    base_name = normalize_branch_name(branch_name)

    if base_name in STANDARD_BRANCHES:
        return "표준"
    if base_name in MBO_BRANCHES:
        return "영업소관리"
    return "TC"


def parse_prepayment(file):
    df = pd.read_excel(file)

    df["조직명"] = df["조직코드"].astype(str).str.split("\n").str[-1].str.strip()
    df["조직기준명"] = df["조직명"].apply(normalize_branch_name)

    # 선납 파일 원본 기준
    # 실적: 계약건수
    # 소계: 선납건수
    # 선납비율: 선납율(%)
    df["실적"] = pd.to_numeric(df["실적"], errors="coerce").fillna(0)
    df["소계"] = pd.to_numeric(df["소계"], errors="coerce").fillna(0)
    df["선납비율"] = pd.to_numeric(df["선납비율"], errors="coerce").fillna(0)

    # 본부 제거
    df = df[~df["조직명"].isin(["충청호남본부"])]
    df["지점구분"] = df["조직명"].apply(classify_branch)

    # 표준지점만 평가 대상
    target_df = df[df["지점구분"] == "표준"].copy()

    if target_df.empty:
        return pd.DataFrame([{
            "평가항목": "선납",
            "조직유형": "지점",
            "조직명": "평가대상 없음",
            "지점구분": "",
            "계약건수": "",
            "선납건수": "",
            "선납율": "",
            "구간점수": "",
            "규모조정계수": "",
            "개선도점수": "",
            "최종득점": "",
            "달성현황": "",
            "비고": "표준지점 매칭 실패 - 조직명 형식 확인 필요"
        }])

    # 평가척도 기준:
    # 달성도 40 / 37 / 34 / 31 / 28 / 25 / 22 / 19 / 16 / 13 / 13↓
    # 득점   4  / 3.7/ 3.4/ 3.1/ 2.8/ 2.5/ 2.2/ 1.9/ 1.6/ 1.3/ 0
    xs = [13, 16, 19, 22, 25, 28, 31, 34, 37, 40]
    ys = [1.3, 1.6, 1.9, 2.2, 2.5, 2.8, 3.1, 3.4, 3.7, 4.0]

    group_avg = target_df["실적"].mean()

    rows = []
    for _, row in target_df.iterrows():
        rate = float(row["선납비율"])
        actual_contract_count = float(row["실적"])
        prepay_count = float(row["소계"])

        # 13% 미만은 0점
        base_score = 0.0 if rate < 13 else piecewise_score(rate, xs, ys)

        # 규모조정율 적용
        factor = size_factor(actual_contract_count, group_avg, 0.9, 1.1)
        interval_score = round(base_score * factor, 2)

        # 개선도 점수는 기준값 파일이 아직 없으므로 계산 보류
        improvement_score = None

        # 현재는 구간점수만 최종득점으로 표시
        final_score = interval_score

        rows.append({
            "평가항목": "선납",
            "조직유형": "지점",
            "조직명": row["조직명"],
            "지점구분": row["지점구분"],
            "계약건수": actual_contract_count,
            "선납건수": prepay_count,
            "선납율": round(rate, 2),
            "구간점수": round(base_score, 2),
            "규모조정계수": round(factor, 3),
            "개선도점수": improvement_score,
            "최종득점": final_score,
            "달성현황": "최고구간" if rate >= 40 else "구간평가",
            "비고": "개선도 기준값(25.4Q 또는 25.1~5월 실적) 미업로드 → 현재는 선납율 구간점수만 반영"
        })

    return pd.DataFrame(rows)