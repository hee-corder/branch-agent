import pandas as pd


def generate_analysis(df: pd.DataFrame, item_name: str) -> str:
    if df.empty:
        return "분석할 데이터가 없습니다."

    work_df = df.copy()
    work_df["최종득점"] = pd.to_numeric(work_df["최종득점"], errors="coerce").fillna(0)
    work_df["달성율"] = pd.to_numeric(work_df["달성율"], errors="coerce").fillna(0)

    total = len(work_df)
    achieved = len(work_df[work_df["달성현황"] == "달성"])
    not_achieved = total - achieved

    avg_score = round(work_df["최종득점"].mean(), 2)
    avg_rate = round(work_df["달성율"].mean(), 2)

    # 하위 30%, 상위 30%
    df_sorted = work_df.sort_values(by="최종득점", ascending=True)
    bottom_count = max(int(total * 0.3), 1)
    top_count = max(int(total * 0.3), 1)

    bottom_df = df_sorted.head(bottom_count)
    top_df = df_sorted.tail(top_count)
    middle_df = df_sorted.iloc[bottom_count: total - top_count] if total > (bottom_count + top_count) else pd.DataFrame()

    bottom_names = ", ".join(bottom_df["조직명"].head(5).astype(str).tolist())
    top_names = ", ".join(top_df.sort_values(by="최종득점", ascending=False)["조직명"].head(5).astype(str).tolist())

    analysis = f"""
현재 **{item_name}** 항목은 전체 **{total}개 지점** 기준으로 집계되었습니다.

목표 달성 지점은 **{achieved}개**, 목표 미달성 지점은 **{not_achieved}개**입니다.  
평균 달성율은 **{avg_rate}%**, 평균 최종득점은 **{avg_score}점**입니다.

하위 구간은 **{len(bottom_df)}개 지점**이며, 대표적으로 **{bottom_names if bottom_names else '-'}** 등이 포함됩니다.  
이 구간의 지점들이 최소 중간권 수준까지 올라오면 본부 평균 득점과 평균 달성율이 함께 개선될 가능성이 큽니다.

중간 구간은 **{len(middle_df)}개 지점**입니다.  
이 구간은 약간의 추가 달성만으로도 상위권으로 이동할 수 있어, 단기적으로 평균을 끌어올리기 가장 효율적인 구간입니다.

상위 구간은 **{len(top_df)}개 지점**이며, 대표적으로 **{top_names if top_names else '-'}** 등이 현재 본부 성과를 견인하고 있습니다.  
이 구간은 유지 관리가 중요합니다.

종합하면, 현재 본부 성과는 **하위 지점 개선 여부**에 크게 좌우되는 구조입니다.  
하위 지점이 일정 수준 이상 득점을 확보하면 평균이 올라가고, 결과적으로 충청호남본부 전체 평가 경쟁력도 좋아질 수 있습니다.
""".strip()

    return analysis