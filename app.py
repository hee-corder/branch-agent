import math
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="선납·유효 진척 현황", layout="wide")

st.title("선납·유효 평가 자동 분석기")
st.write("선납 파일과 유효 파일을 각각 업로드하면 지점 선납과 본부/지역단/지점 유효 결과를 자동으로 계산합니다.")


# ----------------------------
# 공통 함수
# ----------------------------
def clean_numeric(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce"
    )


def to_excel_single(df, sheet_name="결과"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output


def to_excel_multi(sheets_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, index=False, sheet_name=str(sheet_name)[:31])
    output.seek(0)
    return output


# ----------------------------
# 선납 점수 함수
# ----------------------------
def get_sunnap_score(achievement):
    if pd.isna(achievement):
        return 0.0
    elif achievement >= 40:
        return 4.0
    elif achievement >= 37:
        return 3.7
    elif achievement >= 34:
        return 3.4
    elif achievement >= 31:
        return 3.1
    elif achievement >= 28:
        return 2.8
    elif achievement >= 25:
        return 2.5
    elif achievement >= 22:
        return 2.2
    elif achievement >= 19:
        return 1.9
    elif achievement >= 16:
        return 1.6
    elif achievement >= 13:
        return 1.3
    else:
        return 0.0


def get_sunnap_band(achievement):
    if pd.isna(achievement):
        return "값없음"
    elif achievement >= 40:
        return "40% 이상"
    elif achievement >= 37:
        return "37% 이상"
    elif achievement >= 34:
        return "34% 이상"
    elif achievement >= 31:
        return "31% 이상"
    elif achievement >= 28:
        return "28% 이상"
    elif achievement >= 25:
        return "25% 이상"
    elif achievement >= 22:
        return "22% 이상"
    elif achievement >= 19:
        return "19% 이상"
    elif achievement >= 16:
        return "16% 이상"
    elif achievement >= 13:
        return "13% 이상"
    else:
        return "13% 미만"


# ----------------------------
# 유효 점수 함수
# ----------------------------
def get_valid_base_score(achievement):
    if pd.isna(achievement):
        return 0.0
    elif achievement >= 105:
        return 2.0
    elif achievement >= 104:
        return 1.8
    elif achievement >= 103:
        return 1.6
    elif achievement >= 102:
        return 1.4
    elif achievement >= 101:
        return 1.2
    elif achievement >= 100:
        return 1.0
    elif achievement >= 99:
        return 0.8
    elif achievement >= 98:
        return 0.6
    elif achievement >= 97:
        return 0.4
    elif achievement >= 96:
        return 0.2
    else:
        return 0.0


def get_valid_band(achievement):
    if pd.isna(achievement):
        return "값없음"
    elif achievement >= 105:
        return "105% 이상"
    elif achievement >= 104:
        return "104%"
    elif achievement >= 103:
        return "103%"
    elif achievement >= 102:
        return "102%"
    elif achievement >= 101:
        return "101%"
    elif achievement >= 100:
        return "100%"
    elif achievement >= 99:
        return "99%"
    elif achievement >= 98:
        return "98%"
    elif achievement >= 97:
        return "97%"
    elif achievement >= 96:
        return "96%"
    else:
        return "96% 미만"


def get_valid_bonus(achievement):
    if pd.isna(achievement) or achievement <= 105:
        return 0.0
    bonus = math.floor((achievement - 105) / 5) * 0.2
    return round(min(max(bonus, 0.0), 2.0), 1)


def get_scale_factor(actual, avg_actual):
    if pd.isna(actual) or pd.isna(avg_actual) or avg_actual == 0:
        return 1.0
    factor = 0.7 + 0.3 * (actual / avg_actual)
    return round(min(max(factor, 0.9), 1.1), 3)


# ----------------------------
# 파일 읽기 함수
# ----------------------------
def load_sunnap_file(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = [str(col).strip() for col in df.columns]

    required_cols = ["조직코드", "선납비율"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"선납 파일 필수 컬럼이 없습니다: {missing_cols}")

    if "조직명" not in df.columns:
        if "지점명" in df.columns:
            df["조직명"] = df["지점명"]
        else:
            df["조직명"] = df["조직코드"]

    df["조직코드"] = df["조직코드"].astype(str).str.strip()
    df["조직명"] = df["조직명"].astype(str).str.strip()
    df["선납비율"] = clean_numeric(df["선납비율"])

    df = df[~df["조직코드"].str.contains("합계|총계", na=False)].copy()
    df = df[df["조직코드"].str.startswith("E", na=False)].copy()

    return df


def load_valid_file(uploaded_file):
    raw = pd.read_excel(uploaded_file, header=[0, 1, 2])

    flat_cols = []
    for col in raw.columns:
        lv1, lv2, lv3 = [str(x).strip() for x in col]

        if lv1 in ["마감년월", "조직코드", "조직명"]:
            flat_cols.append(lv1)
        else:
            lv2 = lv2.replace("\n", "")
            if lv2 == "당월":
                flat_cols.append(f"{lv1}_{lv3}")
            else:
                flat_cols.append(f"{lv1}_{lv2}")

    raw.columns = flat_cols

    raw["조직코드"] = raw["조직코드"].astype(str).str.strip()
    raw["조직명"] = raw["조직명"].astype(str).str.strip()

    raw = raw[raw["조직코드"] != "합계"].copy()
    raw = raw[raw["조직코드"].str[0].isin(["C", "D", "E"])].copy()

    keep_cols = [
        "마감년월", "조직코드", "조직명",
        "유효100_목표", "유효100_진척", "유효100_달성율",
        "유효200_목표", "유효200_진척", "유효200_달성율",
    ]

    missing_cols = [col for col in keep_cols if col not in raw.columns]
    if missing_cols:
        raise ValueError(f"유효 파일 필수 컬럼이 없습니다: {missing_cols}")

    valid_df = raw[keep_cols].copy()

    numeric_cols = [
        "유효100_목표", "유효100_진척", "유효100_달성율",
        "유효200_목표", "유효200_진척", "유효200_달성율",
    ]
    for col in numeric_cols:
        valid_df[col] = clean_numeric(valid_df[col])

    valid_df["조직레벨"] = valid_df["조직코드"].str[0].map({
        "C": "본부",
        "D": "지역단",
        "E": "지점"
    })

    return valid_df


# ----------------------------
# 유효 결과 생성 함수
# ----------------------------
def build_valid_result(df, metric, level_name):
    work = df.copy()

    target_col = f"{metric}_목표"
    actual_col = f"{metric}_진척"

    work[target_col] = clean_numeric(work[target_col])
    work[actual_col] = clean_numeric(work[actual_col])

    work = work[(work[target_col].notna()) & (work[actual_col].notna())].copy()

    work["달성도(%)"] = np.where(
        work[target_col] > 0,
        (work[actual_col] / work[target_col]) * 100,
        np.nan
    )
    work["달성도(%)"] = work["달성도(%)"].round(1)

    work["평가척도구간"] = work["달성도(%)"].apply(get_valid_band)
    work["기본득점"] = work["달성도(%)"].apply(get_valid_base_score)
    work["가점"] = work["달성도(%)"].apply(get_valid_bonus)

    avg_actual = work[actual_col].mean()
    work["실적규모계수"] = work[actual_col].apply(lambda x: get_scale_factor(x, avg_actual))
    work["최종점수"] = (work["기본득점"] * work["실적규모계수"] + work["가점"]).round(2)
    work["목표대비부족인원"] = (work[target_col] - work[actual_col]).clip(lower=0).round(0).astype("Int64")
    work["달성여부"] = work["달성도(%)"].apply(
        lambda x: "달성" if pd.notna(x) and x >= 100 else "미달성"
    )
    work["구분"] = level_name

    display = work[
        [
            "조직코드", "조직명", "구분",
            target_col, actual_col,
            "달성도(%)", "평가척도구간", "기본득점",
            "실적규모계수", "가점", "최종점수",
            "목표대비부족인원", "달성여부"
        ]
    ].copy()

    display = display.rename(columns={
        target_col: "목표",
        actual_col: "진척"
    })

    display = display.sort_values(
        by=["최종점수", "달성도(%)", "진척"],
        ascending=[False, False, False]
    ).reset_index(drop=True)

    return display


# ----------------------------
# 입력 영역
# ----------------------------
c1, c2 = st.columns(2)
with c1:
    goal_rate = st.number_input("목표 선납률(%)", min_value=0.1, value=10.0, step=0.1)
with c2:
    st.caption("유효100 / 유효200은 업로드한 유효 파일의 목표·진척 값으로 자동 계산됩니다.")

sunnap_file = st.file_uploader("선납 엑셀 파일 업로드", type=["xlsx"], key="sunnap")
valid_file = st.file_uploader("유효 엑셀 파일 업로드", type=["xlsx"], key="valid")

tab1, tab2 = st.tabs(["지점 선납", "유효 성과"])


# ----------------------------
# 탭 1. 지점 선납
# ----------------------------
with tab1:
    if sunnap_file is None:
        st.info("선납 파일을 업로드하면 지점 선납 결과판이 표시됩니다.")
    else:
        try:
            sunnap_df = load_sunnap_file(sunnap_file)

            result_df = sunnap_df.copy()
            result_df["달성도(%)"] = ((result_df["선납비율"] / goal_rate) * 100).round(1)
            result_df["득점"] = result_df["달성도(%)"].apply(get_sunnap_score)
            result_df["평가척도구간"] = result_df["달성도(%)"].apply(get_sunnap_band)
            result_df["목표대비부족분(%p)"] = (goal_rate - result_df["선납비율"]).clip(lower=0).round(2)
            result_df["달성여부"] = result_df["선납비율"].apply(
                lambda x: "달성" if pd.notna(x) and x >= goal_rate else "미달성"
            )

            result_df = result_df.sort_values(
                by=["득점", "선납비율", "목표대비부족분(%p)"],
                ascending=[False, False, True]
            ).reset_index(drop=True)

            total_branches = len(result_df)
            achieved_count = int((result_df["달성여부"] == "달성").sum())
            avg_rate = round(result_df["선납비율"].mean(), 2) if total_branches > 0 else 0
            avg_score = round(result_df["득점"].mean(), 2) if total_branches > 0 else 0

            s1, s2, s3, s4 = st.columns(4)
            s1.metric("전체 지점 수", total_branches)
            s2.metric("평균 선납률", f"{avg_rate}%")
            s3.metric("평균 득점", f"{avg_score}점")
            s4.metric("목표 달성 지점", achieved_count)

            display_cols = [
                "조직코드", "조직명", "선납비율", "달성도(%)",
                "평가척도구간", "득점", "목표대비부족분(%p)", "달성여부"
            ]

            st.markdown("### 지점별 선납 상세")
            st.dataframe(
                result_df[display_cols],
                use_container_width=True,
                hide_index=True
            )

            left, right = st.columns(2)

            with left:
                st.markdown("### 상위 5지점")
                st.dataframe(
                    result_df[display_cols].head(5),
                    use_container_width=True,
                    hide_index=True
                )

            with right:
                st.markdown("### 관리 필요 5지점")
                need_df = result_df.sort_values(
                    by=["득점", "목표대비부족분(%p)", "선납비율"],
                    ascending=[True, False, True]
                )
                st.dataframe(
                    need_df[display_cols].head(5),
                    use_container_width=True,
                    hide_index=True
                )

            excel_data = to_excel_single(result_df[display_cols], sheet_name="지점선납")
            st.download_button(
                label="지점 선납 결과 엑셀 다운로드",
                data=excel_data,
                file_name="지점선납_분석결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"선납 파일 처리 중 오류가 발생했습니다: {e}")


# ----------------------------
# 탭 2. 유효 성과
# ----------------------------
with tab2:
    if valid_file is None:
        st.info("유효 파일을 업로드하면 본부/지역단/지점 유효 결과판이 표시됩니다.")
    else:
        try:
            valid_df = load_valid_file(valid_file)

            hq_df = valid_df[valid_df["조직레벨"] == "본부"].copy()
            region_df = valid_df[valid_df["조직레벨"] == "지역단"].copy()
            branch_df = valid_df[valid_df["조직레벨"] == "지점"].copy()

            hq_200 = build_valid_result(hq_df, "유효200", "본부")
            region_200 = build_valid_result(region_df, "유효200", "지역단")
            branch_100 = build_valid_result(branch_df, "유효100", "지점")
            branch_200 = build_valid_result(branch_df, "유효200", "지점")

            v1, v2, v3, v4 = st.columns(4)
            v1.metric("본부 유효200", f"{len(hq_200)}개")
            v2.metric("지역단 유효200", f"{len(region_200)}개")
            v3.metric("지점 유효100", f"{len(branch_100)}개")
            v4.metric("지점 유효200", f"{len(branch_200)}개")

            st.markdown("### 본부 유효200")
            st.dataframe(hq_200, use_container_width=True, hide_index=True)

            st.markdown("### 지역단 유효200")
            st.dataframe(region_200, use_container_width=True, hide_index=True)

            st.markdown("### 지점 유효100")
            st.dataframe(branch_100, use_container_width=True, hide_index=True)

            st.markdown("### 지점 유효200")
            st.dataframe(branch_200, use_container_width=True, hide_index=True)

            summary_df = pd.DataFrame({
                "구분": ["본부 유효200", "지역단 유효200", "지점 유효100", "지점 유효200"],
                "건수": [len(hq_200), len(region_200), len(branch_100), len(branch_200)],
                "평균달성도(%)": [
                    round(hq_200["달성도(%)"].mean(), 1) if len(hq_200) > 0 else 0,
                    round(region_200["달성도(%)"].mean(), 1) if len(region_200) > 0 else 0,
                    round(branch_100["달성도(%)"].mean(), 1) if len(branch_100) > 0 else 0,
                    round(branch_200["달성도(%)"].mean(), 1) if len(branch_200) > 0 else 0,
                ],
                "평균최종점수": [
                    round(hq_200["최종점수"].mean(), 2) if len(hq_200) > 0 else 0,
                    round(region_200["최종점수"].mean(), 2) if len(region_200) > 0 else 0,
                    round(branch_100["최종점수"].mean(), 2) if len(branch_100) > 0 else 0,
                    round(branch_200["최종점수"].mean(), 2) if len(branch_200) > 0 else 0,
                ],
                "달성건수": [
                    int((hq_200["달성여부"] == "달성").sum()) if len(hq_200) > 0 else 0,
                    int((region_200["달성여부"] == "달성").sum()) if len(region_200) > 0 else 0,
                    int((branch_100["달성여부"] == "달성").sum()) if len(branch_100) > 0 else 0,
                    int((branch_200["달성여부"] == "달성").sum()) if len(branch_200) > 0 else 0,
                ]
            })

            excel_data = to_excel_multi({
                "유효요약": summary_df,
                "본부유효200": hq_200,
                "지역단유효200": region_200,
                "지점유효100": branch_100,
                "지점유효200": branch_200,
            })
            st.download_button(
                label="유효 결과 엑셀 다운로드",
                data=excel_data,
                file_name="유효_분석결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"유효 파일 처리 중 오류가 발생했습니다: {e}")