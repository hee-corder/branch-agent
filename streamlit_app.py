import streamlit as st
import pandas as pd

from parsers.prepayment import parse_prepayment
from parsers.effective_income import parse_effective_income
from utils.analysis import generate_analysis

st.set_page_config(
    page_title="KB 평가항목 AI 대시보드",
    page_icon="📊",
    layout="wide"
)

st.title("KB 평가항목 AI 대시보드")
st.caption("평가척도 기준으로만 계산합니다. 척도 없는 항목은 계산하지 않습니다.")

st.markdown("## 1. 파일 업로드")

col1, col2 = st.columns(2)

with col1:
    prepayment_file = st.file_uploader("선납 파일 업로드", type=["xlsx", "xls"])

with col2:
    effective_file = st.file_uploader("유효소득 파일 업로드", type=["xlsx", "xls"])

st.markdown("---")
st.markdown("## 2. 결과")


def safe_mean(series: pd.Series) -> float:
    return round(pd.to_numeric(series, errors="coerce").fillna(0).mean(), 2)


if st.button("결과 산출하기"):
    pre_df = pd.DataFrame()
    eff_df = pd.DataFrame()

    if prepayment_file is not None:
        try:
            pre_df = parse_prepayment(prepayment_file)
            st.success("선납 파일 처리 완료")
        except Exception as e:
            st.error(f"선납 파일 처리 오류: {e}")

    if effective_file is not None:
        try:
            eff_df = parse_effective_income(effective_file)
            st.success("유효소득 파일 처리 완료")
        except Exception as e:
            st.error(f"유효소득 파일 처리 오류: {e}")

    if not pre_df.empty or not eff_df.empty:
        tab1, tab2 = st.tabs(["선납", "유효"])

        with tab1:
            st.markdown("### 선납 분석결과")

            if not pre_df.empty:
                st.dataframe(pre_df, use_container_width=True)

                st.markdown("### 3. 요약")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("건수", f"{len(pre_df)}개")
                c2.metric("미달성", f"{len(pre_df[pre_df['달성현황'] == '미달성'])}개")
                c3.metric("평균 득점", f"{safe_mean(pre_df['최종득점'])}점")
                c4.metric("평균 선납율", f"{safe_mean(pre_df['선납율'])}%")

                st.markdown("### 📌 분석 결과")
                pre_analysis_df = pre_df.copy()
                pre_analysis_df["달성율"] = pre_analysis_df["선납율"]
                st.write(generate_analysis(pre_analysis_df, "선납"))

            else:
                st.info("선납 결과가 없습니다.")

        with tab2:
            st.markdown("### 유효 분석결과")

            if not eff_df.empty:
                st.dataframe(eff_df, use_container_width=True)

                st.markdown("### 3. 요약")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("건수", f"{len(eff_df)}개")
                c2.metric("미달성", f"{len(eff_df[eff_df['달성현황'] == '미달성'])}개")
                c3.metric("평균 득점", f"{safe_mean(eff_df['최종득점'])}점")
                c4.metric("평균 달성율", f"{safe_mean(eff_df['달성율'])}%")

                # 유효100 / 유효200 따로 보기
                st.markdown("### 유효 세부 결과")
                eff_tab1, eff_tab2 = st.tabs(["유효100", "유효200"])

                with eff_tab1:
                    eff100_df = eff_df[eff_df["평가항목"] == "유효100"].copy()
                    if not eff100_df.empty:
                        st.dataframe(eff100_df, use_container_width=True)
                        st.markdown("#### 📌 유효100 분석 결과")
                        st.write(generate_analysis(eff100_df, "유효100"))
                    else:
                        st.info("유효100 결과가 없습니다.")

                with eff_tab2:
                    eff200_df = eff_df[eff_df["평가항목"] == "유효200"].copy()
                    if not eff200_df.empty:
                        st.dataframe(eff200_df, use_container_width=True)
                        st.markdown("#### 📌 유효200 분석 결과")
                        st.write(generate_analysis(eff200_df, "유효200"))
                    else:
                        st.info("유효200 결과가 없습니다.")

                st.markdown("### 📌 유효 전체 분석 결과")
                st.write(generate_analysis(eff_df, "유효"))

            else:
                st.info("유효 결과가 없습니다.")

    else:
        st.warning("업로드한 파일이 없거나, 처리 가능한 결과가 없습니다.")