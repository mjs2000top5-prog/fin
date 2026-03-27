import streamlit as st
import pandas as pd
import io

# 1. 페이지 설정
st.set_page_config(page_title="통합 정산 보고서", layout="wide")
st.title("📑 데이터 교정 및 고정비 시트 연동 보고서")

# --- 구글 시트 정보 ---
SHEET_ID = "1VCJZqLL4EoaPYTClk9Ovc6Bd106CGlHT-CL1G_RRUSk"
# 수입/지출 내역 (메인 데이터)
EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
# 고정비 데이터 (gid=1372939881) - CSV 형식으로 호출
FIXED_COST_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=1372939881"

# [양식 구조 정의]
TEMPLATE_STRUCTURE = [
    ("위멤버스", "매출", "수강료"), ("위멤버스", "매출", "신규가입 포인트"), ("위멤버스", "매출", "이용료 수납"),
    ("위멤버스", "매출", "비즈 포인트 매출"), ("위멤버스", "매출", "모두싸인 매출"),
    ("위멤버스", "매입", "광고비"), ("위멤버스", "매입", "브랜드 광고비"), ("위멤버스", "매입", "판촉물"),
    ("위멤버스", "매입", "강사료 배분"), ("위멤버스", "매입", "통신비"), ("위멤버스", "매입", "솔루션 비용"), 
    ("위멤버스", "매입", "비즈포인트 이용료"), ("위멤버스", "매입", "신용카드 수수료"), ("위멤버스", "매입", "이벤터스 수수료"), 
    ("위멤버스", "매입", "CMS 수수료"), ("위멤버스", "매입", "KCP 수수료"), ("위멤버스", "매입", "EM3"), 
    ("위멤버스", "매입", "카카오 비즈메세지"), ("위멤버스", "매입", "모두싸인 배분"),
    ("세모R", "매출", "이용료 수납"), ("세모R", "매입", "CMS 수수료"), ("세모R", "매입", "신용카드 수수료"),
    ("세모장부", "매출", "경남 오락"), ("세모장부", "매출", "전북 오락"), ("세모장부", "매출", "부산 오락"),
    ("세모장부", "매출", "우리 원스퀘어"), ("세모장부", "매출", "NH소상공인파트너"), ("세모장부", "매출", "세모장부"),
    ("세모장부", "매입", "은행 배분_전북"), ("세모장부", "매입", "은행 배분_부산"), ("세모장부", "매입", "은행 배분_NH"),
    ("세모장부", "매입", "쿠콘 배분"), ("세모장부", "매입", "로움 배분"),
    ("링크패스", "매출", "링크패스 매출"), ("경리나라", "매입", "포인트")
]

MONTH_ORDER = [f"{i}월" for i in range(1, 13)]
PARENT_ITEMS = ["위멤버스", "세모R", "세모장부", "링크패스", "경리나라"]
DETAIL_OPTIONS = sorted(list(set([item[2] for item in TEMPLATE_STRUCTURE])) + ["미분류"])

# --- 데이터 로드 함수 ---
@st.cache_data
def load_main_data(url):
    try:
        df = pd.read_excel(url)
        df.columns = [str(c).strip() for c in df.columns]
        if '예정(발행)일' in df.columns:
            df['예정(발행)일'] = pd.to_datetime(df['예정(발행)일'], errors='coerce')
            df = df.dropna(subset=['예정(발행)일'])
            df['월'] = df['예정(발행)일'].dt.month.apply(lambda x: f"{x}월")
        if '금액' in df.columns:
            df['금액'] = pd.to_numeric(df['금액'].astype(str).str.replace(',', '').replace('-', '0'), errors='coerce').fillna(0)

        def classify_by_strict_standard(row):
            content = str(row.get('세부매출내용', '')).replace(" ", "")
            client = str(row.get('거래처명', '')).replace(" ", "")
            
            # 기존 분류 로직 유지
            if '쿠콘제휴배분' in content: return pd.Series(["미분류", "미분류", "미분류"])
            if '쿠콘' in content and '배분' in content: return pd.Series(["세모장부", "매입", "쿠콘 배분"])
            if '로움' in content and '배분' in content: return pd.Series(["세모장부", "매입", "로움 배분"])
            if '세모R매출' in content: return pd.Series(["세모R", "매출", "이용료 수납"])
            if '설명회다과구매' in content: return pd.Series(["위멤버스", "매입", "광고비"])
            if '해피디자인' in client: return pd.Series(["위멤버스", "매입", "판촉물"])
            if '구글광고비용' in content: return pd.Series(["위멤버스", "매입", "광고비"])
            if '비즈플레이' in client and '비즈포인트' in content: return pd.Series(["위멤버스", "매입", "비즈포인트 이용료"])
            if '쿠콘' in client or '비즈메세지' in content or '비즈메시지' in content: return pd.Series(["위멤버스", "매입", "카카오 비즈메세지"])
            if 'NH소상공인' in content and '배분' in content: return pd.Series(["세모장부", "매입", "은행 배분_NH"])
            if '전북' in content and ('배분' in content or '수수료' in content): return pd.Series(["세모장부", "매입", "은행 배분_전북"])
            if '부산' in content and ('배분' in content or '수수료' in content): return pd.Series(["세모장부", "매입", "은행 배분_부산"])
            if '위멤버스수강료' in content: return pd.Series(["위멤버스", "매출", "수강료"])
            elif '위멤버스포인트' in content and '매출' in client: return pd.Series(["위멤버스", "매출", "신규가입 포인트"])
            elif '위멤버스매출' in content: return pd.Series(["위멤버스", "매출", "이용료 수납"])
            elif 'Biz포인트' in content or '비즈플레이' in client: return pd.Series(["위멤버스", "매출", "비즈 포인트 매출"])
            elif '경남오락' in content: return pd.Series(["세모장부", "매출", "경남 오락"])
            elif '전북오락' in content: return pd.Series(["세모장부", "매출", "전북 오락"])
            elif '부산오락' in content: return pd.Series(["세모장부", "매출", "부산 오락"])
            elif '우리원스퀘어' in content or '우리은행' in client: return pd.Series(["세모장부", "매출", "우리 원스퀘어"])
            elif 'NH소상공인' in content: return pd.Series(["세모장부", "매출", "NH소상공인파트너"])
            elif '세모매출' in content: return pd.Series(["세모장부", "매출", "세모장부"])
            elif '모두싸인매출' in content: return pd.Series(["위멤버스", "매출", "모두싸인 매출"])
            elif '링크패스' in content: return pd.Series(["링크패스", "매출", "링크패스 매출"])
            if '위멤버스포인트' in content and '매입' in client: return pd.Series(["경리나라", "매입", "포인트"])
            if '브랜드' in content or any(k in content for k in ['구글위멤버스', '카카오광고', '네이버광고']): return pd.Series(["위멤버스", "매입", "브랜드 광고비"])
            if any(k in content for k in ['구글애즈', '프로모션', '사은품']): return pd.Series(["위멤버스", "매입", "광고비"])
            if '강사료' in content or '강의료' in content: return pd.Series(["위멤버스", "매입", "강사료 배분"])
            if '통신비' in content: return pd.Series(["위멤버스", "매입", "통신비"])
            if '솔루션' in content or '스케줄러' in content: return pd.Series(["위멤버스", "매입", "솔루션 비용"])
            if '이벤터스' in content: return pd.Series(["위멤버스", "매입", "이벤터스 수수료"])
            if 'KCP' in content: return pd.Series(["위멤버스", "매입", "KCP 수수료"])
            if 'EM3' in content or '업무도급' in content: return pd.Series(["위멤버스", "매입", "EM3"])
            if '모두싸인배분' in content: return pd.Series(["위멤버스", "매입", "모두싸인 배분"])
            if 'CMS' in content:
                return pd.Series(["세모R", "매입", "CMS 수수료"]) if '세모' in content else pd.Series(["위멤버스", "매입", "CMS 수수료"])
            if '신용카드' in content:
                return pd.Series(["세모R", "매입", "신용카드 수수료"]) if '세모' in content else pd.Series(["위멤버스", "매입", "신용카드 수수료"])
            return pd.Series(["미분류", "미분류", "미분류"])

        df[['상위항목', '구분', '상세항목']] = df.apply(classify_by_strict_standard, axis=1)
        return df[['월', '구분', '상위항목', '상세항목', '금액', '세부매출내용', '거래처명']]
    except Exception as e:
        st.error(f"메인 데이터 로드 실패: {e}")
        return None

@st.cache_data
def load_fixed_costs(url):
    try:
        # 고정비 시트(gid=1372939881) 로드
        df = pd.read_csv(url)
        # 컬럼명 앞뒤 공백 제거
        df.columns = [str(c).strip() for c in df.columns]
        
        # '월' 컬럼이 없는 경우 첫 번째 컬럼을 월로 지정
        if '월' not in df.columns:
            df.rename(columns={df.columns[0]: '월'}, inplace=True)
            
        # 데이터 정제
        df['월'] = df['월'].astype(str).str.strip()
        df['월'] = df['월'].apply(lambda x: x if '월' in x else f"{x}월")
        
        # 금액 데이터 변환
        for col in df.columns:
            if col != '월':
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        return df
    except Exception as e:
        st.warning(f"고정비 시트 로드 실패: {e}")
        return None

# 데이터 초기화
if "master_df" not in st.session_state:
    st.session_state.master_df = load_main_data(EXPORT_URL)
if "fixed_costs_df" not in st.session_state:
    st.session_state.fixed_costs_df = load_fixed_costs(FIXED_COST_URL)

if st.sidebar.button("🔄 데이터 및 고정비 새로고침"):
    st.cache_data.clear()
    st.session_state.master_df = load_main_data(EXPORT_URL)
    st.session_state.fixed_costs_df = load_fixed_costs(FIXED_COST_URL)
    st.rerun()

tab1, tab2, tab3 = st.tabs(["💰 정산 보고서", "📊 월별 손익 요약", "📄 상세 내역 확인 및 교정"])

# --- Tab 3: 상세 내역 확인 및 교정 ---
with tab3:
    st.subheader("🛠️ 상세 내역 확인 및 교정")
    edited_df = st.data_editor(
        st.session_state.master_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "월": st.column_config.SelectboxColumn(options=MONTH_ORDER, required=True),
            "구분": st.column_config.SelectboxColumn(options=["매출", "매입", "미분류"], required=True),
            "상위항목": st.column_config.SelectboxColumn(options=PARENT_ITEMS + ["미분류"], required=True),
            "상세항목": st.column_config.SelectboxColumn(options=DETAIL_OPTIONS, required=True),
            "금액": st.column_config.NumberColumn(format="%d")
        },
        key="main_editor"
    )
    st.session_state.master_df = edited_df

# --- Tab 1: 정산 보고서 ---
with tab1:
    st.subheader("📋 실시간 반영 정산 보고서")
    df_curr = st.session_state.master_df
    active_months = [m for m in MONTH_ORDER if m in df_curr['월'].unique()]
    base_df = pd.DataFrame(TEMPLATE_STRUCTURE, columns=['상위항목', '구분', '상세항목'])
    actual_summary = df_curr[df_curr['상위항목'] != "미분류"].pivot_table(
        index=['상위항목', '구분', '상세항목'], columns='월', values='금액', aggfunc='sum', fill_value=0
    ).reset_index()
    final_report = pd.merge(base_df, actual_summary, on=['상위항목', '구분', '상세항목'], how='left').fillna(0)
    final_report = final_report[['상위항목', '구분', '상세항목'] + active_months]
    final_report['합계'] = final_report[active_months].sum(axis=1)
    st.dataframe(final_report.set_index(['상위항목', '구분', '상세항목']).style.format("{:,.0f}"), use_container_width=True)

# --- Tab 2: 월별 손익 요약 (연동 완료) ---
with tab2:
    st.subheader("⚖️ 월별 종합 손익 요약 (시트 연동)")
    
    f_df = st.session_state.fixed_costs_df
    d_costs, l_costs = {}, {}

    st.sidebar.subheader("🛠️ 고정비 연동 확인")
    for m in active_months:
        default_d, default_l = 0, 0
        
        # 1. 고정비 시트에서 해당 월의 데이터를 찾음
        if f_df is not None and '월' in f_df.columns:
            target_row = f_df[f_df['월'] == m]
            if not target_row.empty:
                # 시트의 열 이름이 '운영비', '인건비'인 경우 가져옴
                if '운영비' in target_row.columns:
                    default_d = target_row['운영비'].values[0]
                if '인건비' in target_row.columns:
                    default_l = target_row['인건비'].values[0]
        
        # 2. 사이드바에 표시 (시트 데이터가 기본값으로 들어감)
        d_costs[m] = st.sidebar.number_input(f"{m} 운영비", value=int(default_d), key=f"d_{m}")
        l_costs[m] = st.sidebar.number_input(f"{m} 인건비", value=int(default_l), key=f"l_{m}")

    summary_list = []
    for m in active_months:
        s = df_curr[(df_curr['월']==m) & (df_curr['구분']=='매출')]['금액'].sum()
        p = df_curr[(df_curr['월']==m) & (df_curr['구분']=='매입')]['금액'].sum()
        
        # 시트 혹은 사이드바에서 가져온 고정비를 차감하여 손익 계산
        summary_list.append({
            "월": m, 
            "총 매출": s, 
            "총 매입(변동비)": p, 
            "운영비(고정)": d_costs[m], 
            "인건비(고정)": l_costs[m], 
            "최종 손익": s - p - d_costs[m] - l_costs[m]
        })
    
    st.table(pd.DataFrame(summary_list).set_index("월").style.format("{:,.0f}"))

    # --- 상위 항목별 상세 손익 ---
    st.markdown("---")
    st.subheader("🏢 상위 항목별 월별 상세 손익")
    
    parent_detail_data = []
    for parent in PARENT_ITEMS:
        for m in active_months:
            rev = df_curr[(df_curr['월']==m) & (df_curr['상위항목']==parent) & (df_curr['구분']=='매출')]['금액'].sum()
            exp = df_curr[(df_curr['월']==m) & (df_curr['상위항목']==parent) & (df_curr['구분']=='매입')]['금액'].sum()
            parent_detail_data.append({"상위항목": parent, "월": m, "지표": "1.매출금액", "금액": rev})
            parent_detail_data.append({"상위항목": parent, "월": m, "지표": "2.매입금액", "금액": exp})
            parent_detail_data.append({"상위항목": parent, "월": m, "지표": "3.손익", "금액": rev - exp})
    
    if parent_detail_data:
        pd_df = pd.DataFrame(parent_detail_data)
        pd_pivot = pd_df.pivot_table(index=["상위항목", "지표"], columns="월", values="금액", aggfunc="sum", fill_value=0)
        pd_pivot = pd_pivot[[m for m in active_months if m in pd_pivot.columns]]
        
        def highlight_profit(s):
            return ['background-color: #f0f2f6; font-weight: bold' if "3.손익" in s.name[1] else '' for _ in s]

        st.dataframe(pd_pivot.style.format("{:,.0f}").apply(highlight_profit, axis=1), use_container_width=True)
    else:
        st.info("데이터가 없습니다.")

st.sidebar.markdown("---")
st.sidebar.download_button("📊 최종 엑셀 저장", data=io.BytesIO().getvalue(), file_name="최종_정산_보고서.xlsx")