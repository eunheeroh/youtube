import re
from collections import Counter
import streamlit as st
import pandas as pd
import importlib
import utils

# Streamlit 핫 리로드 시 utils 모듈이 최신 상태로 유지되도록 reload
importlib.reload(utils)

from utils import (
    get_video_info,
    fetch_youtube_comments,
    convert_to_excel,
    convert_to_csv,
    convert_to_txt
)

# 텍스트 분석 및 시간 추이 함수 (모듈 캐시 불일치 방지를 위한 안전한 자체 정의)
DEFAULT_STOPWORDS = {
    "너무", "진짜", "정말", "그냥", "이거", "이건", "영상", "보고", "댓글", "좋아요",
    "구독", "있는", "하는", "그리고", "근데", "대한", "통해", "많이", "계속", "지금",
    "오늘", "생각", "사람", "시간", "다들", "항상", "감사합니다", "어떻게", "이런",
    "저런", "그렇게", "모두", "누가", "내가", "네가", "우리", "유튜브", "이게", "저게",
    "like", "video", "this", "that", "with", "have", "from", "they", "will", "would",
    "there", "their", "what", "about", "which", "when", "make", "can", "just", "more"
}

def extract_top_keywords(df: pd.DataFrame, top_n: int = 20, extra_stopwords: str = "") -> pd.DataFrame:
    if df is None or df.empty or "댓글내용" not in df.columns:
        return pd.DataFrame(columns=["키워드", "언급횟수"])

    stopwords = set(DEFAULT_STOPWORDS)
    if extra_stopwords:
        user_words = [w.strip() for w in extra_stopwords.split(",") if w.strip()]
        stopwords.update(user_words)

    counter = Counter()
    for text in df["댓글내용"].dropna():
        words = re.findall(r'[가-힣a-zA-Z]{2,}', str(text).lower())
        for word in words:
            if word not in stopwords and len(word) >= 2:
                counter[word] += 1

    top_items = counter.most_common(top_n)
    if not top_items:
        return pd.DataFrame(columns=["키워드", "언급횟수"])

    return pd.DataFrame(top_items, columns=["키워드", "언급횟수"])

def extract_time_trend(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or "댓글작성일시" not in df.columns:
        return pd.DataFrame()

    temp_series = df["댓글작성일시"].astype(str)
    dates = temp_series.str.extract(r'(\d{4}-\d{2}-\d{2})')[0].dropna()
    if dates.empty:
        return pd.DataFrame()

    trend = dates.value_counts().sort_index().reset_index()
    trend.columns = ["날짜", "댓글수"]
    return trend

# 페이지 설정
st.set_page_config(
    page_title="유튜브 댓글 수집기 (YouTube Comment Collector)",
    page_icon="💬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 커스텀 CSS 스타일 적용
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #FF0000;
        margin-bottom: 0.2rem;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .sub-title {
        font-size: 1rem;
        color: #555;
        margin-bottom: 1.5rem;
    }
    .video-card {
        background-color: #f8f9fa;
        border-radius: 12px;
        padding: 16px;
        margin-bottom: 20px;
        border: 1px solid #e2e8f0;
        display: flex;
        gap: 16px;
        align-items: center;
    }
    .stat-box {
        background: #ffffff;
        border: 1px solid #e9ecef;
        border-radius: 10px;
        padding: 14px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .stat-number {
        font-size: 1.6rem;
        font-weight: bold;
        color: #0f172a;
    }
    .stat-label {
        font-size: 0.85rem;
        color: #64748b;
    }
    .best-card {
        background: #f0fdf4;
        border-left: 4px solid #22c55e;
        padding: 12px 16px;
        border-radius: 6px;
        margin-bottom: 12px;
    }
</style>
""", unsafe_allow_html=True)

# 헤더
st.markdown('<div class="main-title">🎬 유튜브 댓글 수집기</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">YouTube 주소(URL)만 입력하면 API 키 없이도 댓글을 간편하게 수집하고 엑셀/CSV로 다운로드할 수 있습니다.</div>', unsafe_allow_html=True)

# 사이드바 설정
with st.sidebar:
    st.header("⚙️ 수집 옵션 설정")
    
    sort_option = st.radio(
        "댓글 정렬 기준",
        options=["인기순 (Top)", "최신순 (Newest)"],
        index=0,
        help="유튜브에서 댓글을 정렬하는 기준입니다."
    )
    sort_code = 0 if "인기순" in sort_option else 1

    unlimited = st.checkbox("댓글 전체 수집 (수량이 많으면 시간이 걸릴 수 있습니다)", value=False)
    
    if unlimited:
        max_comments = 0  # 0이면 제한 없이 끝까지 수집
        st.info("⚠️ 전체 수집 모드: 영상에 수만 개의 댓글이 있는 경우 수집에 수 분 이상 소요될 수 있습니다.")
    else:
        max_comments = st.slider(
            "수집할 최대 댓글 수",
            min_value=20,
            max_value=2000,
            value=200,
            step=20,
            help="수집할 댓글의 최대 개수를 설정합니다."
        )

    include_replies = st.checkbox("답글(대댓글) 포함하기", value=True, help="원댓글 외에 답글(대댓글)도 함께 수집할지 여부를 선택합니다.")

    st.markdown("---")
    st.markdown("### 💡 사용 가이드")
    st.markdown("""
    - **지원하는 링크 형태**:
      - 일반 영상: `https://www.youtube.com/watch?v=...`
      - 단축 링크: `https://youtu.be/...`
      - 유튜브 쇼츠: `https://www.youtube.com/shorts/...`
    - **다운로드 형식**:
      - `Excel (.xlsx)`: 바로 열기 좋은 서식
      - `CSV (.csv)`: 한글 깨짐 방지(UTF-8-BOM)
      - `TXT (.txt)`: AI(ChatGPT/Gemini) 요약 분석용
    """)

# 세션 상태 초기화
if "comments_df" not in st.session_state:
    st.session_state.comments_df = None
if "current_url" not in st.session_state:
    st.session_state.current_url = ""
if "video_info" not in st.session_state:
    st.session_state.video_info = None

# 유튜브 URL 입력
url_input = st.text_input(
    "📺 유튜브 영상 주소를 입력하세요:",
    placeholder="예: https://www.youtube.com/watch?v=dQw4w9WgXcQ 또는 https://youtu.be/..."
)

# 영상 정보 미리보기
if url_input and url_input.strip() != "":
    cleaned_url = url_input.strip()
    # URL이 바뀌면 영상 정보 재조회
    if st.session_state.current_url != cleaned_url:
        st.session_state.current_url = cleaned_url
        st.session_state.video_info = get_video_info(cleaned_url)
        st.session_state.comments_df = None  # 이전 수집결과 초기화

    info = st.session_state.video_info
    if info:
        col_thumb, col_detail = st.columns([1, 3])
        with col_thumb:
            if info.get("thumbnail_url"):
                st.image(info["thumbnail_url"], use_container_width=True)
        with col_detail:
            st.subheader(info.get("title", "제목 없음"))
            st.caption(f"채널명: **{info.get('author', '알 수 없음')}**")
            st.write(f"[👉 유튜브에서 직접 보기]({cleaned_url})")

# 수집 실행 버튼
col_btn, _ = st.columns([1, 3])
with col_btn:
    start_collect = st.button("🚀 댓글 수집 시작하기", type="primary", use_container_width=True)

if start_collect:
    if not url_input or not url_input.strip():
        st.error("⚠️ 유튜브 영상 주소를 먼저 입력해주세요.")
    else:
        target_url = url_input.strip()
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.info("댓글 수집 엔진을 준비하는 중입니다... 잠시만 기다려주세요.")
        
        def update_progress(count, total):
            if total > 0:
                pct = min(count / total, 1.0)
                progress_bar.progress(pct)
                status_text.info(f"댓글 수집 중... ({count} / {total}개 수집 완료)")
            else:
                status_text.info(f"댓글 수집 중... (현재 {count}개 수집 완료)")

        try:
            with st.spinner("댓글을 불러오는 중입니다..."):
                comments = fetch_youtube_comments(
                    url=target_url,
                    max_comments=max_comments,
                    sort_by=sort_code,
                    include_replies=include_replies,
                    progress_callback=update_progress
                )
            
            progress_bar.progress(1.0)
            
            if not comments:
                status_text.warning("댓글을 찾을 수 없거나, 댓글 작성이 비활성화된 영상입니다.")
                st.session_state.comments_df = None
            else:
                status_text.success(f"🎉 성공적으로 총 {len(comments):,}개의 댓글을 수집했습니다!")
                df = pd.DataFrame(comments)
                st.session_state.comments_df = df

        except Exception as ex:
            progress_bar.empty()
            status_text.error(f"댓글 수집 중 오류가 발생했습니다: {str(ex)}")

# 수집 결과 표시
if st.session_state.comments_df is not None and not st.session_state.comments_df.empty:
    df = st.session_state.comments_df

    st.markdown("---")
    st.subheader("📊 수집 결과 요약 및 다운로드")

    # 통계 메트릭 계산
    total_count = len(df)
    unique_authors = df["작성자"].nunique()
    reply_count = (df["구분"] == "답글(대댓글)").sum()
    heart_count = (df["크리에이터하트"] == "O").sum()
    max_likes = df["좋아요수(숫자)"].max()

    m1, m2, m3, m4, m5 = st.columns(5)
    with m1:
        st.markdown(f'<div class="stat-box"><div class="stat-number">{total_count:,}</div><div class="stat-label">총 수집 댓글 수</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="stat-box"><div class="stat-number">{unique_authors:,}</div><div class="stat-label">고유 작성자 수</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="stat-box"><div class="stat-number">{reply_count:,}</div><div class="stat-label">대댓글 수</div></div>', unsafe_allow_html=True)
    with m4:
        st.markdown(f'<div class="stat-box"><div class="stat-number">{heart_count:,}</div><div class="stat-label">하트 받은 댓글</div></div>', unsafe_allow_html=True)
    with m5:
        st.markdown(f'<div class="stat-box"><div class="stat-number">{max_likes:,}</div><div class="stat-label">최다 좋아요 수</div></div>', unsafe_allow_html=True)

    st.write("")

    # 다운로드 섹션
    col_d1, col_d2, col_d3 = st.columns(3)
    
    excel_data = convert_to_excel(df)
    csv_data = convert_to_csv(df)
    txt_data = convert_to_txt(df)

    # 비디오 제목으로 파일명 정규화
    clean_title = "youtube_comments"
    if st.session_state.video_info and st.session_state.video_info.get("title"):
        clean_title = "".join(c for c in st.session_state.video_info["title"] if c.isalnum() or c in (' ', '_', '-')).strip()
        clean_title = clean_title[:30] if clean_title else "youtube_comments"

    with col_d1:
        st.download_button(
            label="📥 엑셀 다운로드 (.xlsx)",
            data=excel_data,
            file_name=f"{clean_title}_댓글.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_d2:
        st.download_button(
            label="📄 CSV 다운로드 (.csv)",
            data=csv_data,
            file_name=f"{clean_title}_댓글.csv",
            mime="text/csv",
            use_container_width=True
        )
    with col_d3:
        st.download_button(
            label="📝 AI분석용 텍스트 (.txt)",
            data=txt_data,
            file_name=f"{clean_title}_댓글.txt",
            mime="text/plain",
            use_container_width=True
        )

    st.markdown("---")

    # 탭 구성: 데이터 목록 / 인기 댓글 TOP 5 / 텍스트 분석 & 차트 / 댓글 검색
    tab1, tab2, tab3, tab4 = st.tabs(["📋 전체 댓글 데이터", "🔥 베스트 댓글 TOP 5", "📊 텍스트 분석 & 차트", "🔍 검색 및 필터"])

    with tab1:
        # 데이터프레임 표시 (보기 좋은 기본 컬럼 위주)
        display_columns = ["번호", "구분", "작성자", "댓글내용", "좋아요수", "답글수", "댓글작성일시", "수집일시", "크리에이터하트"]
        st.dataframe(
            df[display_columns],
            use_container_width=True,
            height=450,
            column_config={
                "댓글내용": st.column_config.TextColumn("댓글 내용", width="large"),
                "작성자": st.column_config.TextColumn("작성자", width="medium"),
                "좋아요수": st.column_config.TextColumn("좋아요", width="small"),
                "댓글작성일시": st.column_config.TextColumn("댓글 작성일시", width="medium"),
                "수집일시": st.column_config.TextColumn("수집일시", width="medium"),
                "번호": st.column_config.NumberColumn("No.", width="small")
            }
        )

    with tab2:
        st.subheader("👍 좋아요가 가장 많은 베스트 댓글")
        # 좋아요수(숫자) 기준 내림차순 정렬
        top_df = df.sort_values(by="좋아요수(숫자)", ascending=False).head(5)
        for idx, row in top_df.iterrows():
            write_time_display = row.get('댓글작성일시', '')
            if row.get('작성시기') and row.get('작성시기') != write_time_display:
                write_time_display = f"{write_time_display} ({row.get('작성시기')})"
            st.markdown(f"""
            <div class="best-card">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                    <strong>👤 {row['작성자']}</strong>
                    <span style="color:#e11d48; font-weight:bold;">❤️ 좋아요 {row['좋아요수']}개</span>
                </div>
                <div style="font-size:0.95rem; white-space:pre-wrap; margin-bottom:6px;">{row['댓글내용']}</div>
                <div style="font-size:0.8rem; color:#6b7280;">📅 작성일시: {write_time_display} · 🕒 수집일시: {row.get('수집일시', '')} · 구분: {row['구분']}</div>
            </div>
            """, unsafe_allow_html=True)

    with tab3:
        st.subheader("📊 수집된 댓글 텍스트 분석 및 시각화")

        # 섹션 1: 주요 키워드 빈도 분석
        st.markdown("#### 🔤 최다 언급 키워드 TOP N")
        col_k1, col_k2 = st.columns([1, 2])
        with col_k1:
            top_n = st.slider("표시할 키워드 개수", min_value=5, max_value=30, value=15, step=5)
            extra_sw = st.text_input("제외할 단어(쉼표로 구분):", placeholder="예: 진짜, 너무, ㅋㅋ")
        
        kw_df = extract_top_keywords(df, top_n=top_n, extra_stopwords=extra_sw)

        if not kw_df.empty:
            col_chart, col_table = st.columns([2, 1])
            with col_chart:
                st.caption(f"언급 횟수가 높은 상위 {len(kw_df)}개 단어 막대 차트")
                chart_data = kw_df.set_index("키워드")
                st.bar_chart(chart_data, color="#3b82f6")
            with col_table:
                st.caption("키워드 순위 표")
                st.dataframe(kw_df, use_container_width=True, height=300)
        else:
            st.info("분석할 충분한 한글/영문 텍스트 단어가 없습니다.")

        st.markdown("---")

        # 섹션 2: 시간대별 작성 추이 & 댓글 길이 분포
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.markdown("#### 📈 일자별 댓글 등록 추이")
            trend_df = extract_time_trend(df)
            if not trend_df.empty:
                st.caption("날짜별 댓글 작성 수 변화")
                st.line_chart(trend_df.set_index("날짜"), color="#10b981")
            else:
                st.info("작성일시 데이터가 없거나 변환할 수 없습니다.")

        with col_c2:
            st.markdown("#### 📏 댓글 글자 수(길이) 분포")
            lengths = df["댓글내용"].astype(str).str.len()
            bins = [0, 10, 30, 50, 100, 10000]
            labels = ["10자 미만 (단문)", "10~30자", "30~50자", "50~100자", "100자 이상 (장문)"]
            len_binned = pd.cut(lengths, bins=bins, labels=labels, right=False).value_counts().reindex(labels)
            len_df = pd.DataFrame({"구간": len_binned.index, "댓글수": len_binned.values}).set_index("구간")
            st.caption("댓글 글자 수 구간별 빈도")
            st.bar_chart(len_df, color="#f59e0b")

        st.markdown("---")

        # 섹션 3: 좋아요 상위 댓글 순위 차트
        st.markdown("#### 🏆 좋아요를 많이 받은 댓글 TOP 10")
        top10_likes = df.sort_values(by="좋아요수(숫자)", ascending=False).head(10).copy()
        if not top10_likes.empty and top10_likes["좋아요수(숫자)"].max() > 0:
            # 표시용 짧은 라벨 생성
            top10_likes["라벨"] = top10_likes["작성자"] + ": " + top10_likes["댓글내용"].str.slice(0, 15) + "..."
            like_chart_data = top10_likes[["라벨", "좋아요수(숫자)"]].set_index("라벨")
            st.bar_chart(like_chart_data, color="#ef4444")
        else:
            st.info("좋아요가 1개 이상 기록된 댓글이 없습니다.")

    with tab4:
        st.subheader("🔎 댓글 검색 및 필터링")
        search_query = st.text_input("검색할 키워드 또는 작성자 이름을 입력하세요:", "")
        filter_type = st.selectbox("댓글 종류 필터", ["전체 보기", "일반댓글만 보기", "답글(대댓글)만 보기"])

        filtered_df = df.copy()
        if search_query.strip():
            q = search_query.strip().lower()
            filtered_df = filtered_df[
                filtered_df["댓글내용"].str.lower().str.contains(q, na=False) |
                filtered_df["작성자"].str.lower().str.contains(q, na=False)
            ]

        if filter_type == "일반댓글만 보기":
            filtered_df = filtered_df[filtered_df["구분"] == "일반댓글"]
        elif filter_type == "답글(대댓글)만 보기":
            filtered_df = filtered_df[filtered_df["구분"] == "답글(대댓글)"]

        st.caption(f"검색 결과: 총 {len(filtered_df):,}건")
        st.dataframe(filtered_df[["번호", "구분", "작성자", "댓글내용", "좋아요수", "댓글작성일시", "수집일시"]], use_container_width=True)

