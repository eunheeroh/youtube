import re
import io
import datetime
import requests
import pandas as pd
from youtube_comment_downloader import YoutubeCommentDownloader, SORT_BY_POPULAR, SORT_BY_RECENT

def parse_votes(vote_str):
    """
    좋아요 수 문자열('1.5만', '3.2천', '10K', '1.2M', '54' 등)을 숫자로 변환합니다.
    """
    if not vote_str:
        return 0
    vote_str = str(vote_str).strip().replace(',', '')
    if vote_str.isdigit():
        return int(vote_str)
    
    # 한국어 단위
    try:
        if '억' in vote_str:
            num = float(re.sub(r'[^\d.]', '', vote_str.split('억')[0]))
            return int(num * 100_000_000)
        if '만' in vote_str:
            num = float(re.sub(r'[^\d.]', '', vote_str.split('만')[0]))
            return int(num * 10_000)
        if '천' in vote_str:
            num = float(re.sub(r'[^\d.]', '', vote_str.split('천')[0]))
            return int(num * 1_000)
        
        # 영문 단위
        if 'M' in vote_str or 'm' in vote_str:
            num = float(re.sub(r'[^\d.]', '', vote_str.upper().split('M')[0]))
            return int(num * 1_000_000)
        if 'K' in vote_str or 'k' in vote_str:
            num = float(re.sub(r'[^\d.]', '', vote_str.upper().split('K')[0]))
            return int(num * 1_000)
        
        cleaned = re.sub(r'[^\d.]', '', vote_str)
        return int(float(cleaned)) if cleaned else 0
    except Exception:
        return 0

def extract_video_id(url: str):
    """
    유튜브 URL에서 비디오 ID를 추출합니다.
    """
    if not url:
        return None
    patterns = [
        r'(?:v=|\/)([0-9A-Za-z_-]{11}).*',
        r'(?:embed\/)([0-9A-Za-z_-]{11})',
        r'(?:shorts\/)([0-9A-Za-z_-]{11})',
        r'youtu\.be\/([0-9A-Za-z_-]{11})'
    ]
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    return None

def get_video_info(url: str):
    """
    유튜브 oEmbed API를 사용하여 영상의 제목, 작성자, 썸네일 정보를 가져옵니다.
    """
    try:
        oembed_url = f"https://www.youtube.com/oembed?url={url}&format=json"
        res = requests.get(oembed_url, timeout=5)
        if res.status_code == 200:
            data = res.json()
            return {
                "title": data.get("title", "제목 없음"),
                "author": data.get("author_name", "알 수 없음"),
                "author_url": data.get("author_url", ""),
                "thumbnail_url": data.get("thumbnail_url", ""),
                "html": data.get("html", "")
            }
    except Exception:
        pass
    
    # 비디오 ID만이라도 있으면 기본 썸네일 반환
    vid = extract_video_id(url)
    if vid:
        return {
            "title": f"영상 (ID: {vid})",
            "author": "알 수 없음",
            "author_url": "",
            "thumbnail_url": f"https://i.ytimg.com/vi/{vid}/hqdefault.jpg",
            "html": ""
        }
    return None

def fetch_youtube_comments(url: str, max_comments: int = 100, sort_by: int = SORT_BY_POPULAR, include_replies: bool = True, progress_callback=None):
    """
    youtube_comment_downloader를 사용하여 댓글을 수집하고 리스트로 반환합니다.
    """
    downloader = YoutubeCommentDownloader()
    comments = []
    
    # sort_by: 0 (SORT_BY_POPULAR), 1 (SORT_BY_RECENT)
    try:
        collected_at = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        comment_generator = downloader.get_comments_from_url(url, sort_by=sort_by)
        
        count = 0
        for item in comment_generator:
            # 대댓글 포함 여부 필터
            is_reply = item.get('reply', False)
            if not include_replies and is_reply:
                continue

            # 날짜 변환 (time_parsed: timestamp in seconds)
            time_parsed = item.get('time_parsed')
            datetime_str = ""
            if time_parsed:
                try:
                    dt = datetime.datetime.fromtimestamp(time_parsed)
                    datetime_str = dt.strftime('%Y-%m-%d %H:%M:%S')
                except Exception:
                    datetime_str = ""

            raw_votes = item.get('votes', 0)
            votes_num = parse_votes(raw_votes)

            comment_data = {
                "번호": count + 1,
                "구분": "답글(대댓글)" if is_reply else "일반댓글",
                "작성자": item.get('author', ''),
                "댓글내용": item.get('text', ''),
                "좋아요수": raw_votes if raw_votes else "0",
                "좋아요수(숫자)": votes_num,
                "답글수": item.get('replies', 0),
                "댓글작성일시": datetime_str if datetime_str else item.get('time', ''),
                "작성시기": item.get('time', ''),
                "수집일시": collected_at,
                "크리에이터하트": "O" if item.get('heart', False) else "X",
                "댓글ID": item.get('cid', ''),
                "작성자프로필": item.get('photo', ''),
                "채널ID": item.get('channel', '')
            }
            comments.append(comment_data)
            count += 1
            
            if progress_callback and (count % 10 == 0 or count == max_comments):
                progress_callback(count, max_comments)

            if max_comments > 0 and count >= max_comments:
                break

    except Exception as e:
        if not comments:
            raise e
        # 일부 수집 도중 에러가 나더라도 지금까지 수집된 건 반환
        pass

    return comments

def convert_to_excel(df: pd.DataFrame) -> bytes:
    """
    데이터프레임을 엑셀 바이너리로 변환합니다.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='유튜브_댓글')
    return output.getvalue()

def convert_to_csv(df: pd.DataFrame) -> bytes:
    """
    데이터프레임을 UTF-8-SIG CSV 바이트로 변환합니다. (Excel 한글 깨짐 방지)
    """
    return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

def convert_to_txt(df: pd.DataFrame) -> bytes:
    """
    댓글 본문만 텍스트 파일로 내보내기 위한 변환 (AI 요약 / 분석용)
    """
    lines = []
    for idx, row in df.iterrows():
        write_time = row.get('댓글작성일시', '') or row.get('작성시기', '')
        collect_time = row.get('수집일시', '')
        lines.append(f"[{row.get('작성자', '')}] (작성일: {write_time} | 수집일: {collect_time} | 좋아요: {row.get('좋아요수', '0')})\n{row.get('댓글내용', '')}\n{'-'*40}")
    content = "\n".join(lines)
    return content.encode('utf-8')

# 기본 불용어(Stopwords) 목록
DEFAULT_STOPWORDS = {
    "너무", "진짜", "정말", "그냥", "이거", "이건", "영상", "보고", "댓글", "좋아요",
    "구독", "있는", "하는", "그리고", "근데", "대한", "통해", "많이", "계속", "지금",
    "오늘", "생각", "사람", "시간", "다들", "항상", "감사합니다", "어떻게", "이런",
    "저런", "그렇게", "모두", "누가", "내가", "네가", "우리", "유튜브", "이게", "저게",
    "like", "video", "this", "that", "with", "have", "from", "they", "will", "would",
    "there", "their", "what", "about", "which", "when", "make", "can", "just", "more"
}

def extract_top_keywords(df: pd.DataFrame, top_n: int = 20, extra_stopwords: str = "") -> pd.DataFrame:
    """
    댓글 본문에서 최빈 출현 키워드를 추출하여 DataFrame으로 반환합니다.
    """
    from collections import Counter
    if df is None or df.empty or "댓글내용" not in df.columns:
        return pd.DataFrame(columns=["키워드", "언급횟수"])

    stopwords = set(DEFAULT_STOPWORDS)
    if extra_stopwords:
        user_words = [w.strip() for w in extra_stopwords.split(",") if w.strip()]
        stopwords.update(user_words)

    counter = Counter()
    for text in df["댓글내용"].dropna():
        # 한글, 영문 2글자 이상 단어 추출
        words = re.findall(r'[가-힣a-zA-Z]{2,}', str(text).lower())
        for word in words:
            if word not in stopwords and len(word) >= 2:
                counter[word] += 1

    top_items = counter.most_common(top_n)
    if not top_items:
        return pd.DataFrame(columns=["키워드", "언급횟수"])

    res_df = pd.DataFrame(top_items, columns=["키워드", "언급횟수"])
    return res_df

def extract_time_trend(df: pd.DataFrame) -> pd.DataFrame:
    """
    댓글작성일시를 기준으로 날짜별(일자별) 작성 추이를 집계합니다.
    """
    if df is None or df.empty or "댓글작성일시" not in df.columns:
        return pd.DataFrame()

    temp_series = df["댓글작성일시"].astype(str)
    # YYYY-MM-DD 형식 추출
    dates = temp_series.str.extract(r'(\d{4}-\d{2}-\d{2})')[0].dropna()
    if dates.empty:
        return pd.DataFrame()

    trend = dates.value_counts().sort_index().reset_index()
    trend.columns = ["날짜", "댓글수"]
    return trend


