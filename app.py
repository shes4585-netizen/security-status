import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import re
import os

st.set_page_config(layout="wide")

EXCEL_PATH = '비초소_태스크_리스트.xlsx'

# ────────────────────────────────────────────────
# 1. 데이터 로드 및 정제
# ────────────────────────────────────────────────
def clean_str(text):
    if pd.isna(text):
        return ""
    return re.sub(r'[\s\t\n\r]+', '', str(text))


@st.cache_data
def load_data(file_mtime):
    """
    file_mtime을 캐시 키에 포함시켜, 엑셀 파일이 수정될 때마다(mtime 변경)
    캐시가 자동으로 무효화되고 새 데이터를 다시 읽도록 함.
    (엑셀만 덮어쓰고 앱을 재시작하지 않아도 최신 데이터가 반영됨)
    """
    file_path = EXCEL_PATH
    배치 = pd.read_excel(file_path, sheet_name='배치')
    근무데이터 = pd.read_excel(file_path, sheet_name='근무데이터')
    대체근무자 = pd.read_excel(file_path, sheet_name='대체근무자')
    매핑 = pd.read_excel(file_path, sheet_name='태스크_코드 매핑')

    for df in [배치, 근무데이터, 대체근무자, 매핑]:
        df.columns = df.columns.astype(str).str.strip()
        for col in df.select_dtypes(include=['object', 'string']).columns:
            df[col] = df[col].apply(clean_str)

    근무데이터['날짜'] = pd.to_datetime(근무데이터['날짜']).dt.normalize()
    대체근무자['날짜'] = pd.to_datetime(대체근무자['날짜']).dt.normalize()

    for col in ['시작시각', '종료시각']:
        매핑[col] = pd.to_datetime(매핑[col], format='%H:%M:%S', errors='coerce').dt.time

    return 배치, 근무데이터, 대체근무자, 매핑


배치, 근무데이터, 대체근무자, 매핑 = load_data(os.path.getmtime(EXCEL_PATH))
st.caption(f"📄 데이터 파일 최종 수정: {datetime.fromtimestamp(os.path.getmtime(EXCEL_PATH)).strftime('%Y-%m-%d %H:%M:%S')}")

# 조회 속도를 위해 매핑/배치 정보는 dict로 미리 변환 (반복 조회 오버헤드 제거)
매핑_dict = {
    row['코드명']: (row['시작시각'], row['종료시각'], row['태스크명'])
    for _, row in 매핑.iterrows()
}
기본배치_dict = {n: p for n, p in zip(배치['경비원명'], 배치['소속초소'])}
전체_경비원명 = list(배치['경비원명'])

# 배치시트 '비고' 컬럼(예: "북초소대체담당")에서 지정 대체근무자 → 담당초소 매핑 추출
담당초소_dict = {}
for _, row in 배치.iterrows():
    비고 = row.get('비고', '')
    if 비고 and 비고.endswith('대체담당'):
        담당초소_dict[row['경비원명']] = 비고.replace('대체담당', '')

북초소_대체담당자 = next((n for n, p in 담당초소_dict.items() if p == '북초소'), None)

# 히트맵에 인원수 대신 약칭을 표시하기 위한 이름→약칭 매핑
약칭_dict = {
    '이춘도': '춘', '이근재': '근', '이학근': '학', '유복현': '복', '김민겸': '민',
    '신운기': '운', '정창화': '창', '여연동': '연', '이동철': '동', '박동원': '원',
}


# ────────────────────────────────────────────────
# 2. 공통 판정 함수 (탭1 / 탭2가 반드시 동일 로직을 쓰도록 통일)
#    → 지금까지 대화에서 나온 '탭1과 탭2 결과 불일치' 문제의 근본 원인이
#      두 탭이 서로 다른 코드로 계산했기 때문이므로, 이 함수 하나로 합칩니다.
# ────────────────────────────────────────────────
def in_range(t, s, e):
    """
    시간 포함 여부 판정.
    휴-3A(22:00~01:59)처럼 자정을 넘기는 구간까지 정확히 처리합니다.
    (기존 코드의 s <= t <= e 비교는 자정을 넘는 구간에서 항상 False가 되는
     숨은 오류가 있었어서 이번에 같이 바로잡았습니다.)
    """
    if s is None or e is None or pd.isna(s) or pd.isna(e):
        return False
    if s <= e:
        return s <= t <= e
    else:
        return t >= s or t <= e


def get_location(name, t, today_daeche):
    """해당 시각 기준 지리적 위치 확정 (기본 배치 → 대체근무 시 대체 초소로 이동)"""
    location = 기본배치_dict.get(name, "미배치")
    sub_rows = today_daeche[today_daeche['경비원명'] == name]
    for _, row in sub_rows.iterrows():
        info = 매핑_dict.get(row['코드명'])
        if info and in_range(t, info[0], info[1]):
            location = row['초소']
    return location


def get_status(name, t, today_data):
    """
    초소인원/비초소인원 판정.
    우선순위 1: 연차(연-1/2/3)는 해당 연차 유형의 실제 시간대에만 적용.
               연-1(전일)만 하루 종일(07:30~익일07:30) 적용되고,
               연-2(오전 07:30~19:29)/연-3(오후 19:30~07:29)는 그 시간대에만 연차로 판정.
               (연차자는 명단에서 지우지 않고, 해당 시간대엔 '연차'로 비초소 테이블에 표시)
    우선순위 2: 연차 시간대가 아니거나 연차가 없으면, 휴게/기타 태스크를 시간 범위로 비교.
    """
    s_df = today_data[today_data['경비원명'] == name]

    leave_rows = s_df[s_df['코드명'].str.startswith('연-')]
    if not leave_rows.empty:
        leave_code = leave_rows.iloc[0]['코드명']
        if leave_code == '연-1':
            return True, "연차"
        info = 매핑_dict.get(leave_code)
        if info and in_range(t, info[0], info[1]):
            return True, "연차"

    is_off, task_name = False, ""
    for _, row in s_df.iterrows():
        code = row['코드명']
        if code.startswith('연-'):
            continue  # 연차 코드는 위에서 이미 처리했으므로 제외
        info = 매핑_dict.get(code)
        if info and in_range(t, info[0], info[1]):
            is_off, task_name = True, info[2]
    return is_off, task_name


def north_backup_active(t, day_data, day_daeche):
    """
    이 시각에 박동원(북초소 대체담당)이 북초소를 채워야 하는지 여부.
    연차 발생 여부와 무관하게, 북초소에 '물리적으로 남아있는' 인원
    (연차자는 위치가 그대로 북초소이므로 포함되고, 다른 초소로 파견 나간 인원은 제외)
    중 초소인원(근무 중)이 0명이 되는 순간에만 True.
    """
    북초소_소속 = [n for n, p in 기본배치_dict.items() if p == '북초소']
    present = [n for n in 북초소_소속 if get_location(n, t, day_daeche) == '북초소']
    if not present:
        return True  # 전원 파견/부재 → 당연히 지원 필요
    return all(get_status(n, t, day_data)[0] for n in present)


def get_person_status(name, t, day_data, day_daeche):
    """
    위치 + 초소인원/비초소인원 상태를 한 번에 반환하는 최종 통합 함수.
    - 북초소 대체담당자(박동원)는 대체근무자 시트의 시간코드 대신,
      북초소 0명방지 동적 판정(north_backup_active)에 따라 위치가 결정됨.
    - 그 외 모든 인원(동/서/남초소 대체 포함)은 대체근무자 시트의 코드/시간 기반으로 결정.
    """
    if name == 북초소_대체담당자:
        if north_backup_active(t, day_data, day_daeche):
            return '북초소', False, ""  # 북초소 안에서 근무 중 → 북초소 초소인원
        own_off, own_task = get_status(name, t, day_data)
        return 기본배치_dict.get(name, "미배치"), own_off, own_task  # 본인 소속(남초소)에서 평소대로

    location = get_location(name, t, day_daeche)
    is_off, task_name = get_status(name, t, day_data)
    return location, is_off, task_name


def get_active_own_code(name, t, today_data):
    """
    이 시각에 근무데이터 시트에서 본인이 실제로 수행 중인 코드명(연차 포함)을 반환.
    get_status()와 동일한 우선순위(연차 → 휴게/기타 태스크)를 따르되,
    태스크명이 아니라 원본 코드명(예: 휴-1A, 연-2)을 그대로 반환.
    """
    s_df = today_data[today_data['경비원명'] == name]

    leave_rows = s_df[s_df['코드명'].str.startswith('연-')]
    if not leave_rows.empty:
        leave_code = leave_rows.iloc[0]['코드명']
        if leave_code == '연-1':
            return leave_code
        info = 매핑_dict.get(leave_code)
        if info and in_range(t, info[0], info[1]):
            return leave_code

    active_code = None
    for _, row in s_df.iterrows():
        code = row['코드명']
        if code.startswith('연-'):
            continue
        info = 매핑_dict.get(code)
        if info and in_range(t, info[0], info[1]):
            active_code = code
    return active_code


def get_active_sub_code(name, t, today_daeche):
    """
    이 시각에 대체근무자 시트에서 실제로 적용 중인 코드명을 반환.
    (북초소 대체담당자의 동적 판정처럼 대체근무자 시트를 아예 안 쓰는 경우엔 None)
    """
    active_code = None
    sub_rows = today_daeche[today_daeche['경비원명'] == name]
    for _, row in sub_rows.iterrows():
        info = 매핑_dict.get(row['코드명'])
        if info and in_range(t, info[0], info[1]):
            active_code = row['코드명']
    return active_code


def get_shift_day_data(base_date):
    """
    07:30~익일 07:30을 '하나의 근무일'로 볼 때 필요한 데이터를 모두 모읍니다.
    (야간 데이터가 익일 날짜로 기록되어 있으므로 base_date + 익일 데이터를 함께 조회)
    """
    base_ts = pd.to_datetime(base_date)
    next_ts = base_ts + pd.Timedelta(days=1)
    day_data = 근무데이터[(근무데이터['날짜'] == base_ts) | (근무데이터['날짜'] == next_ts)]
    day_daeche = 대체근무자[(대체근무자['날짜'] == base_ts) | (대체근무자['날짜'] == next_ts)]
    return day_data, day_daeche


st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nanum+Pen+Script&display=swap');

html, body, [class*="css"] {
    font-family: 'Nanum Pen Script', cursive !important;
    font-size: 70% !important;
}

/* 인쇄(프린트) 전용 스타일
   - 표의 컬럼 헤더(초소명)를 페이지마다 자동으로 반복 출력
   - 행이 페이지 경계에서 잘리지 않도록 함
   - 날짜선택/버튼 등 조회 컨트롤은 인쇄 시 숨겨서, 표가 1페이지부터 바로 시작되도록 함 */
@media print {
    thead { display: table-header-group; }
    tr, td, th { page-break-inside: avoid; }

    [data-testid="stSelectbox"],
    [data-testid="stButton"],
    [data-testid="stDateInput"],
    [data-testid="stTimeInput"],
    [data-testid="stCaptionContainer"] {
        display: none !important;
    }
}
</style>
""", unsafe_allow_html=True)

st.title("근무현황 조회")

tab1, tab2 = st.tabs(["🕒 초소별 실시간 인원", "📊 시간별 인원 히트맵"])

# ────────────────────────────────────────────────
# 탭 1: 실시간 조회
# ────────────────────────────────────────────────
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        target_date = st.date_input("날짜 선택", datetime.today(), key="tab1_date")
    with col2:
        target_time = st.time_input("시간 선택", value=time(12, 0), key="tab1_time")

    if st.button("실시간 조회", key="tab1_btn"):
        # 00:00~07:29 시간대는 전날 07:30에 시작한 근무일에 속하므로, 그 근무일 기준으로 데이터를 모음
        if target_time >= time(7, 30):
            shift_start_date = target_date
        else:
            shift_start_date = target_date - timedelta(days=1)
        today_data, today_daeche = get_shift_day_data(shift_start_date)

        result = {}
        for post in 배치['소속초소'].unique():
            on_duty, off_duty = [], []
            for name in 전체_경비원명:
                loc, is_off, task_name = get_person_status(name, target_time, today_data, today_daeche)
                if loc != post:
                    continue
                if is_off:
                    off_duty.append({'경비원명': name, '상태': task_name})
                else:
                    on_duty.append(name)
            result[post] = (on_duty, off_duty)

        st.session_state.last_query_result = result
        st.session_state.last_query_time = f"{target_date} {target_time.strftime('%H:%M')}"

    # 탭 이동 후에도 마지막 조회 결과를 유지
    if 'last_query_result' in st.session_state:
        st.caption(f"조회 시각: {st.session_state.last_query_time}")

        for post, (on_duty, off_duty) in st.session_state.last_query_result.items():
            st.subheader(f"🏢 {post}")
            st.write(f"**초소인원 ({len(on_duty)}명)**: {', '.join(on_duty) if on_duty else '없음'}")
            if off_duty:
                st.table(pd.DataFrame(off_duty))

# ────────────────────────────────────────────────
# 탭 2: 30분 단위 히트맵
# ────────────────────────────────────────────────
with tab2:
    available_dates = sorted(근무데이터['날짜'].dt.date.unique())
    selected_date = st.selectbox("조회할 날짜 선택", available_dates, key="tab2_date")

    if st.button("히트맵 생성", key="tab2_btn"):
        with st.spinner("계산 중..."):
            start_dt = datetime.combine(selected_date, time(7, 30))
            time_slots = [(start_dt + timedelta(minutes=30 * i)).time() for i in range(48)]
            posts = list(배치['소속초소'].unique())

            # [중요] 야간 근무(00:00~07:29) 데이터는 익일 날짜로 기록되어 있으므로
            # 07:30~익일 07:30을 하나의 근무일로 보려면 두 날짜 데이터를 모두 모아야 합니다.
            today_data, today_daeche = get_shift_day_data(selected_date)

            # 각 초소·시간대별로 초소인원(on)/비초소인원(off) 항목을 각각 모아둠
            results = {t.strftime('%H:%M'): {post: {'on': [], 'off': []} for post in posts} for t in time_slots}

            for t in time_slots:
                for name in 전체_경비원명:
                    loc, is_off, _ = get_person_status(name, t, today_data, today_daeche)
                    약칭 = 약칭_dict.get(name, name[:1])
                    if not is_off:
                        # 북초소 대체담당자가 동적 백업(코드 없이 자동 판정)으로 북초소에 들어온 경우는
                        # 대체근무자 시트 코드가 없으므로 약칭만 표시
                        if name == 북초소_대체담당자 and loc == '북초소' and get_active_sub_code(name, t, today_daeche) is None:
                            entry = 약칭
                        else:
                            sub_code = get_active_sub_code(name, t, today_daeche)
                            entry = f"{약칭}({sub_code})" if sub_code else 약칭
                        results[t.strftime('%H:%M')][loc]['on'].append(entry)
                    else:
                        own_code = get_active_own_code(name, t, today_data)
                        entry = f"{약칭}({own_code})" if own_code else 약칭
                        results[t.strftime('%H:%M')][loc]['off'].append(entry)

            # 초소인원은 쉼표로, 비초소인원은 그 뒤에 " / "로 구분해 이어붙임
            # 예: "근(대체-남A),학 / 창(휴-1B),운(주-3)"
            def build_cell(entry_dict):
                on_part = ','.join(entry_dict['on'])
                off_part = ','.join(entry_dict['off'])
                if off_part:
                    return f"{on_part} / {off_part}" if on_part else f" / {off_part}"
                return on_part

            matrix = pd.DataFrame.from_dict(
                {t: {post: build_cell(entry_dict) for post, entry_dict in posts_dict.items()} for t, posts_dict in results.items()},
                orient='index'
            )
            st.session_state.last_heatmap_data = matrix

    # 탭 이동 후에도 마지막 히트맵 결과를 유지
    if 'last_heatmap_data' in st.session_state:
        matrix = st.session_state.last_heatmap_data

        def color_coding(val):
            # 색상은 초소인원(근무 중) 수 기준으로만 판정 - " / " 앞부분(초소인원)만 계산
            on_part = str(val).split(' / ')[0] if val else ''
            인원수 = len([v for v in on_part.split(',') if v]) if on_part else 0
            if 인원수 == 0:
                color = '#FF0000'   # 0명: 빨강
            elif 인원수 == 1:
                color = '#0000FF'   # 1명: 파랑
            else:
                color = '#008000'   # 2명 이상: 녹색
            return f'background-color: {color}; color: white; font-weight: bold; text-align: center; font-size: 16px;'

        styled = matrix.style.map(color_coding)

        # y축(시간대 인덱스)과 초소명 헤더가 어두운 배경에 묻히지 않도록 고대비 스타일 명시 적용
        styled = styled.set_table_styles([
            {'selector': 'th.row_heading', 'props': [
                ('background-color', '#222222'),
                ('color', '#FFFF00'),
                ('font-weight', 'bold'),
                ('font-size', '16px'),
                ('text-align', 'center'),
            ]},
            {'selector': 'th.col_heading', 'props': [
                ('background-color', '#222222'),
                ('color', '#FFFF00'),
                ('font-weight', 'bold'),
                ('font-size', '16px'),
                ('text-align', 'center'),
            ]},
        ], overwrite=False)

        st.table(styled)
