import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import re
import requests
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="보전팀 통합 분석 시스템",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .main-title {font-size:26px; font-weight:700; color:#1e3a5f; margin-bottom:2px;}
    .sub-title  {font-size:12px; color:#888; margin-bottom:16px;}
    .kpi-box    {background:#f0f4fa; border-radius:10px; padding:14px 18px; text-align:center; height:90px;}
    .kpi-val    {font-size:28px; font-weight:700; color:#1e3a5f;}
    .kpi-unit   {font-size:14px; color:#888;}
    .kpi-label  {font-size:11px; color:#666; margin-top:4px;}
    .warn-box   {background:#fff3cd; border-left:4px solid #ffc107;
                 padding:10px 14px; border-radius:4px; margin:8px 0; font-size:13px;}
    .ok-box     {background:#d4edda; border-left:4px solid #28a745;
                 padding:10px 14px; border-radius:4px; margin:8px 0; font-size:13px;}
    .err-box    {background:#f8d7da; border-left:4px solid #dc3545;
                 padding:10px 14px; border-radius:4px; margin:8px 0; font-size:13px;}
    .stTabs [data-baseweb="tab"] {font-size:14px; font-weight:600;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────
# 상수 및 정규화 맵
# ─────────────────────────────────────────────────────
VALID_YEAR_MIN, VALID_YEAR_MAX = 2010, 2030

EQUIP_NORM = {
    '로보트':'로봇', '로보트 ':'로봇', ' 로보트':'로봇', '로봇 ':'로봇', ' 로봇':'로봇',
    '블래킹':'블랭킹', '플랭킹':'블랭킹', 'B/K':'블랭킹',
    '파일러1':'파일러', '파일러 ':'파일러',
    '지하컨베어':'컨베어', '텔레스코픽컨베어':'텔레스코프',
    '1500T':'프레스-1500T', '1200T':'프레스-1200T',
    '800T':'프레스-800T',   '600T':'프레스-600T',
}

# ─────────────────────────────────────────────────────
# 유틸 함수
# ─────────────────────────────────────────────────────
def parse_dt(val):
    """다양한 datetime 형식 파싱 (날짜 전용 문자열 / pandas Timestamp 포함)"""
    if val is None:
        return None
    if isinstance(val, float):
        if np.isnan(val): return None
        return None
    if isinstance(val, datetime):
        return val
    # pandas Timestamp 처리
    if hasattr(val, 'to_pydatetime'):
        try: return val.to_pydatetime()
        except: return None
    if isinstance(val, str):
        val = val.strip()
        if val in ('', '00:00:00', 'None', '0', 'NaT', 'nan', 'NaN'):
            return None
        for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d',
                    '%Y/%m/%d %H:%M:%S', '%Y/%m/%d %H:%M', '%Y/%m/%d',
                    '%Y.%m.%d %H:%M:%S', '%Y.%m.%d %H:%M', '%Y.%m.%d']:
            try:
                return datetime.strptime(val, fmt)
            except:
                pass
    return None

def sanitize_dt(dt):
    """이상 연도 제거 (2010~2030 외 None)"""
    if dt is None:
        return None
    try:
        return dt if VALID_YEAR_MIN <= dt.year <= VALID_YEAR_MAX else None
    except:
        return None

def to_float_safe(v):
    try:
        f = float(v)
        # 0이하 또는 1440분(24h) 초과는 오입력 제거
        return f if 0 < f <= 1440 else None
    except:
        return None

def norm_equip(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    # 숫자(int/float)는 str 변환 후 정규화
    s = str(v).strip() if not isinstance(v, str) else v.strip()
    if s in ('', 'nan', 'None', '0'): return None
    return EQUIP_NORM.get(s, s)

def classify_fault(현상, 원인):
    text = (str(현상 or '') + str(원인 or '')).lower()
    if any(k in text for k in ['케이블','단선','전기','전원','퓨즈','eocr','인버터','센서','스위치','접촉','누전']):
        return '전기/제어'
    if any(k in text for k in ['오일','유압','누유','펌프','실린더','밸브','압력','윤활','그리스']):
        return '유압/윤활'
    if any(k in text for k in ['베어링','볼트','마모','균열','파손','벨트','체인','기어','용접','변형']):
        return '기계적결함'
    if any(k in text for k in ['티칭','teaching','프로그램','통신','plc','로직','파라미터','설정값']):
        return '제어/프로그램'
    if any(k in text for k in ['작업자','이종','투입','조작','인위','사람']):
        return '작업자과실'
    if any(k in text for k in ['예방','pm','계획','정기점검']):
        return '예방보전'
    return '기타/불명'

def classify_action(조치):
    text = str(조치 or '').lower()
    if any(k in text for k in ['교체','교환','신품','부품','spare']):
        return '부품교체'
    if any(k in text for k in ['티칭','teaching','티이칭']):
        return '티칭수정'
    if any(k in text for k in ['예방','pm ','p/m','계획','정기']):
        return '예방보전'
    if any(k in text for k in ['조정','셋팅','설정','세팅','재조임','조임']):
        return '조정/설정'
    if any(k in text for k in ['리셋','reset','재기동','재가동','원복','복구']):
        return '긴급복구'
    return '기타'

def parse_workers(조치자_val):
    """조치자 문자열 → 개인 리스트
    구분자: , / . 공백 줄바꿈 + 모두 처리
    조합 문자열 자체는 선택지에 노출되지 않음"""
    NOISE = {'야간', '주간', '주간조', '야간조', '주야간', '업체', '가동중', '조치', '기타'}
    if not 조치자_val or not isinstance(조치자_val, str):
        return []
    # 구분자: 콤마, 슬래시, 마침표, 공백, 줄바꿈, + 모두 처리
    workers = [w.strip() for w in re.split(r'[,/.\s\+]+', 조치자_val) if w.strip()]
    workers = [w for w in workers if len(w) >= 2]
    # 한글 포함된 이름만 유효 (숫자/영문만인 잡음 제거)
    workers = [w for w in workers if re.search(r'[가-힣]', w)]
    # 잡음 단어 제거
    workers = [w for w in workers if w not in NOISE]
    return workers

# ─────────────────────────────────────────────────────
# ★ 로봇/지그 세부분류 함수
# ─────────────────────────────────────────────────────
def get_세부분류(row):
    """설비유형 + 고장설비 + 고장부위 기반 세부분류 생성"""
    유형 = str(row.get('설비유형') or '').strip()
    설비 = str(row.get('고장설비') or '').strip().upper()
    부위 = str(row.get('고장부위') or '').strip()

    if '로봇' in 유형 or '로보트' in 유형:
        # 번호 추출
        m = re.match(r'R(\d+)', 설비)
        if m:
            n = int(m.group(1))
            if 'R1-' in 설비 or 'R2-' in 설비:
                그룹 = '로봇-서브(R1-x)'
            elif n <= 5:
                그룹 = f'로봇-{설비}(메인)'
            elif n <= 10:
                그룹 = f'로봇-{설비}'
            elif n <= 20:
                그룹 = f'로봇-{설비}'
            else:
                그룹 = f'로봇-{설비}'
        else:
            그룹 = '로봇-기타'
        return 그룹

    if '지그' in 유형:
        if 설비.startswith('A'):
            return '지그-A계열(조립)'
        if 설비.startswith('S'):
            return '지그-S계열(스터드)'
        if 'PLT' in 설비:
            return '지그-PLT(팔레트)'
        if 설비.startswith('CS'):
            return '지그-CS계열'
        if 설비.startswith('FA'):
            return '지그-FA계열'
        if 설비.startswith('M'):
            return '지그-M계열'
        return '지그-기타'

    return 유형 if 유형 else '기타'

def get_고장부위_그룹(부위_raw):
    """고장부위_STD 기반 표준 그룹 분류"""
    v = str(부위_raw or '').strip()
    if not v or v == 'nan': return '기타'
    if '일시정지' in v: return '일시정지'
    if '에러' in v: return '에러'
    if 'L/S' in v or 'LS' in v: return 'L/S이상'
    if '스터드' in v: return '스터드'
    if 'T/C' in v or 'TC' in v: return 'T/C이상'
    if '센서' in v: return '센서'
    if '실러' in v: return '실러'
    if '품질' in v: return '품질'
    if '냉각' in v: return '냉각수'
    if '그리퍼' in v: return '그리퍼'
    if '티칭' in v: return '티칭수정'
    if 'ATD' in v: return 'ATD'
    if 'PW' in v: return 'PW'
    if '파트' in v: return '파트이상'
    if 'AIR' in v: return 'AIR'
    return '기타'

# ─────────────────────────────────────────────────────
# 파일 로드
# ─────────────────────────────────────────────────────
def load_press(file_obj):
    # 시트명 유연하게 처리
    xf = pd.ExcelFile(file_obj)
    sheet = '설비보전현황_통합' if '설비보전현황_통합' in xf.sheet_names else xf.sheet_names[0]
    file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name=sheet, header=0)
    df.columns = ['년','월','일','주','라인','정지시각','출동시각','완료시각',
                  '소요시간','설비유형','고장설비','고장부위','현상','원인','조치내역','조치자','비고']

    # 문자열 컬럼: int/float 혼재 값을 str로 통일, NaN은 None 처리
    for col in ['라인','설비유형','고장설비','고장부위','현상','원인','조치내역','조치자','비고']:
        df[col] = df[col].apply(
            lambda v: None if (v is None or (isinstance(v,float) and pd.isna(v)))
            else str(v).strip() if str(v).strip() not in ('nan','None','') else None
        )

    for col in ['정지시각','출동시각','완료시각']:
        df[col] = df[col].apply(parse_dt).apply(sanitize_dt)

    # 발생일시: 정지 -> 출동 -> 완료 -> 년월일 순 폴백
    # pd.notna() 사용 — Timestamp/NaT bool 평가 오류 방지
    def make_dt(row):
        if pd.notna(row['정지시각']): return row['정지시각']
        if pd.notna(row['출동시각']): return row['출동시각']
        if pd.notna(row['완료시각']): return row['완료시각']
        try:
            y, m, d = int(row['년']), int(row['월']), int(row['일'])
            if 1 <= m <= 12 and 1 <= d <= 31:
                return datetime(y, m, d)
        except:
            pass
        return None

    df['발생일시'] = df.apply(make_dt, axis=1).apply(sanitize_dt)

    # 소요시간: 원본 → 완료-출동
    def fix_dur(row):
        v = to_float_safe(row['소요시간'])
        if v: return v
        if pd.notna(row['완료시각']) and pd.notna(row['출동시각']):
            try:
                d = (row['완료시각'] - row['출동시각']).total_seconds() / 60
                return round(d, 1) if d > 0 else None
            except: return None
        return None

    df['소요시간'] = df.apply(fix_dur, axis=1)
    df['설비유형'] = df['설비유형'].apply(norm_equip)
    df['고장분류'] = df.apply(lambda r: classify_fault(r['현상'], r['원인']), axis=1)
    df['조치유형'] = df['조치내역'].apply(classify_action)
    df['파일출처'] = '프레스'
    df['차종'] = None
    df['조치'] = df['조치내역']

    # ★ 라인+설비 복합키 (프레스는 차종 없음 → 라인 그대로)
    df['라인_차종'] = df['라인'].astype(str).str.strip()
    df['설비_KEY'] = df['라인_차종'] + ' | ' + df['고장설비'].astype(str).str.strip()
    return df

def load_robot(file_obj):
    df = pd.read_excel(file_obj, sheet_name='Sheet1', header=0)
    df.columns = ['발생일시_raw','월','일','주','라인','라인_KEY','차종','설비유형',
                  '고장설비','고장부위','고장부위_STD','현상','원인','조치',
                  '소요시간','정지시각','출동시각','완료시각','조치자','비고','NO']

    for col in ['정지시각','출동시각','완료시각']:
        df[col] = df[col].apply(parse_dt).apply(sanitize_dt)

    def make_dt(row):
        dt = sanitize_dt(parse_dt(row['발생일시_raw']))
        if dt: return dt
        if pd.notna(row['출동시각']): return row['출동시각']
        if pd.notna(row['정지시각']): return row['정지시각']
        try: return sanitize_dt(datetime(2025, int(row['월']), int(row['일'])))
        except: return None

    df['발생일시'] = df.apply(make_dt, axis=1)

    def fix_dur(row):
        v = to_float_safe(row['소요시간'])
        if v: return v
        if row['완료시각'] and row['출동시각']:
            try:
                d = (row['완료시각'] - row['출동시각']).total_seconds() / 60
                return round(d, 1) if d > 0 else None
            except: return None
        return None

    df['소요시간'] = df.apply(fix_dur, axis=1)
    df['설비유형'] = df['설비유형'].apply(norm_equip)
    df['고장분류'] = df.apply(lambda r: classify_fault(r['현상'], r['원인']), axis=1)
    df['조치유형'] = df['조치'].apply(classify_action)
    df['파일출처'] = '로봇/지그'
    df['조치내역'] = df['조치']
    df['년'] = df['발생일시'].apply(lambda x: x.year if x else None)

    # ★ 라인+차종 복합키 (차종 있는 파일만)
    # 차종 nan 정제 후 라인_차종 생성
    df['차종_clean'] = df['차종'].fillna('').astype(str).str.strip().replace({'nan':'', 'NaN':'', 'None':''})
    def make_차종라인(r):
        라인 = str(r['라인']).strip() if pd.notna(r['라인']) and str(r['라인']).strip() not in ('','nan','NaN') else ''
        차종 = r['차종_clean']
        if 차종 and 라인: return f"{차종} / {라인}"
        if 라인:           return 라인
        if 차종:           return f"{차종} / (미상)"
        return '미상'
    df['라인_차종'] = df.apply(make_차종라인, axis=1)
    # ★ 라인+차종+설비 복합키
    df['설비_KEY'] = df['라인_차종'] + ' | ' + df['고장설비'].astype(str).str.strip()
    return df

def detect_and_load(file_obj, fname=''):
    try:
        xf = pd.ExcelFile(file_obj)
        sheets = xf.sheet_names
        file_obj.seek(0)
        # 프레스 파일: 설비보전현황_통합 시트
        if '설비보전현황_통합' in sheets:
            return load_press(file_obj), 'press'
        # 로봇/지그 파일: Sheet1 또는 첫번째 시트
        target = 'Sheet1' if 'Sheet1' in sheets else sheets[0]
        df_h = pd.read_excel(file_obj, sheet_name=target, nrows=2)
        file_obj.seek(0)
        cols = df_h.columns.tolist()
        # 로봇/지그 파일 판별: 라인_KEY 또는 차종 컬럼 존재
        if any(c in cols for c in ['라인_KEY','차종','고장부위_STD']):
            return load_robot(file_obj), 'robot'
        # 프레스 파일 판별: 년,월,일,라인 컬럼 존재
        if all(c in cols for c in ['년','월','라인','설비유형']):
            return load_press(file_obj), 'press'
        # 컬럼수로 구분: 17개=프레스, 21개=로봇/지그
        if len(cols) >= 20:
            return load_robot(file_obj), 'robot'
        if len(cols) >= 15:
            return load_press(file_obj), 'press'
        return None, f'인식불가 (시트: {target}, 컬럼수: {len(cols)})'
    except Exception as e:
        return None, f'오류: {e}'

def merge_dfs(press_df, robot_df):
    COMMON = ['발생일시','년','월','일','라인','차종','라인_차종','설비유형','고장설비','고장부위','설비_KEY',
              '현상','원인','조치내역','조치자','소요시간','정지시각','출동시각',
              '완료시각','비고','조치유형','고장분류','파일출처']
    frames = []
    if press_df is not None:
        frames.append(press_df[[c for c in COMMON if c in press_df.columns]])
    if robot_df is not None:
        frames.append(robot_df[[c for c in COMMON if c in robot_df.columns]])
    if not frames:
        return None
    merged = pd.concat(frames, ignore_index=True)
    merged = merged[merged['발생일시'].notna()].copy()
    merged['년'] = merged['발생일시'].dt.year
    merged['월'] = merged['발생일시'].dt.month
    # 라인_차종 없으면 라인으로 채움
    if '라인_차종' not in merged.columns:
        merged['라인_차종'] = merged['라인'].astype(str).str.strip()
    else:
        merged['라인_차종'] = merged['라인_차종'].fillna(merged['라인'].astype(str).str.strip())
    # 차종 없는 경우(프레스 등) 라인_차종 = 라인
    mask_no_car = merged['라인_차종'].astype(str).str.endswith(' / ') |                   merged['라인_차종'].astype(str).str.contains(r'/ $', regex=True)
    merged.loc[mask_no_car, '라인_차종'] = merged.loc[mask_no_car, '라인'].astype(str).str.strip()
    # ★ 세부분류 컬럼 생성
    merged['세부분류'] = merged.apply(get_세부분류, axis=1)
    merged['부위그룹'] = merged['고장부위'].apply(get_고장부위_그룹) if '고장부위' in merged.columns else '기타'
    merged.sort_values('발생일시', inplace=True)
    merged.reset_index(drop=True, inplace=True)
    return merged

def load_from_gdrive(url):
    try:
        url = url.strip()
        file_id = None
        for p in [r'spreadsheets/d/([a-zA-Z0-9_-]+)',
                  r'/d/([a-zA-Z0-9_-]+)',
                  r'id=([a-zA-Z0-9_-]+)']:
            m = re.search(p, url)
            if m:
                file_id = m.group(1)
                break
        if not file_id:
            return None, '파일 ID를 찾을 수 없습니다.'
        # 스프레드시트 URL -> xlsx export 사용
        if 'spreadsheets' in url or 'docs.google.com' in url:
            dl_url = ('https://docs.google.com/spreadsheets/d/'
                      + file_id + '/export?format=xlsx')
        else:
            dl_url = ('https://drive.google.com/uc'
                      '?export=download&id=' + file_id)
        headers = {'User-Agent': 'Mozilla/5.0'}
        resp = requests.get(dl_url, timeout=30, headers=headers)
        # 바이러스 경고 페이지 처리
        if resp.status_code == 200 and b'confirm=' in resp.content[:2000]:
            cm = re.search(rb'confirm=([0-9A-Za-z_-]+)', resp.content)
            if cm:
                resp = requests.get(
                    dl_url + '&confirm=' + cm.group(1).decode(),
                    timeout=30, headers=headers)
        if resp.status_code != 200:
            return None, f'다운로드 실패 (HTTP {resp.status_code})'
        if len(resp.content) < 500:
            return None, ('파일이 너무 작습니다 — '
                          '공유 권한을 "링크 있는 모든 사용자"로 설정해주세요.')
        return io.BytesIO(resp.content), None
    except Exception as e:
        return None, str(e)

# ─────────────────────────────────────────────────────
# 분석 함수
# ─────────────────────────────────────────────────────
def calc_mttr_mtbf(df):
    """라인+설비 복합키 기준 MTTR/MTBF"""
    results = []
    df2 = df[df['소요시간'].notna()].copy()
    for key, grp in df2.groupby('설비_KEY'):
        grp = grp.sort_values('발생일시')
        parts = key.split(' | ')
        라인 = parts[0] if len(parts) > 0 else ''
        설비 = parts[1] if len(parts) > 1 else ''
        mttr = grp['소요시간'].mean()
        cnt = len(grp)
        if cnt >= 2:
            gaps = grp['발생일시'].diff().dropna().dt.total_seconds() / 3600
            gaps = gaps[gaps > 0]
            mtbf = gaps.mean() if len(gaps) else None
        else:
            mtbf = None
        results.append({
            '라인': 라인,
            '고장설비': 설비,
            '설비_KEY': key,
            '발생건수': cnt,
            'MTTR(분)': round(mttr, 1),
            'MTBF(시간)': round(mtbf, 1) if mtbf else None,
            '총정지시간(분)': round(grp['소요시간'].sum(), 1),
            '설비유형': grp['설비유형'].mode()[0] if not grp['설비유형'].isna().all() else '',
        })
    return pd.DataFrame(results).sort_values('총정지시간(분)', ascending=False)

def get_worker_df(df):
    """조치자 파싱 → 인원별 롱포맷"""
    rows = []
    for _, r in df.iterrows():
        for w in parse_workers(r.get('조치자')):
            rows.append({
                '조치자': w,
                '소요시간': r['소요시간'] if pd.notna(r.get('소요시간')) else 0,
                '발생일시': r['발생일시'],
                '라인': r.get('라인'),
                '라인_차종': r.get('라인_차종', r.get('라인','')),
                '설비유형': r.get('설비유형'),
                '고장설비': r.get('고장설비'),
            })
    return pd.DataFrame(rows) if rows else pd.DataFrame()

def calc_response_time(df):
    mask = df['정지시각'].notna() & df['출동시각'].notna()
    sub = df[mask].copy()
    if sub.empty: return None
    sub['응답시간_분'] = sub.apply(
        lambda r: (r['출동시각'] - r['정지시각']).total_seconds() / 60
        if r['출동시각'] > r['정지시각'] else None, axis=1)
    return sub[sub['응답시간_분'].notna() & (sub['응답시간_분'] > 0) & (sub['응답시간_분'] < 240)]

def to_excel(df_dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
        for name, df in df_dict.items():
            df.to_excel(w, sheet_name=name[:31], index=False)
    return buf.getvalue()

# ─────────────────────────────────────────────────────
# 세션 상태
# ─────────────────────────────────────────────────────
for k in ['press_df','robot_df','merged_df']:
    if k not in st.session_state:
        st.session_state[k] = None

# ─────────────────────────────────────────────────────
# 헤더
# ─────────────────────────────────────────────────────
st.markdown('<div class="main-title">🔧 보전팀 통합 분석 시스템</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">호원오토 평택공장 | 보전관리 2팀</div>', unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(
    ["📂 데이터 불러오기", "📊 고장현황 (Pareto)", "⚙️ 설비분석 (MTTR/MTBF)",
     "👷 인원분석",
     "🏆 설비 위험도", "⏱️ 유실시간 분석", "🔧 예방정비 추천",
     "📥 출력"])

# ══════════════════════════════════════════════════════
# TAB 1 — 데이터 불러오기
# ══════════════════════════════════════════════════════
with tab1:
    st.subheader("데이터 불러오기")
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### ① 구글 드라이브 공유 링크")
        gdrive_url = st.text_input("링크 붙여넣기", placeholder="https://drive.google.com/file/d/...")
        if st.button("🔗 드라이브에서 불러오기", use_container_width=True):
            if gdrive_url:
                with st.spinner("다운로드 중..."):
                    fobj, err = load_from_gdrive(gdrive_url)
                    if err:
                        st.error(f"오류: {err}")
                    else:
                        result, ftype = detect_and_load(fobj)
                        if ftype == 'press':
                            st.session_state.press_df = result
                            st.success(f"✅ 프레스 파일 로드 — {len(result):,}건")
                        elif ftype == 'robot':
                            st.session_state.robot_df = result
                            st.success(f"✅ 로봇/지그 파일 로드 — {len(result):,}건")
                        else:
                            if 'unknown' in str(ftype) or '인식불가' in str(ftype):
                                st.error(
                                    f"파일 형식 인식 실패: {ftype}\n"
                                    "해결방법: 구글시트는 파일 > 다운로드 > .xlsx 로 저장 후 직접 업로드 해주세요.")
                            else:
                                st.warning(f"파일 형식 인식 실패: {ftype}")
            else:
                st.warning("링크를 입력해주세요.")

    with c2:
        st.markdown("#### ② 직접 파일 업로드 (.xlsx / .csv)")
        uploaded = st.file_uploader("파일 선택 (여러 파일 동시 가능)",
                                    type=['xlsx','csv'], accept_multiple_files=True)
        if uploaded:
            for uf in uploaded:
                with st.spinner(f"{uf.name} 처리 중..."):
                    try:
                        if uf.name.endswith('.csv'):
                            st.info(f"CSV '{uf.name}' — xlsx 권장")
                        else:
                            result, ftype = detect_and_load(uf, uf.name)
                            if ftype == 'press':
                                st.session_state.press_df = result
                                st.success(f"✅ 프레스 '{uf.name}' — {len(result):,}건")
                            elif ftype == 'robot':
                                st.session_state.robot_df = result
                                st.success(f"✅ 로봇/지그 '{uf.name}' — {len(result):,}건")
                            else:
                                st.warning(f"'{uf.name}' 형식 인식 실패: {ftype}")
                    except Exception as e:
                        st.error(f"'{uf.name}' 처리 오류: {e}")

    st.divider()
    if st.button("🔄 데이터 통합 실행", type="primary", use_container_width=True):
        merged = merge_dfs(st.session_state.press_df, st.session_state.robot_df)
        if merged is not None:
            st.session_state.merged_df = merged
            yr_min = int(merged['년'].min())
            yr_max = int(merged['년'].max())
            st.success(f"✅ 통합 완료 — {len(merged):,}건 ({yr_min}~{yr_max}년)")
        else:
            st.warning("불러온 데이터가 없습니다.")

    s1, s2, s3 = st.columns(3)
    with s1:
        if st.session_state.press_df is not None:
            st.markdown(f'<div class="ok-box">✅ 프레스 파일<br>{len(st.session_state.press_df):,}건 로드됨</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="warn-box">⬜ 프레스 파일 미로드</div>', unsafe_allow_html=True)
    with s2:
        if st.session_state.robot_df is not None:
            st.markdown(f'<div class="ok-box">✅ 로봇/지그 파일<br>{len(st.session_state.robot_df):,}건 로드됨</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="warn-box">⬜ 로봇/지그 파일 미로드</div>', unsafe_allow_html=True)
    with s3:
        if st.session_state.merged_df is not None:
            df = st.session_state.merged_df
            st.markdown(f'<div class="ok-box">✅ 통합 완료<br>{len(df):,}건 ({int(df["년"].min())}~{int(df["년"].max())}년)</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="warn-box">⬜ 통합 미실행</div>', unsafe_allow_html=True)

    if st.session_state.merged_df is not None:
        with st.expander("📋 통합 데이터 미리보기 (상위 200건)"):
            preview_cols = ['발생일시','라인','설비_KEY','설비유형','고장설비','고장부위',
                            '현상','원인','조치내역','조치자','소요시간','고장분류','조치유형']
            st.dataframe(st.session_state.merged_df[[c for c in preview_cols
                         if c in st.session_state.merged_df.columns]].head(200),
                         use_container_width=True)

# ══════════════════════════════════════════════════════
# TAB 2 — 고장현황 (Pareto)
# ══════════════════════════════════════════════════════
with tab2:
    df = st.session_state.merged_df
    if df is None:
        st.info("Tab1에서 데이터를 불러오고 '통합 실행'을 눌러주세요.")
    else:
        st.subheader("고장현황 분석")

        # ── 필터 ──
        fc1, fc2, fc3, fc4, fc5 = st.columns([2,2,2,2,1])
        with fc1:
            yrs = sorted(df['년'].dropna().unique().astype(int))
            sel_yr = st.multiselect("연도", yrs, default=yrs, key='t2y')
        with fc2:
            equips = ['전체'] + sorted(df['설비유형'].dropna().unique().tolist(), key=str)
            sel_eq = st.selectbox("설비유형", equips, key='t2e')
        with fc3:
            lines = ['전체'] + sorted(df['라인'].dropna().unique().tolist(), key=str)
            sel_ln = st.selectbox("라인", lines, key='t2l')
        with fc4:
            # 차종 필터 (라인_차종이 있는 경우만)
            if '차종' in df.columns:
                cars = ['전체'] + sorted(df['차종'].dropna().unique().tolist(), key=str)
            else:
                cars = ['전체']
            sel_car = st.selectbox("차종", cars, key='t2car')
        with fc5:
            top_n = st.slider("Top N", 5, 30, 20, key='t2n')

        fdf = df[df['년'].isin(sel_yr)].copy() if sel_yr else df.copy()
        if sel_eq  != '전체': fdf = fdf[fdf['설비유형'] == sel_eq]
        if sel_ln  != '전체': fdf = fdf[fdf['라인'] == sel_ln]
        if sel_car != '전체' and '차종' in fdf.columns:
            fdf = fdf[fdf['차종'] == sel_car]

        # ── KPI ──
        total_cnt  = len(fdf)
        total_stop = fdf['소요시간'].sum()
        avg_mttr   = fdf['소요시간'].mean()
        resp_df    = calc_response_time(fdf)
        avg_resp   = resp_df['응답시간_분'].mean() if resp_df is not None and not resp_df.empty else 0
        k1,k2,k3,k4 = st.columns(4)
        for col, label, val, unit in [
            (k1,'총 고장건수',f'{total_cnt:,}','건'),
            (k2,'총 정지시간',f'{total_stop/60:.0f}','시간'),
            (k3,'평균 MTTR',f'{avg_mttr:.0f}','분'),
            (k4,'평균 응답시간',f'{avg_resp:.0f}','분'),
        ]:
            col.markdown(
                f'<div class="kpi-box"><div class="kpi-val">{val}'
                f'<span class="kpi-unit"> {unit}</span></div>'
                f'<div class="kpi-label">{label}</div></div>',
                unsafe_allow_html=True)
        st.divider()

        # ══ ① Pareto: 설비유형별 (상위 2줄) ══
        p1, p2 = st.columns(2)
        with p1:
            st.markdown("##### 설비유형별 Pareto — 건수")
            grp = (fdf.groupby('설비유형')
                   .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'))
                   .reset_index().sort_values('건수', ascending=False).head(top_n))
            grp['누적%'] = (grp['건수'].cumsum() / grp['건수'].sum() * 100).round(1)
            grp['평균MTTR'] = (grp['총정지시간'] / grp['건수']).round(1)
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_trace(go.Bar(
                x=grp['설비유형'], y=grp['건수'], name='건수',
                marker_color='#1e3a5f',
                customdata=np.stack([grp['총정지시간'], grp['평균MTTR'], grp['누적%']], axis=-1),
                hovertemplate='<b>%{x}</b><br>건수: %{y:,}건<br>총정지: %{customdata[0]:.0f}분<br>평균MTTR: %{customdata[1]:.1f}분<br>누적: %{customdata[2]:.1f}%<extra></extra>'
            ), secondary_y=False)
            fig.add_trace(go.Scatter(
                x=grp['설비유형'], y=grp['누적%'], name='누적%',
                line=dict(color='#e74c3c', width=2), mode='lines+markers',
                hovertemplate='누적: %{y:.1f}%<extra></extra>'
            ), secondary_y=True)
            fig.add_hline(y=80, line_dash='dash', line_color='orange',
                          annotation_text='80%', secondary_y=True)
            fig.update_yaxes(title_text='건수', secondary_y=False)
            fig.update_yaxes(title_text='누적 %', range=[0,105], secondary_y=True)
            fig.update_layout(height=320, margin=dict(t=10,b=40,l=10,r=60))
            st.plotly_chart(fig, use_container_width=True)

        with p2:
            st.markdown("##### 설비유형별 Pareto — 정지시간")
            grp2 = (fdf.groupby('설비유형')
                    .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'))
                    .reset_index().sort_values('총정지시간', ascending=False).head(top_n))
            grp2['누적%'] = (grp2['총정지시간'].cumsum() / grp2['총정지시간'].sum() * 100).round(1)
            grp2['평균MTTR'] = (grp2['총정지시간'] / grp2['건수']).round(1)
            fig2 = make_subplots(specs=[[{"secondary_y": True}]])
            fig2.add_trace(go.Bar(
                x=grp2['설비유형'], y=grp2['총정지시간'], name='정지시간(분)',
                marker_color='#c0392b',
                customdata=np.stack([grp2['건수'], grp2['평균MTTR'], grp2['누적%']], axis=-1),
                hovertemplate='<b>%{x}</b><br>총정지: %{y:,.0f}분<br>건수: %{customdata[0]:,}건<br>평균MTTR: %{customdata[1]:.1f}분<br>누적: %{customdata[2]:.1f}%<extra></extra>'
            ), secondary_y=False)
            fig2.add_trace(go.Scatter(
                x=grp2['설비유형'], y=grp2['누적%'], name='누적%',
                line=dict(color='#2980b9', width=2), mode='lines+markers',
                hovertemplate='누적: %{y:.1f}%<extra></extra>'
            ), secondary_y=True)
            fig2.add_hline(y=80, line_dash='dash', line_color='orange', secondary_y=True)
            fig2.update_layout(height=320, margin=dict(t=10,b=40,l=10,r=60))
            st.plotly_chart(fig2, use_container_width=True)

        # ── 라인/차종별 Pareto (전체 너비) ──
        st.markdown("##### 차종·라인별 Pareto — 건수 / 정지시간")
        lc1, lc2 = st.columns(2)
        with lc1:
            # 라인_차종 건수 Pareto
            lc_grp = (fdf.groupby('라인_차종')
                      .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'),
                           라인=('라인','first'), 차종=('차종','first') if '차종' in fdf.columns else ('라인','first'))
                      .reset_index().sort_values('건수', ascending=False).head(top_n))
            lc_grp['누적%'] = (lc_grp['건수'].cumsum() / lc_grp['건수'].sum() * 100).round(1)
            lc_grp['평균MTTR'] = (lc_grp['총정지시간'] / lc_grp['건수']).round(1)
            fig_lc1 = make_subplots(specs=[[{"secondary_y": True}]])
            fig_lc1.add_trace(go.Bar(
                x=lc_grp['라인_차종'], y=lc_grp['건수'], name='건수',
                marker_color='#2471a3',
                customdata=np.stack([lc_grp['총정지시간'], lc_grp['평균MTTR'], lc_grp['누적%']], axis=-1),
                hovertemplate='<b>%{x}</b><br>건수: %{y:,}건<br>총정지: %{customdata[0]:.0f}분<br>평균MTTR: %{customdata[1]:.1f}분<br>누적: %{customdata[2]:.1f}%<extra></extra>'
            ), secondary_y=False)
            fig_lc1.add_trace(go.Scatter(
                x=lc_grp['라인_차종'], y=lc_grp['누적%'], name='누적%',
                line=dict(color='#e74c3c', width=2), mode='lines+markers',
                hovertemplate='누적: %{y:.1f}%<extra></extra>'
            ), secondary_y=True)
            fig_lc1.add_hline(y=80, line_dash='dash', line_color='orange', secondary_y=True)
            fig_lc1.update_yaxes(title_text='건수', secondary_y=False)
            fig_lc1.update_yaxes(title_text='누적 %', range=[0,105], secondary_y=True)
            fig_lc1.update_layout(height=380, margin=dict(t=10,b=100,l=10,r=60),
                                  xaxis_tickangle=-45, xaxis_title='차종 / 라인')
            st.plotly_chart(fig_lc1, use_container_width=True)

        with lc2:
            # 라인_차종 정지시간 Pareto
            lc_grp2 = (fdf.groupby('라인_차종')
                       .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'))
                       .reset_index().sort_values('총정지시간', ascending=False).head(top_n))
            lc_grp2['누적%'] = (lc_grp2['총정지시간'].cumsum() / lc_grp2['총정지시간'].sum() * 100).round(1)
            lc_grp2['평균MTTR'] = (lc_grp2['총정지시간'] / lc_grp2['건수']).round(1)
            fig_lc2 = make_subplots(specs=[[{"secondary_y": True}]])
            fig_lc2.add_trace(go.Bar(
                x=lc_grp2['라인_차종'], y=lc_grp2['총정지시간'], name='정지시간(분)',
                marker_color='#c0392b',
                customdata=np.stack([lc_grp2['건수'], lc_grp2['평균MTTR'], lc_grp2['누적%']], axis=-1),
                hovertemplate='<b>%{x}</b><br>총정지: %{y:,.0f}분<br>건수: %{customdata[0]:,}건<br>평균MTTR: %{customdata[1]:.1f}분<br>누적: %{customdata[2]:.1f}%<extra></extra>'
            ), secondary_y=False)
            fig_lc2.add_trace(go.Scatter(
                x=lc_grp2['라인_차종'], y=lc_grp2['누적%'], name='누적%',
                line=dict(color='#2980b9', width=2), mode='lines+markers',
                hovertemplate='누적: %{y:.1f}%<extra></extra>'
            ), secondary_y=True)
            fig_lc2.add_hline(y=80, line_dash='dash', line_color='orange', secondary_y=True)
            fig_lc2.update_layout(height=380, margin=dict(t=10,b=100,l=10,r=60),
                                  xaxis_tickangle=-45, xaxis_title='차종 / 라인')
            st.plotly_chart(fig_lc2, use_container_width=True)

        st.divider()

        # ══ ② ★ 드릴다운 — 1단계 전체 현황 + 2단계 멀티셀렉트 Pareto ══
        st.markdown("##### 🔍 세부설비 분석 — 차종·라인 선택 후 설비 Pareto")

        # ── 1단계: 차종/라인별 전체 현황 (가로 막대) ──
        st.markdown("###### ① 차종·라인별 전체 고장현황")
        lc_ovr = (fdf.groupby(['라인_차종','설비유형'])
                  .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'))
                  .reset_index())
        lc_ovr['평균MTTR'] = (lc_ovr['총정지시간'] / lc_ovr['건수']).round(1)
        lc_total_order = (fdf.groupby('라인_차종').size()
                          .sort_values(ascending=True).tail(top_n).index.tolist())
        lc_ovr_f = lc_ovr[lc_ovr['라인_차종'].isin(lc_total_order)]

        fig_step1 = px.bar(
            lc_ovr_f, x='건수', y='라인_차종', orientation='h',
            color='설비유형',
            color_discrete_map={'로봇':'#1e3a5f','지그':'#27ae60',
                                '프레스':'#e67e22','실러':'#8e44ad','기타':'#95a5a6'},
            barmode='stack',
            custom_data=['설비유형','총정지시간','평균MTTR'],
            category_orders={'라인_차종': lc_total_order})
        fig_step1.update_traces(
            hovertemplate='<b>%{y}</b><br>설비유형: %{customdata[0]}<br>건수: %{x:,}건<br>총정지: %{customdata[1]:.0f}분<br>평균MTTR: %{customdata[2]:.1f}분<extra></extra>')
        fig_step1.update_layout(
            height=max(420, len(lc_total_order)*28),
            margin=dict(t=10,b=20,l=10,r=10),
            xaxis_title='건수', yaxis_title='차종 / 라인',
            legend=dict(orientation='h', y=1.02, x=0))
        st.plotly_chart(fig_step1, use_container_width=True)

        st.divider()

        # ── 2단계: 멀티셀렉트 → 선택 라인별 세부분류 Pareto ──
        st.markdown("###### ② 차종·라인 선택 → 설비별 세부 Pareto")

        # 기본값: 건수 상위 3개 라인 선택
        all_lc_sorted = (fdf.groupby('라인_차종').size()
                         .sort_values(ascending=False).index.tolist())
        default_lc = all_lc_sorted[:3]

        sel_lc = st.multiselect(
            "🔽 차종·라인 선택 (복수 선택 가능 — 1개: Pareto, 2개+: 비교 차트)",
            options=all_lc_sorted,
            default=default_lc,
            key='t2_drilldown'
        )

        if not sel_lc:
            st.info("차종·라인을 1개 이상 선택하세요.")
        elif len(sel_lc) == 1:
            # ── 1개 선택: 단독 Pareto (누적% 포함) ──
            lc_name = sel_lc[0]
            sub = fdf[fdf['라인_차종'] == lc_name].copy()
            sd_1 = (sub.groupby('세부분류')
                    .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'))
                    .reset_index().sort_values('건수', ascending=False))
            sd_1['누적%'] = (sd_1['건수'].cumsum() / sd_1['건수'].sum() * 100).round(1)
            sd_1['평균MTTR'] = (sd_1['총정지시간'] / sd_1['건수']).round(1)
            sd_1 = sd_1.head(top_n)

            def color_세부(n):
                if '로봇' in str(n): return '#1e3a5f'
                if '지그' in str(n): return '#27ae60'
                return '#95a5a6'
            sd_1['색상'] = sd_1['세부분류'].apply(color_세부)

            fig_p1 = make_subplots(specs=[[{"secondary_y": True}]])
            fig_p1.add_trace(go.Bar(
                x=sd_1['세부분류'], y=sd_1['건수'], name='건수',
                marker_color=sd_1['색상'].tolist(),
                customdata=np.stack([sd_1['총정지시간'], sd_1['평균MTTR'], sd_1['누적%']], axis=-1),
                hovertemplate='<b>%{x}</b><br>건수: %{y:,}건<br>총정지: %{customdata[0]:.0f}분<br>평균MTTR: %{customdata[1]:.1f}분<br>누적: %{customdata[2]:.1f}%<extra></extra>'
            ), secondary_y=False)
            fig_p1.add_trace(go.Scatter(
                x=sd_1['세부분류'], y=sd_1['누적%'], name='누적%',
                line=dict(color='#e74c3c', width=2), mode='lines+markers',
                hovertemplate='누적: %{y:.1f}%<extra></extra>'
            ), secondary_y=True)
            fig_p1.add_hline(y=80, line_dash='dash', line_color='orange',
                             annotation_text='80%', secondary_y=True)
            fig_p1.update_yaxes(title_text='건수', secondary_y=False)
            fig_p1.update_yaxes(title_text='누적 %', range=[0,105], secondary_y=True)
            fig_p1.update_layout(
                title=f'📍 {lc_name} — 세부설비 Pareto',
                height=420, margin=dict(t=40,b=90,l=10,r=60),
                xaxis_tickangle=-45)
            st.plotly_chart(fig_p1, use_container_width=True)

            # 요약 지표
            총건수 = int(sd_1['건수'].sum())
            top3건수 = int(sd_1.head(3)['건수'].sum())
            top3비율 = round(top3건수/총건수*100, 1) if 총건수 else 0
            m1,m2,m3 = st.columns(3)
            m1.metric("총 고장건수", f"{총건수:,}건")
            m2.metric("상위 3개 설비 집중도", f"{top3비율}%")
            m3.metric("평균 MTTR", f"{sd_1['평균MTTR'].mean():.1f}분")

        else:
            # ── 2개+ 선택: 라인 간 세부분류 비교 (Grouped bar) ──
            sub_multi = fdf[fdf['라인_차종'].isin(sel_lc)].copy()
            sd_multi = (sub_multi.groupby(['라인_차종','세부분류'])
                        .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'))
                        .reset_index())
            sd_multi['평균MTTR'] = (sd_multi['총정지시간'] / sd_multi['건수']).round(1)

            # 선택 라인 전체에서 상위 세부분류만
            top_세부_sel = (sub_multi.groupby('세부분류').size()
                           .sort_values(ascending=False).head(top_n).index.tolist())
            sd_multi_f = sd_multi[sd_multi['세부분류'].isin(top_세부_sel)]

            # 선택 라인 수에 따라 높이 조정
            chart_h = 400 if len(sel_lc) <= 3 else min(600, 350 + len(sel_lc)*20)

            fig_multi = px.bar(
                sd_multi_f, x='세부분류', y='건수', color='라인_차종',
                barmode='group',
                custom_data=['라인_차종','총정지시간','평균MTTR'],
                color_discrete_sequence=px.colors.qualitative.Alphabet,
                title=f'📊 선택 {len(sel_lc)}개 라인 세부설비 비교')
            fig_multi.update_traces(
                hovertemplate='<b>%{x}</b><br>차종·라인: %{customdata[0]}<br>건수: %{y:,}건<br>총정지: %{customdata[1]:.0f}분<br>평균MTTR: %{customdata[2]:.1f}분<extra></extra>')
            fig_multi.update_layout(
                height=chart_h,
                margin=dict(t=40,b=110,l=10,r=10),
                xaxis_tickangle=-45,
                xaxis_title='세부분류', yaxis_title='건수',
                legend=dict(orientation='h', y=-0.45, x=0, font=dict(size=10),
                            title='차종·라인'))
            st.plotly_chart(fig_multi, use_container_width=True)

            # 선택 라인 간 요약 비교 테이블
            with st.expander("📋 선택 라인 비교 요약표"):
                summary_rows = []
                for lc in sel_lc:
                    s = fdf[fdf['라인_차종']==lc]
                    총건수_s = len(s)
                    총정지_s = s['소요시간'].sum()
                    평균mttr_s = s['소요시간'].mean()
                    top1_설비 = s.groupby('세부분류').size().idxmax() if 총건수_s else '-'
                    top1_건수 = s.groupby('세부분류').size().max() if 총건수_s else 0
                    summary_rows.append({
                        '차종·라인': lc,
                        '총건수': 총건수_s,
                        '총정지시간(분)': round(총정지_s, 0),
                        '평균MTTR(분)': round(평균mttr_s, 1),
                        '최다고장설비': top1_설비,
                        '최다건수': top1_건수,
                    })
                st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        # ── 고장부위 분포 (선택 라인 연동) ──
        st.markdown("###### 고장부위 분포")
        if sel_lc:
            부위_src = fdf[fdf['라인_차종'].isin(sel_lc)]
        else:
            top_lc12 = (fdf.groupby('라인_차종').size()
                        .sort_values(ascending=False).head(12).index.tolist())
            부위_src = fdf[fdf['라인_차종'].isin(top_lc12)]

        top_부위 = (부위_src.groupby('부위그룹').size()
                   .sort_values(ascending=False).head(12).index.tolist())
        부위_ln2 = (부위_src.groupby(['부위그룹','라인_차종'])
                   .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'))
                   .reset_index())
        부위_ln2['평균MTTR'] = (부위_ln2['총정지시간'] / 부위_ln2['건수']).round(1)
        부위_f2 = 부위_ln2[부위_ln2['부위그룹'].isin(top_부위)]
        fig_bw2 = px.bar(
            부위_f2, x='건수', y='부위그룹', color='라인_차종',
            orientation='h', barmode='stack',
            custom_data=['라인_차종','총정지시간','평균MTTR'],
            color_discrete_sequence=px.colors.qualitative.Alphabet)
        fig_bw2.update_traces(
            hovertemplate='<b>%{y}</b><br>차종·라인: %{customdata[0]}<br>건수: %{x:,}건<br>총정지: %{customdata[1]:.0f}분<br>평균MTTR: %{customdata[2]:.1f}분<extra></extra>')
        fig_bw2.update_layout(
            height=480, margin=dict(t=10,b=10,l=10,r=10),
            yaxis_title='고장부위', xaxis_title='건수',
            legend=dict(orientation='h', y=-0.15, x=0, font=dict(size=10)))
        st.plotly_chart(fig_bw2, use_container_width=True)

        st.divider()

        # ══ ③ ★ 라인·설비별 고장현황 — 기본값: 라인+설비 복합 ══
        st.markdown("##### 라인별 · 설비별 고장현황")
        view_mode = st.radio(
            "보기 방식",
            ["🏭 라인+설비 복합 Top", "📋 라인별 집계", "🔥 라인×설비 히트맵", "📦 설비별 집계"],
            horizontal=True, key='t2v'
        )

        if view_mode == "🏭 라인+설비 복합 Top":
            kgrp = (fdf.groupby('설비_KEY')
                    .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'),
                         라인=('라인','first'), 설비=('고장설비','first'),
                         유형=('설비유형','first'), 세부=('세부분류','first'))
                    .reset_index())
            kgrp['평균MTTR'] = (kgrp['총정지시간'] / kgrp['건수']).round(1)
            kgrp = kgrp.sort_values('건수', ascending=True).tail(top_n)
            row_h = max(480, len(kgrp) * 26)
            fig_k = px.bar(kgrp, x='건수', y='설비_KEY', orientation='h',
                           color='유형',
                           color_discrete_map={'로봇':'#1e3a5f','지그':'#27ae60',
                                               '프레스':'#e67e22','컨베어':'#8e44ad'},
                           custom_data=['라인','설비','유형','세부','총정지시간','평균MTTR'])
            fig_k.update_traces(
                hovertemplate=(
                    '<b>%{y}</b><br>'
                    '라인: %{customdata[0]}<br>'
                    '설비: %{customdata[1]} (%{customdata[2]})<br>'
                    '세부분류: %{customdata[3]}<br>'
                    '건수: %{x:,}건<br>'
                    '총정지: %{customdata[4]:.0f}분<br>'
                    '평균MTTR: %{customdata[5]:.1f}분'
                    '<extra></extra>'))
            fig_k.update_layout(height=row_h, margin=dict(t=10,b=20,l=10,r=20),
                                yaxis_title='차종 / 라인 | 설비', xaxis_title='건수')
            st.plotly_chart(fig_k, use_container_width=True)

        elif view_mode == "📋 라인별 집계":
            lgrp = (fdf.groupby('라인_차종')
                    .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'),
                         라인=('라인','first'))
                    .reset_index())
            lgrp['평균MTTR'] = (lgrp['총정지시간'] / lgrp['건수']).round(1)
            lgrp = lgrp.sort_values('건수', ascending=True).tail(top_n)
            row_h = max(420, len(lgrp) * 26)
            # 라인별로 색상 구분
            fig_l = px.bar(lgrp, x='건수', y='라인_차종', orientation='h',
                           color='라인',
                           custom_data=['라인','총정지시간','평균MTTR'])
            fig_l.update_traces(
                hovertemplate='<b>%{y}</b><br>라인: %{customdata[0]}<br>건수: %{x:,}건<br>총정지: %{customdata[1]:.0f}분<br>평균MTTR: %{customdata[2]:.1f}분<extra></extra>')
            fig_l.update_layout(height=row_h, margin=dict(t=10,b=20,l=10,r=20),
                                xaxis_title='건수', yaxis_title='차종 / 라인')
            st.plotly_chart(fig_l, use_container_width=True)

        elif view_mode == "🔥 라인×설비 히트맵":
            # 히트맵 표시 범위 설정
            ht_c1, ht_c2, ht_c3 = st.columns(3)
            with ht_c1:
                ht_top_line = st.slider("표시 라인 수", 5, 52, 25, key='ht_l')
            with ht_c2:
                ht_top_col  = st.slider("표시 세부분류 수", 5, 30, 20, key='ht_c')
            with ht_c3:
                ht_sort     = st.selectbox("정렬 기준", ["건수 합계 ↓","건수 합계 ↑"], key='ht_s')

            pivot_data = (fdf.groupby(['라인_차종','세부분류'])
                          .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'))
                          .reset_index())
            pivot   = pivot_data.pivot(index='라인_차종', columns='세부분류', values='건수').fillna(0)
            pivot_t = pivot_data.pivot(index='라인_차종', columns='세부분류', values='총정지시간').fillna(0)

            asc = (ht_sort == "건수 합계 ↑")
            top_lines = pivot.sum(axis=1).sort_values(ascending=asc).head(ht_top_line).index
            top_cols  = pivot.sum(axis=0).sort_values(ascending=False).head(ht_top_col).index
            pivot_f   = pivot.loc[top_lines, [c for c in top_cols if c in pivot.columns]]
            pivot_f   = pivot_f.loc[pivot_f.sum(axis=1).sort_values(ascending=False).index]

            n_rows = len(pivot_f)
            n_cols = len(pivot_f.columns)
            ht    = max(650, n_rows * 32 + 220)

            fig_h = px.imshow(
                pivot_f.astype(int),
                color_continuous_scale='YlOrRd',
                text_auto=True,
                labels=dict(x='세부분류(설비)', y='차종/라인', color='건수'),
                title=f'차종/라인 × 세부분류 고장건수 히트맵  ({n_rows}개 라인·차종 × {n_cols}개 설비)')
            fig_h.update_traces(
                hovertemplate=(
                    '<b>%{y}</b><br>'
                    '설비: %{x}<br>'
                    '건수: %{z:,}건<extra></extra>'))
            fig_h.update_layout(
                height=ht,
                margin=dict(t=55, b=140, l=160, r=20),
                xaxis=dict(tickangle=-45, tickfont=dict(size=11), side='bottom'),
                yaxis=dict(tickfont=dict(size=11)),
                coloraxis_colorbar=dict(title='건수', thickness=14, len=0.7))
            st.plotly_chart(fig_h, use_container_width=True)

            with st.expander("📋 차종·라인별 세부분류 집계 테이블"):
                summary = pivot_f.copy()
                summary.insert(0, '합계', summary.sum(axis=1))
                summary = summary.sort_values('합계', ascending=False)
                st.dataframe(summary.astype(int), use_container_width=True)

        else:  # 설비별 집계
            egrp = (fdf.groupby(['설비유형','고장설비'])
                    .agg(건수=('소요시간','count'), 총정지시간=('소요시간','sum'))
                    .reset_index())
            egrp['평균MTTR'] = (egrp['총정지시간'] / egrp['건수']).round(1)
            egrp = egrp.sort_values('건수', ascending=True).tail(top_n)
            row_h = max(420, len(egrp) * 24)
            fig_e = px.bar(egrp, x='건수', y='고장설비', orientation='h',
                           color='설비유형',
                           color_discrete_map={'로봇':'#1e3a5f','지그':'#27ae60',
                                               '프레스':'#e67e22','컨베어':'#8e44ad'},
                           custom_data=['설비유형','총정지시간','평균MTTR'])
            fig_e.update_traces(
                hovertemplate='<b>%{y}</b> (%{customdata[0]})<br>건수: %{x:,}건<br>총정지: %{customdata[1]:.0f}분<br>평균MTTR: %{customdata[2]:.1f}분<extra></extra>')
            fig_e.update_layout(height=row_h, margin=dict(t=10,b=20,l=10,r=20),
                                xaxis_title='건수', yaxis_title='설비명')
            st.plotly_chart(fig_e, use_container_width=True)

        st.divider()

        # ══ ④ 고장분류 파이 + 조치유형 ══
        p3, p4 = st.columns(2)
        with p3:
            st.markdown("##### 고장분류별 비율")
            fgrp = fdf.groupby('고장분류').agg(
                건수=('소요시간','count'), 총정지시간=('소요시간','sum')
            ).reset_index()
            fgrp['평균MTTR'] = (fgrp['총정지시간'] / fgrp['건수']).round(1)
            fig_pie = px.pie(fgrp, values='건수', names='고장분류',
                             color_discrete_sequence=px.colors.qualitative.Set2, hole=0.4,
                             custom_data=['총정지시간','평균MTTR'])
            fig_pie.update_traces(
                hovertemplate='<b>%{label}</b><br>건수: %{value:,}건 (%{percent})<br>총정지: %{customdata[0][0]:.0f}분<br>평균MTTR: %{customdata[0][1]:.1f}분<extra></extra>')
            fig_pie.update_layout(height=320, margin=dict(t=10,b=10))
            st.plotly_chart(fig_pie, use_container_width=True)

        with p4:
            st.markdown("##### 조치유형 분포")
            atype = fdf.groupby('조치유형').agg(
                건수=('소요시간','count'), 총정지시간=('소요시간','sum')
            ).reset_index().sort_values('건수', ascending=False)
            atype['평균MTTR'] = (atype['총정지시간'] / atype['건수']).round(1)
            fig_at = px.bar(atype, x='조치유형', y='건수',
                            color='조치유형',
                            color_discrete_sequence=px.colors.qualitative.Pastel,
                            custom_data=['총정지시간','평균MTTR'])
            fig_at.update_traces(
                hovertemplate='<b>%{x}</b><br>건수: %{y:,}건<br>총정지: %{customdata[0]:.0f}분<br>평균MTTR: %{customdata[1]:.1f}분<extra></extra>')
            fig_at.update_layout(height=320, margin=dict(t=10,b=20), showlegend=False)
            st.plotly_chart(fig_at, use_container_width=True)

        # ══ ⑤ 월별 트렌드 (세부분류 옵션) ══
        st.markdown("##### 월별 고장건수 트렌드")
        trend_by = st.radio("트렌드 기준", ["설비유형", "세부분류(상위10)"], horizontal=True, key='t2tb')
        fdf2 = fdf.copy()
        fdf2['년월'] = fdf2['발생일시'].dt.to_period('M').astype(str)

        if trend_by == "설비유형":
            mgrp = fdf2.groupby(['년월','설비유형']).agg(
                건수=('소요시간','count'), 총정지시간=('소요시간','sum')
            ).reset_index()
            mgrp['평균MTTR'] = (mgrp['총정지시간'] / mgrp['건수']).round(1)
            fig_tr = px.line(mgrp, x='년월', y='건수', color='설비유형', markers=True,
                             custom_data=['총정지시간','평균MTTR','설비유형'])
            fig_tr.update_traces(
                hovertemplate='<b>%{customdata[2]}</b> — %{x}<br>건수: %{y:,}건<br>총정지: %{customdata[0]:.0f}분<br>평균MTTR: %{customdata[1]:.1f}분<extra></extra>')
        else:
            top10세부 = (fdf2.groupby('세부분류').size()
                        .sort_values(ascending=False).head(10).index.tolist())
            fdf2_f = fdf2[fdf2['세부분류'].isin(top10세부)]
            mgrp = fdf2_f.groupby(['년월','세부분류']).agg(
                건수=('소요시간','count'), 총정지시간=('소요시간','sum')
            ).reset_index()
            mgrp['평균MTTR'] = (mgrp['총정지시간'] / mgrp['건수']).round(1)
            fig_tr = px.line(mgrp, x='년월', y='건수', color='세부분류', markers=True,
                             custom_data=['총정지시간','평균MTTR','세부분류'])
            fig_tr.update_traces(
                hovertemplate='<b>%{customdata[2]}</b> — %{x}<br>건수: %{y:,}건<br>총정지: %{customdata[0]:.0f}분<br>평균MTTR: %{customdata[1]:.1f}분<extra></extra>')

        fig_tr.update_layout(height=360, margin=dict(t=10,b=50,l=10,r=10),
                             xaxis_title='년월', yaxis_title='고장건수',
                             xaxis_tickangle=-30)
        st.plotly_chart(fig_tr, use_container_width=True)

# ══════════════════════════════════════════════════════
# TAB 3 — 설비분석 (MTTR/MTBF)
# ══════════════════════════════════════════════════════
with tab3:
    df = st.session_state.merged_df
    if df is None:
        st.info("Tab1에서 데이터를 불러오고 '통합 실행'을 눌러주세요.")
    else:
        st.subheader("설비분석 — MTTR / MTBF (라인+설비 기준)")

        mc1, mc2, mc3, mc4 = st.columns(4)
        with mc1:
            yrs3 = sorted(df['년'].dropna().unique().astype(int))
            sel_y3 = st.multiselect("연도", yrs3, default=yrs3, key='t3y')
        with mc2:
            min_cnt = st.slider("최소 발생건수", 1, 20, 3, key='t3c')
        with mc3:
            lines3 = ['전체'] + sorted(df['라인'].dropna().unique().tolist(), key=str)
            sel_l3 = st.selectbox("라인 필터", lines3, key='t3l')
        with mc4:
            if '차종' in df.columns:
                cars3 = ['전체'] + sorted(df['차종'].dropna().unique().tolist(), key=str)
            else:
                cars3 = ['전체']
            sel_c3 = st.selectbox("차종 필터", cars3, key='t3car')

        mdf = df[df['년'].isin(sel_y3)].copy() if sel_y3 else df.copy()
        if sel_l3 != '전체': mdf = mdf[mdf['라인'] == sel_l3]
        if sel_c3 != '전체' and '차종' in mdf.columns:
            mdf = mdf[mdf['차종'] == sel_c3]

        mttr_df = calc_mttr_mtbf(mdf)
        mttr_df = mttr_df[mttr_df['발생건수'] >= min_cnt]
        st.markdown(f"**분석 대상 설비 (라인+설비 조합): {len(mttr_df):,}개**")

        # ── MTTR 상위 (라인 색상 구분) ──
        st.markdown("##### 평균 복구시간(MTTR) 상위 — 라인별 색상 구분")
        top_mttr = mttr_df.nlargest(20, 'MTTR(분)')
        fig_mttr = px.bar(
            top_mttr, x='MTTR(분)', y='설비_KEY', orientation='h',
            color='라인',
            custom_data=['라인','고장설비','발생건수','MTBF(시간)','총정지시간(분)','설비유형'])
        fig_mttr.update_traces(
            texttemplate='%{x:.0f}분', textposition='outside',
            hovertemplate=(
                '<b>%{y}</b><br>'
                '라인: %{customdata[0]}<br>'
                '설비: %{customdata[1]}<br>'
                '유형: %{customdata[5]}<br>'
                'MTTR: %{x:.1f}분<br>'
                '발생건수: %{customdata[2]}건<br>'
                'MTBF: %{customdata[3]}시간<br>'
                '총정지: %{customdata[4]:.0f}분'
                '<extra></extra>'))
        fig_mttr.update_layout(
            height=max(420, len(top_mttr)*28),
            margin=dict(t=30,b=20),
            yaxis_title='차종 / 라인 | 설비',
            xaxis_title='MTTR (분)')
        st.plotly_chart(fig_mttr, use_container_width=True)

        # ── MTBF 하위 ──
        mtbf_data = mttr_df[mttr_df['MTBF(시간)'].notna()].nsmallest(20, 'MTBF(시간)')
        if not mtbf_data.empty:
            st.markdown("##### 평균 고장간격(MTBF) 하위 — 짧을수록 위험")
            fig_mtbf = px.bar(
                mtbf_data, x='MTBF(시간)', y='설비_KEY', orientation='h',
                color='라인',
                custom_data=['라인','고장설비','발생건수','MTTR(분)','총정지시간(분)','설비유형'])
            fig_mtbf.update_traces(
                texttemplate='%{x:.0f}h', textposition='outside',
                hovertemplate=(
                    '<b>%{y}</b><br>'
                    '라인: %{customdata[0]}<br>'
                    '설비: %{customdata[1]}<br>'
                    '유형: %{customdata[5]}<br>'
                    'MTBF: %{x:.1f}시간<br>'
                    '발생건수: %{customdata[2]}건<br>'
                    'MTTR: %{customdata[3]:.1f}분<br>'
                    '총정지: %{customdata[4]:.0f}분'
                    '<extra></extra>'))
            fig_mtbf.update_layout(height=max(420, len(mtbf_data)*28),
                                   margin=dict(t=30,b=20), yaxis_title='차종 / 라인 | 설비')
            st.plotly_chart(fig_mtbf, use_container_width=True)

        # ── 산점도: MTTR vs MTBF (라인별 버블) ──
        st.markdown("##### MTTR vs MTBF 산점도 — 우상단이 위험 설비")
        scatter_d = mttr_df[mttr_df['MTBF(시간)'].notna()].copy()
        if not scatter_d.empty:
            fig_sc = px.scatter(
                scatter_d, x='MTBF(시간)', y='MTTR(분)',
                color='라인', size='발생건수',
                hover_name='설비_KEY',
                custom_data=['라인','고장설비','발생건수','총정지시간(분)','설비유형'],
                labels={'MTBF(시간)':'MTBF (시간, 길수록 안전)', 'MTTR(분)':'MTTR (분, 낮을수록 좋음)'})
            fig_sc.update_traces(
                hovertemplate=(
                    '<b>%{hovertext}</b><br>'
                    '라인: %{customdata[0]}<br>'
                    '설비: %{customdata[1]}<br>'
                    '유형: %{customdata[4]}<br>'
                    'MTBF: %{x:.1f}시간<br>'
                    'MTTR: %{y:.1f}분<br>'
                    '발생건수: %{customdata[2]}건<br>'
                    '총정지: %{customdata[3]:.0f}분'
                    '<extra></extra>'))
            # 위험 구간 표시
            mtbf_med = scatter_d['MTBF(시간)'].median()
            mttr_med = scatter_d['MTTR(분)'].median()
            fig_sc.add_vline(x=mtbf_med, line_dash='dash', line_color='gray',
                             annotation_text=f'MTBF 중앙값 {mtbf_med:.0f}h')
            fig_sc.add_hline(y=mttr_med, line_dash='dash', line_color='gray',
                             annotation_text=f'MTTR 중앙값 {mttr_med:.0f}분')
            fig_sc.update_layout(height=460, margin=dict(t=30,b=20))
            st.plotly_chart(fig_sc, use_container_width=True)

        # ── 응답시간 ──
        st.markdown("##### ⏱ 응답시간 분포 (정지→출동)")
        resp = calc_response_time(mdf)
        if resp is not None and not resp.empty:
            rc1, rc2 = st.columns(2)
            with rc1:
                fig_rh = px.histogram(resp, x='응답시간_분', nbins=30,
                                      color_discrete_sequence=['#1e3a5f'])
                fig_rh.add_vline(x=resp['응답시간_분'].mean(), line_dash='dash',
                                 line_color='red',
                                 annotation_text=f"평균 {resp['응답시간_분'].mean():.0f}분")
                fig_rh.update_traces(
                    hovertemplate='응답시간: %{x:.0f}분<br>건수: %{y}건<extra></extra>')
                fig_rh.update_layout(height=300, margin=dict(t=20,b=20),
                                     xaxis_title='응답시간(분)', yaxis_title='건수')
                st.plotly_chart(fig_rh, use_container_width=True)
            with rc2:
                fig_rb = px.box(resp, x='설비유형', y='응답시간_분',
                                color='설비유형')
                fig_rb.update_traces(
                    hovertemplate='설비유형: %{x}<br>응답시간: %{y:.0f}분<extra></extra>')
                fig_rb.update_layout(height=300, margin=dict(t=20,b=20), showlegend=False)
                st.plotly_chart(fig_rb, use_container_width=True)
        else:
            st.info("응답시간 계산 데이터 부족 (정지시각 입력 필요)")

        # ── 반복고장 패턴 ──
        st.markdown("##### 🔁 반복고장 패턴 (5회 이상, 라인+설비 기준)")
        repeat = mttr_df[mttr_df['발생건수'] >= 5].sort_values('발생건수', ascending=False)
        if not repeat.empty:
            st.dataframe(repeat[['라인','고장설비','설비_KEY','발생건수',
                                  'MTTR(분)','MTBF(시간)','총정지시간(분)','설비유형']],
                         use_container_width=True)
        else:
            st.info("5회 이상 반복 고장 없음")

        with st.expander("📋 전체 MTTR/MTBF 테이블"):
            st.dataframe(mttr_df, use_container_width=True)

# ══════════════════════════════════════════════════════
# TAB 4 — 인원분석
# ══════════════════════════════════════════════════════
with tab4:
    df = st.session_state.merged_df
    if df is None:
        st.info("Tab1에서 데이터를 불러오고 '통합 실행'을 눌러주세요.")
    else:
        st.subheader("인원 업무부하 분석")

        yrs4 = sorted(df['년'].dropna().unique().astype(int))
        sel_y4 = st.multiselect("연도 선택", yrs4, default=yrs4, key='t4y')
        wdf_base = df[df['년'].isin(sel_y4)].copy() if sel_y4 else df.copy()

        worker_df = get_worker_df(wdf_base)
        if worker_df.empty:
            st.warning("조치자 데이터가 없습니다.")
            st.stop()

        # 인원별 집계
        person_agg = (worker_df.groupby('조치자')
                      .agg(출동건수=('소요시간','count'),
                           총소요시간_분=('소요시간','sum'))
                      .reset_index())
        person_agg['평균소요시간_분'] = (person_agg['총소요시간_분'] / person_agg['출동건수']).round(1)
        person_agg['총소요시간_시'] = (person_agg['총소요시간_분'] / 60).round(1)
        person_agg = person_agg[person_agg['출동건수'] >= 1].sort_values('출동건수', ascending=False)

        if person_agg.empty:
            st.warning("5건 이상 출동 인원이 없습니다.")
        else:
            pa1, pa2 = st.columns(2)
            with pa1:
                st.markdown("##### 인원별 출동건수")
                fig_p1 = px.bar(person_agg.head(20), x='조치자', y='출동건수',
                                color='출동건수', color_continuous_scale='Blues',
                                custom_data=['총소요시간_시','평균소요시간_분'])
                fig_p1.update_traces(
                    texttemplate='%{y}', textposition='outside',
                    hovertemplate='<b>%{x}</b><br>출동건수: %{y}건<br>총소요시간: %{customdata[0]:.1f}시간<br>평균소요: %{customdata[1]:.1f}분<extra></extra>')
                fig_p1.update_layout(height=400, margin=dict(t=20,b=60),
                                     showlegend=False, xaxis_tickangle=-30)
                st.plotly_chart(fig_p1, use_container_width=True)

            with pa2:
                st.markdown("##### 인원별 총 소요시간")
                fig_p2 = px.bar(person_agg.head(20), x='조치자', y='총소요시간_시',
                                color='총소요시간_시', color_continuous_scale='Reds',
                                custom_data=['출동건수','평균소요시간_분'])
                fig_p2.update_traces(
                    texttemplate='%{y:.0f}h', textposition='outside',
                    hovertemplate='<b>%{x}</b><br>총소요: %{y:.1f}시간<br>출동건수: %{customdata[0]}건<br>평균소요: %{customdata[1]:.1f}분<extra></extra>')
                fig_p2.update_layout(height=400, margin=dict(t=20,b=60),
                                     showlegend=False, xaxis_tickangle=-30)
                st.plotly_chart(fig_p2, use_container_width=True)

            # ── 주별 업무시간 ──
            st.markdown("##### ⚠️ 주별 업무시간 분석 (근로기준법 52시간 기준)")
            wdf2 = worker_df[worker_df['발생일시'].notna()].copy()
            wdf2['연도'] = wdf2['발생일시'].dt.year
            wdf2['주차'] = wdf2['발생일시'].dt.isocalendar().week.astype(int)
            weekly = (wdf2.groupby(['조치자','연도','주차'])['소요시간']
                      .sum().reset_index())
            weekly['소요시간_시'] = (weekly['소요시간'] / 60).round(1)
            weekly['초과위험'] = weekly['소요시간_시'] > 20

            over_workers = weekly[weekly['초과위험']]['조치자'].unique()
            if len(over_workers) > 0:
                st.markdown(
                    f'<div class="warn-box">⚠️ 주간 보전업무 20시간 초과 이력: '
                    f'<b>{", ".join(over_workers[:10])}</b> ({len(over_workers)}명) '
                    f'— 실제 근무시간 포함 시 52시간 초과 위험</div>',
                    unsafe_allow_html=True)

            top_workers = person_agg['조치자'].tolist()
            sel_w = st.selectbox("인원 선택 (주별 차트)", top_workers, key='t4w')
            w_data = weekly[weekly['조치자'] == sel_w].copy()
            w_data['년주'] = (w_data['연도'].astype(str) + '-W' +
                             w_data['주차'].astype(str).str.zfill(2))
            fig_wk = px.bar(w_data, x='년주', y='소요시간_시',
                            color='초과위험',
                            color_discrete_map={True:'#dc3545', False:'#1e3a5f'},
                            custom_data=['연도','주차','소요시간'])
            fig_wk.update_traces(
                hovertemplate='<b>%{x}</b><br>소요시간: %{y:.1f}시간<br>(%{customdata[2]:.0f}분)<extra></extra>')
            fig_wk.add_hline(y=20, line_dash='dash', line_color='orange',
                             annotation_text='경고 20h/주')
            fig_wk.update_layout(height=340, margin=dict(t=30,b=50),
                                  xaxis_tickangle=-45, showlegend=False,
                                  title=f'{sel_w} — 주별 보전업무 소요시간')
            st.plotly_chart(fig_wk, use_container_width=True)

            # ── 인원별 담당 라인/설비 분포 ──
            st.markdown("##### 인원별 담당 라인 분포")
            sel_w2 = st.selectbox("인원 선택 (라인 분포)", top_workers, key='t4w2')
            w_line = worker_df[worker_df['조치자'] == sel_w2]
            if not w_line.empty:
                lc1, lc2 = st.columns(2)
                with lc1:
                    ld = w_line.groupby('라인').size().reset_index(name='건수').sort_values('건수',ascending=False)
                    fig_ld = px.pie(ld, values='건수', names='라인',
                                   title=f'{sel_w2} — 라인별 출동비율', hole=0.35)
                    fig_ld.update_traces(
                        hovertemplate='<b>%{label}</b><br>%{value}건 (%{percent})<extra></extra>')
                    fig_ld.update_layout(height=320, margin=dict(t=40,b=10))
                    st.plotly_chart(fig_ld, use_container_width=True)
                with lc2:
                    ed = w_line.groupby('설비유형').size().reset_index(name='건수').sort_values('건수',ascending=False)
                    fig_ed = px.pie(ed, values='건수', names='설비유형',
                                   title=f'{sel_w2} — 설비유형별 출동비율',
                                   color_discrete_sequence=px.colors.qualitative.Set3, hole=0.35)
                    fig_ed.update_traces(
                        hovertemplate='<b>%{label}</b><br>%{value}건 (%{percent})<extra></extra>')
                    fig_ed.update_layout(height=320, margin=dict(t=40,b=10))
                    st.plotly_chart(fig_ed, use_container_width=True)

            # ── 시간대별 히트맵 ──
            st.markdown("##### 🗓 시간대별 출동 히트맵")
            wdf3 = worker_df[worker_df['발생일시'].notna()].copy()
            wdf3['요일_en'] = wdf3['발생일시'].dt.day_name()
            wdf3['시간'] = wdf3['발생일시'].dt.hour
            heat = wdf3.groupby(['요일_en','시간']).size().reset_index(name='건수')
            day_order = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
            day_kr = {'Monday':'월','Tuesday':'화','Wednesday':'수',
                      'Thursday':'목','Friday':'금','Saturday':'토','Sunday':'일'}
            heat['요일'] = pd.Categorical(
                heat['요일_en'].map(day_kr),
                categories=[day_kr[d] for d in day_order], ordered=True)
            heat = heat.sort_values('요일')
            pivot_h = heat.pivot(index='요일', columns='시간', values='건수').fillna(0)
            fig_heat = px.imshow(pivot_h, color_continuous_scale='YlOrRd',
                                 text_auto=True,
                                 labels=dict(x='시간(시)', y='요일', color='건수'),
                                 title='요일 × 시간대 출동건수')
            fig_heat.update_traces(
                hovertemplate='요일: %{y}<br>%{x}시<br>건수: %{z}건<extra></extra>')
            fig_heat.update_layout(height=320, margin=dict(t=40,b=20))
            st.plotly_chart(fig_heat, use_container_width=True)

            with st.expander("📋 인원별 상세 집계"):
                st.dataframe(person_agg, use_container_width=True)

# ══════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════
# TAB 5 — 설비 위험도 Ranking
# ══════════════════════════════════════════════════════
with tab5:
    df = st.session_state.merged_df
    if df is None:
        st.info("Tab1에서 데이터를 불러오고 '통합 실행'을 눌러주세요.")
    else:
        st.markdown("## 🏆 설비 위험도 Ranking")
        st.caption("위험도 = 고장빈도 × 가중치 + 총정지시간 × 가중치 + 평균MTTR × 가중치")

        # ── 필터 ──
        rf1, rf2, rf3 = st.columns([2, 2, 2])
        with rf1:
            yrs_r = sorted(df['년'].dropna().unique().astype(int))
            sel_yr_r = st.multiselect("연도", yrs_r, default=yrs_r, key='r_yr')
        with rf2:
            eq_r = ['전체'] + sorted(df['설비유형'].dropna().unique().tolist(), key=str)
            sel_eq_r = st.selectbox("설비유형", eq_r, key='r_eq')
        with rf3:
            top_n_r = st.slider("Top N 표시", 10, 50, 20, key='r_top')

        rw1, rw2, rw3 = st.columns(3)
        with rw1:
            w_freq = st.slider("빈도 가중치 (%)", 0, 100, 40, 5, key='r_wf')
        with rw2:
            w_time = st.slider("정지시간 가중치 (%)", 0, 100, 40, 5, key='r_wt')
        with rw3:
            w_mttr = st.slider("MTTR 가중치 (%)", 0, 100, 20, 5, key='r_wm')

        rdf = df[df['년'].isin(sel_yr_r)].copy() if sel_yr_r else df.copy()
        if sel_eq_r != '전체':
            rdf = rdf[rdf['설비유형'] == sel_eq_r]

        # ── 위험도 계산 ──
        risk = (rdf.groupby(['라인_차종', '고장설비', '설비유형'])
                .agg(건수=('소요시간', 'count'),
                     총정지시간=('소요시간', 'sum'),
                     평균MTTR=('소요시간', 'mean'))
                .reset_index())
        risk = risk[risk['건수'] >= 2].copy()
        risk['총정지시간'] = risk['총정지시간'].fillna(0)
        risk['평균MTTR'] = risk['평균MTTR'].fillna(0)

        mx_f = risk['건수'].max() or 1
        mx_t = risk['총정지시간'].max() or 1
        mx_m = risk['평균MTTR'].max() or 1

        risk['위험도점수'] = (
            risk['건수'] / mx_f * w_freq +
            risk['총정지시간'] / mx_t * w_time +
            risk['평균MTTR'] / mx_m * w_mttr
        ).round(1)

        p80 = risk['위험도점수'].quantile(0.80)
        p60 = risk['위험도점수'].quantile(0.60)

        def risk_grade(s):
            if s >= p80: return '🔴 위험'
            if s >= p60: return '🟠 주의'
            return '🟢 양호'

        risk['등급'] = risk['위험도점수'].apply(risk_grade)
        risk = risk.sort_values('위험도점수', ascending=False).reset_index(drop=True)
        risk.index += 1
        risk_top = risk.head(top_n_r).copy()

        # ── KPI ──
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("분석 설비 수", f"{len(risk):,}개")
        k2.metric("🔴 위험", f"{(risk['등급']=='🔴 위험').sum()}개")
        k3.metric("🟠 주의", f"{(risk['등급']=='🟠 주의').sum()}개")
        k4.metric("총 정지시간", f"{risk['총정지시간'].sum():,.0f}분")

        st.divider()

        # ── 위험도 가로 막대 ──
        st.markdown(f"##### 위험도 상위 {top_n_r}개 설비")
        risk_top['설비KEY'] = risk_top['라인_차종'].astype(str) + ' | ' + risk_top['고장설비'].astype(str)
        color_map = {'🔴 위험': '#e74c3c', '🟠 주의': '#e67e22', '🟢 양호': '#27ae60'}

        fig_risk = px.bar(
            risk_top.sort_values('위험도점수'),
            x='위험도점수', y='설비KEY', orientation='h',
            color='등급', color_discrete_map=color_map,
            custom_data=['건수', '총정지시간', '평균MTTR', '설비유형'])
        fig_risk.update_traces(
            hovertemplate=(
                '<b>%{y}</b><br>'
                '위험도: %{x:.1f}점<br>'
                '설비유형: %{customdata[3]}<br>'
                '건수: %{customdata[0]}건<br>'
                '총정지: %{customdata[1]:,.0f}분<br>'
                '평균MTTR: %{customdata[2]:.1f}분<extra></extra>'))
        fig_risk.update_layout(
            height=max(480, top_n_r * 26),
            margin=dict(t=30, b=20, l=10, r=20),
            xaxis_title='위험도 점수', yaxis_title='',
            legend=dict(orientation='h', y=1.05))
        st.plotly_chart(fig_risk, use_container_width=True)

        # ── 버블: 빈도 vs 정지시간 ──
        st.markdown("##### 고장빈도 × 총정지시간 분포 (버블 크기 = 평균MTTR)")
        fig_bub = px.scatter(
            risk_top, x='건수', y='총정지시간',
            size='평균MTTR', color='등급',
            color_discrete_map=color_map,
            text='고장설비',
            hover_data={'라인_차종': True, '위험도점수': True,
                        '건수': True, '총정지시간': True})
        fig_bub.update_traces(textposition='top center', textfont_size=9)
        fig_bub.update_layout(height=420, margin=dict(t=20, b=20, l=10, r=10))
        st.plotly_chart(fig_bub, use_container_width=True)

        # ── 상세표 ──
        with st.expander("📋 위험도 전체 순위표"):
            show_risk = risk[['등급', '위험도점수', '라인_차종', '고장설비',
                               '설비유형', '건수', '총정지시간', '평균MTTR']].head(100).copy()
            show_risk['총정지시간'] = show_risk['총정지시간'].round(0).astype(int)
            show_risk['평균MTTR'] = show_risk['평균MTTR'].round(1)
            st.dataframe(show_risk, use_container_width=True)

        # ── Excel 다운로드 ──
        if st.button("📥 위험도 순위표 Excel 다운로드", key='risk_xl'):
            out = to_excel({'설비위험도': show_risk})
            st.download_button("⬇️ 다운로드", data=out,
                               file_name=f"설비위험도_{datetime.now().strftime('%Y%m%d')}.xlsx",
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               key='risk_dl')


# ══════════════════════════════════════════════════════
# TAB 6 — 유실시간 분석
# ══════════════════════════════════════════════════════
with tab6:
    df = st.session_state.merged_df
    if df is None:
        st.info("Tab1에서 데이터를 불러오고 '통합 실행'을 눌러주세요.")
    else:
        st.markdown("## ⏱️ 유실시간 분석")
        st.caption("고장 정지로 인한 유실시간 집계·트렌드·분포 분석")

        # ── 필터 ──
        lf1, lf2, lf3 = st.columns(3)
        with lf1:
            yrs_l = sorted(df['년'].dropna().unique().astype(int))
            sel_yr_l = st.multiselect("연도", yrs_l, default=yrs_l, key='l_yr')
        with lf2:
            eq_l = ['전체'] + sorted(df['설비유형'].dropna().unique().tolist(), key=str)
            sel_eq_l = st.selectbox("설비유형", eq_l, key='l_eq')
        with lf3:
            lc_l = ['전체'] + sorted(df['라인_차종'].dropna().unique().tolist(), key=str)
            sel_lc_l = st.selectbox("차종·라인", lc_l, key='l_lc')

        ldf = df[df['년'].isin(sel_yr_l)].copy() if sel_yr_l else df.copy()
        if sel_eq_l != '전체':
            ldf = ldf[ldf['설비유형'] == sel_eq_l]
        if sel_lc_l != '전체':
            ldf = ldf[ldf['라인_차종'] == sel_lc_l]
        ldf_v = ldf[ldf['소요시간'].notna()].copy()

        if len(ldf_v) == 0:
            st.warning("필터 조건에 해당하는 소요시간 데이터가 없습니다.")
        else:
            # ── KPI ──
            총유실 = ldf_v['소요시간'].sum()
            총건수 = len(ldf_v)
            평균MTTR = ldf_v['소요시간'].mean()
            최대단건 = ldf_v['소요시간'].max()

            lk1, lk2, lk3, lk4 = st.columns(4)
            lk1.metric("총 유실시간", f"{총유실:,.0f}분", f"({총유실/60:.1f}h)")
            lk2.metric("총 고장건수", f"{총건수:,}건")
            lk3.metric("건당 평균 유실", f"{평균MTTR:.1f}분")
            lk4.metric("단건 최대 유실", f"{최대단건:.0f}분")

            st.divider()

            # ── 월별 트렌드 ──
            st.markdown("##### 월별 유실시간 트렌드 (설비유형별 스택)")
            ldf_v['년월'] = ldf_v['발생일시'].dt.to_period('M').astype(str)
            monthly_eq = (ldf_v.groupby(['년월', '설비유형'])['소요시간']
                          .sum().reset_index(name='유실시간'))
            monthly_cnt = (ldf_v.groupby('년월').size()
                           .reset_index(name='건수').sort_values('년월'))

            fig_tr = make_subplots(specs=[[{"secondary_y": True}]])
            equip_list = ldf_v['설비유형'].dropna().unique()
            colors = px.colors.qualitative.Set2
            for i, eq in enumerate(equip_list):
                sub = monthly_eq[monthly_eq['설비유형'] == eq].sort_values('년월')
                fig_tr.add_trace(go.Bar(
                    x=sub['년월'], y=sub['유실시간'], name=eq,
                    marker_color=colors[i % len(colors)],
                    hovertemplate='%{x}<br>' + eq + ': %{y:.0f}분<extra></extra>'),
                    secondary_y=False)
            fig_tr.add_trace(go.Scatter(
                x=monthly_cnt['년월'], y=monthly_cnt['건수'],
                name='고장건수', mode='lines+markers',
                line=dict(color='#c0392b', width=2, dash='dot'),
                hovertemplate='%{x}<br>건수: %{y}건<extra></extra>'),
                secondary_y=True)
            fig_tr.update_layout(
                barmode='stack', height=400,
                margin=dict(t=10, b=80, l=10, r=60),
                xaxis_tickangle=-45,
                yaxis_title='유실시간(분)',
                legend=dict(orientation='h', y=-0.45, x=0))
            fig_tr.update_yaxes(title_text='고장건수', secondary_y=True)
            st.plotly_chart(fig_tr, use_container_width=True)

            st.divider()

            # ── 설비유형 파이 + 라인별 Top ──
            lc1, lc2 = st.columns(2)
            with lc1:
                st.markdown("##### 설비유형별 유실시간 비중")
                eq_loss = (ldf_v.groupby('설비유형')['소요시간']
                           .sum().reset_index(name='유실시간')
                           .sort_values('유실시간', ascending=False))
                fig_pie = px.pie(
                    eq_loss, values='유실시간', names='설비유형',
                    hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                fig_pie.update_traces(
                    textinfo='label+percent',
                    hovertemplate='<b>%{label}</b><br>%{value:,.0f}분 (%{percent})<extra></extra>')
                fig_pie.update_layout(height=340, margin=dict(t=10, b=10, l=10, r=10),
                                       showlegend=False)
                st.plotly_chart(fig_pie, use_container_width=True)

            with lc2:
                st.markdown("##### 차종·라인별 유실시간 Top 15")
                lc_loss = (ldf_v.groupby('라인_차종')
                           .agg(유실시간=('소요시간', 'sum'),
                                건수=('소요시간', 'count'))
                           .reset_index())
                lc_loss['평균MTTR'] = (lc_loss['유실시간'] / lc_loss['건수']).round(1)
                lc_loss = lc_loss.sort_values('유실시간', ascending=True).tail(15)
                fig_lc = px.bar(
                    lc_loss, x='유실시간', y='라인_차종', orientation='h',
                    color='유실시간', color_continuous_scale='Reds',
                    custom_data=['건수', '평균MTTR'])
                fig_lc.update_traces(
                    hovertemplate='<b>%{y}</b><br>유실: %{x:,.0f}분<br>건수: %{customdata[0]}건<br>평균MTTR: %{customdata[1]}분<extra></extra>')
                fig_lc.update_layout(height=340, margin=dict(t=10, b=10, l=10, r=10),
                                      coloraxis_showscale=False)
                st.plotly_chart(fig_lc, use_container_width=True)

            st.divider()

            # ── 요일별 유실 패턴 ──
            st.markdown("##### 요일별 유실시간 분포")
            요일명 = {0: '월', 1: '화', 2: '수', 3: '목', 4: '금', 5: '토', 6: '일'}
            ldf_v['요일'] = ldf_v['발생일시'].dt.dayofweek.map(요일명)
            dow = (ldf_v.groupby('요일')['소요시간']
                   .agg(['sum', 'count', 'mean']).reset_index())
            dow.columns = ['요일', '총유실시간', '건수', '평균유실']
            dow_order = ['월', '화', '수', '목', '금', '토', '일']
            dow['요일'] = pd.Categorical(dow['요일'], categories=dow_order, ordered=True)
            dow = dow.sort_values('요일')
            fig_dow = px.bar(dow, x='요일', y='총유실시간',
                             color='총유실시간', color_continuous_scale='Blues',
                             custom_data=['건수', '평균유실'], text='건수')
            fig_dow.update_traces(
                hovertemplate='<b>%{x}요일</b><br>총유실: %{y:,.0f}분<br>건수: %{customdata[0]}건<br>평균: %{customdata[1]:.1f}분<extra></extra>',
                texttemplate='%{text}건', textposition='outside')
            fig_dow.update_layout(height=320, margin=dict(t=10, b=20, l=10, r=10),
                                   coloraxis_showscale=False, xaxis_title='')
            st.plotly_chart(fig_dow, use_container_width=True)

            # ── 상세표 ──
            with st.expander("📋 설비별 유실시간 상세"):
                detail = (ldf_v.groupby(['라인_차종', '고장설비', '설비유형'])
                          .agg(건수=('소요시간', 'count'),
                               총유실시간_분=('소요시간', 'sum'),
                               평균MTTR_분=('소요시간', 'mean'),
                               최대단건_분=('소요시간', 'max'))
                          .reset_index()
                          .sort_values('총유실시간_분', ascending=False))
                detail['총유실시간_분'] = detail['총유실시간_분'].round(0).astype(int)
                detail['평균MTTR_분'] = detail['평균MTTR_분'].round(1)
                detail['최대단건_분'] = detail['최대단건_분'].round(0).astype(int)
                st.dataframe(detail, use_container_width=True, hide_index=True)

            # ── Excel ──
            if st.button("📥 유실시간 분석 Excel 다운로드", key='loss_xl'):
                out = to_excel({'유실시간_설비별': detail, '유실시간_라인별': lc_loss})
                st.download_button("⬇️ 다운로드", data=out,
                                   file_name=f"유실시간_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                   key='loss_dl')


# ══════════════════════════════════════════════════════
# TAB 7 — 예방정비 추천
# ══════════════════════════════════════════════════════
with tab7:
    df = st.session_state.merged_df
    if df is None:
        st.info("Tab1에서 데이터를 불러오고 '통합 실행'을 눌러주세요.")
    else:
        st.markdown("## 🔧 예방정비 추천")
        st.caption("고장 패턴·반복 주기·정지시간 기반 우선순위 자동 추천")

        # ── 필터 ──
        pf1, pf2, pf3 = st.columns(3)
        with pf1:
            yrs_p = sorted(df['년'].dropna().unique().astype(int))
            def_p = yrs_p[-2:] if len(yrs_p) >= 2 else yrs_p
            sel_yr_p = st.multiselect("연도", yrs_p, default=def_p, key='p_yr')
        with pf2:
            eq_p = ['전체'] + sorted(df['설비유형'].dropna().unique().tolist(), key=str)
            sel_eq_p = st.selectbox("설비유형", eq_p, key='p_eq')
        with pf3:
            min_cnt_p = st.slider("최소 고장건수 (신뢰도 기준)", 2, 20, 5, key='p_min')

        pdf = df[df['년'].isin(sel_yr_p)].copy() if sel_yr_p else df.copy()
        if sel_eq_p != '전체':
            pdf = pdf[pdf['설비유형'] == sel_eq_p]
        pdf_v = pdf[pdf['소요시간'].notna()].copy()

        # ── 설비별 집계 ──
        pm = (pdf_v.groupby(['라인_차종', '고장설비', '설비유형', '고장부위'])
              .agg(건수=('소요시간', 'count'),
                   총정지시간=('소요시간', 'sum'),
                   평균MTTR=('소요시간', 'mean'),
                   최근발생=('발생일시', 'max'),
                   최초발생=('발생일시', 'min'))
              .reset_index())
        pm = pm[pm['건수'] >= min_cnt_p].copy()

        if len(pm) == 0:
            st.warning(f"최소 {min_cnt_p}건 이상 고장 설비가 없습니다. 슬라이더를 낮추거나 연도 범위를 넓혀주세요.")
        else:
            pm['평균MTTR'] = pm['평균MTTR'].round(1)
            pm['총정지시간'] = pm['총정지시간'].round(0)
            pm['분석기간_일'] = (pm['최근발생'] - pm['최초발생']).dt.days + 1
            pm['고장주기_일'] = (pm['분석기간_일'] / pm['건수']).round(1)

            # 우선순위 점수
            mx_c = pm['건수'].max() or 1
            mx_t = pm['총정지시간'].max() or 1
            mx_m = pm['평균MTTR'].max() or 1
            pm['우선순위점수'] = (
                pm['건수'] / mx_c * 40 +
                pm['총정지시간'] / mx_t * 40 +
                pm['평균MTTR'] / mx_m * 20
            ).round(1)

            # 등급
            p80 = pm['우선순위점수'].quantile(0.80)
            p60 = pm['우선순위점수'].quantile(0.60)
            def pm_grade(s):
                if s >= p80: return '🔴 즉시조치'
                if s >= p60: return '🟠 조기예방'
                return '🟢 정기점검'
            pm['우선순위'] = pm['우선순위점수'].apply(pm_grade)

            # 추천 액션
            def get_action(row):
                acts = []
                cyc = row['고장주기_일']
                if pd.notna(cyc):
                    if cyc < 30:
                        acts.append(f"⚡ 점검주기 단축 (평균 {cyc:.0f}일마다 고장)")
                    elif cyc < 90:
                        acts.append(f"📅 월간 집중점검 (평균 {cyc:.0f}일 주기)")
                    else:
                        acts.append(f"📋 분기점검 유지 (평균 {cyc:.0f}일 주기)")
                mttr = row['평균MTTR']
                if mttr >= 60:
                    acts.append(f"🔩 예비부품 사전확보 필수 (수리 평균 {mttr:.0f}분)")
                elif mttr >= 30:
                    acts.append(f"🔩 핵심부품 재고 점검 (수리 {mttr:.0f}분)")
                부위 = row['고장부위']
                if pd.notna(부위) and str(부위) not in ('nan', 'None', ''):
                    acts.append(f"🔍 중점부위: {부위}")
                if row['건수'] >= 10:
                    acts.append(f"📊 근본원인 분석 필요 ({row['건수']}회 반복)")
                return ' | '.join(acts) if acts else '✅ 현행 정기점검 유지'

            pm['추천조치'] = pm.apply(get_action, axis=1)
            pm = pm.sort_values('우선순위점수', ascending=False).reset_index(drop=True)
            pm.index += 1

            # ── KPI ──
            pk1, pk2, pk3, pk4 = st.columns(4)
            pk1.metric("추천 대상 설비", f"{len(pm)}개")
            pk2.metric("🔴 즉시조치", f"{(pm['우선순위']=='🔴 즉시조치').sum()}개")
            pk3.metric("🟠 조기예방", f"{(pm['우선순위']=='🟠 조기예방').sum()}개")
            pk4.metric("평균 고장주기", f"{pm['고장주기_일'].mean():.0f}일")

            st.divider()

            # ── 즉시조치 카드 ──
            urgent = pm[pm['우선순위'] == '🔴 즉시조치'].head(10)
            if len(urgent):
                st.markdown("##### 🔴 즉시조치 대상")
                for idx, row in urgent.iterrows():
                    st.markdown(
                        f'<div style="border-left:5px solid #e74c3c;background:#fff5f5;'
                        f'padding:12px 16px;border-radius:0 8px 8px 0;margin-bottom:8px;">'
                        f'<b>#{idx} {row["라인_차종"]} | {row["고장설비"]}</b>'
                        f' <span style="color:#888;font-size:12px">{row["설비유형"]}</span><br>'
                        f'고장 <b>{row["건수"]}건</b> / 총정지 <b>{row["총정지시간"]:.0f}분</b> / '
                        f'평균MTTR <b>{row["평균MTTR"]}분</b> / 고장주기 <b>{row["고장주기_일"]}일</b><br>'
                        f'<span style="color:#c0392b"><b>📌 추천:</b> {row["추천조치"]}</span>'
                        f'</div>', unsafe_allow_html=True)
                st.divider()

            # ── 조기예방 카드 ──
            early = pm[pm['우선순위'] == '🟠 조기예방'].head(10)
            if len(early):
                st.markdown("##### 🟠 조기예방 대상")
                for idx, row in early.iterrows():
                    st.markdown(
                        f'<div style="border-left:5px solid #e67e22;background:#fff9f0;'
                        f'padding:12px 16px;border-radius:0 8px 8px 0;margin-bottom:8px;">'
                        f'<b>#{idx} {row["라인_차종"]} | {row["고장설비"]}</b>'
                        f' <span style="color:#888;font-size:12px">{row["설비유형"]}</span><br>'
                        f'고장 <b>{row["건수"]}건</b> / 총정지 <b>{row["총정지시간"]:.0f}분</b> / '
                        f'고장주기 <b>{row["고장주기_일"]}일</b><br>'
                        f'<span style="color:#d35400"><b>📌 추천:</b> {row["추천조치"]}</span>'
                        f'</div>', unsafe_allow_html=True)
                st.divider()

            # ── 고장주기 산점도 ──
            st.markdown("##### 고장주기 vs 총정지시간 분포")
            fig_sc = px.scatter(
                pm.head(60), x='고장주기_일', y='총정지시간',
                size='건수', color='우선순위',
                color_discrete_map={
                    '🔴 즉시조치': '#e74c3c',
                    '🟠 조기예방': '#e67e22',
                    '🟢 정기점검': '#27ae60'},
                hover_data={'라인_차종': True, '고장설비': True,
                            '평균MTTR': True, '우선순위점수': True},
                title='버블 크기 = 고장건수')
            fig_sc.add_vline(x=30, line_dash='dash', line_color='#e74c3c',
                             annotation_text='30일')
            fig_sc.add_vline(x=90, line_dash='dash', line_color='#e67e22',
                             annotation_text='90일')
            fig_sc.update_layout(height=420, margin=dict(t=40, b=20, l=10, r=10))
            st.plotly_chart(fig_sc, use_container_width=True)

            # ── 전체 목록 ──
            with st.expander("📋 예방정비 추천 전체 목록"):
                show_pm = pm[['우선순위', '우선순위점수', '라인_차종', '고장설비',
                               '설비유형', '고장부위', '건수', '총정지시간',
                               '평균MTTR', '고장주기_일', '추천조치']].copy()
                st.dataframe(show_pm, use_container_width=True)

            if st.button("📥 예방정비 추천 Excel 다운로드", key='pm_xl'):
                out = to_excel({'예방정비추천': show_pm})
                st.download_button("⬇️ 다운로드", data=out,
                                   file_name=f"예방정비추천_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                   key='pm_dl')


# ══════════════════════════════════════════════════════
# TAB 8 — 출력
# ══════════════════════════════════════════════════════
with tab8:
    df = st.session_state.merged_df
    if df is None:
        st.info("Tab1에서 데이터를 불러오고 '통합 실행'을 눌러주세요.")
    else:
        st.subheader("분석 결과 출력")
        oc1, oc2 = st.columns(2)

        with oc1:
            st.markdown("#### 📊 Excel 분석결과 다운로드")
            if st.button("Excel 파일 생성", use_container_width=True, key='ex1'):
                with st.spinner("생성 중..."):
                    try:
                        mttr_r = calc_mttr_mtbf(df)
                        wdf_all = get_worker_df(df)
                        person_r = (wdf_all.groupby('조치자')
                                    .agg(출동건수=('소요시간', 'count'),
                                         총소요시간_분=('소요시간', 'sum'))
                                    .reset_index().sort_values('출동건수', ascending=False))
                        pareto_r = (df.groupby(['라인', '설비유형'])
                                    .agg(건수=('소요시간', 'count'),
                                         총정지_분=('소요시간', 'sum'))
                                    .reset_index().sort_values('건수', ascending=False))
                        key_r = (df.groupby('설비_KEY')
                                 .agg(건수=('소요시간', 'count'),
                                      총정지_분=('소요시간', 'sum'))
                                 .reset_index().sort_values('건수', ascending=False))
                        export_cols = ['발생일시', '라인', '설비_KEY', '설비유형', '고장설비',
                                       '고장부위', '현상', '원인', '조치내역', '조치자',
                                       '소요시간', '조치유형', '고장분류', '파일출처']
                        excel_data = to_excel({
                            '통합데이터': df[[c for c in export_cols if c in df.columns]],
                            'Pareto_라인별설비': pareto_r,
                            '라인설비_복합키': key_r,
                            'MTTR_MTBF': mttr_r,
                            '인원별_부하': person_r,
                        })
                        st.download_button(
                            "⬇️ Excel 다운로드", data=excel_data,
                            file_name=f"보전팀_분석_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            use_container_width=True)
                    except Exception as e:
                        st.error(f"Excel 생성 오류: {e}")

        with oc2:
            st.markdown("#### 📋 통합 원본 CSV 다운로드")
            if st.button("CSV 생성", use_container_width=True, key='csv1'):
                csv = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    "⬇️ CSV 다운로드", data=csv.encode('utf-8-sig'),
                    file_name=f"보전팀_통합_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime='text/csv', use_container_width=True)

        st.divider()
        st.markdown("#### 🖼 차트 이미지 저장 (PPT 붙여넣기용)")
        st.info("각 차트 우측 상단 📷 아이콘 클릭 → PNG 저장")

        st.divider()
        st.markdown("#### 🖨 PDF 출력 방법")
        st.markdown(
            '<div style="background:#eaf4fb;border-left:4px solid #2471a3;'
            'padding:16px 20px;border-radius:6px;font-size:14px;line-height:2.2;">'
            '<b>분석 화면을 PDF로 저장하는 방법</b><br>'
            '1단계 : 저장하려는 탭으로 이동<br>'
            '2단계 : 키보드 <code>Ctrl + P</code> (Mac: <code>Cmd + P</code>) 입력<br>'
            '3단계 : 프린터 선택란에서 <b>PDF로 저장</b> 선택<br>'
            '4단계 : 용지 → <b>A3 가로</b> 권장 / 배율 → <b>맞춤</b> 설정<br>'
            '5단계 : <b>저장</b> 클릭</div>',
            unsafe_allow_html=True)
