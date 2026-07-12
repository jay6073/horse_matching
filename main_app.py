import io

import pandas as pd
import streamlit as st

from excel_service import MultiDataMatcherService

st.set_page_config(page_title="다중 통합 매칭 시스템 v6.9", layout="wide")

svc = MultiDataMatcherService()


# ----------------------------------------------------------------------
# 유틸: 업로드된 파일(BytesIO 유사 객체)에서 시트 목록 읽기
#       (원본 get_sheets는 경로 문자열 기준이라 업로드 객체에 맞게 보완)
# ----------------------------------------------------------------------
def get_sheets_from_upload(uploaded_file) -> list[str]:
    try:
        from python_calamine import CalamineWorkbook
        uploaded_file.seek(0)
        return CalamineWorkbook.from_filelike(uploaded_file).sheet_names
    except Exception:
        try:
            uploaded_file.seek(0)
            return pd.ExcelFile(uploaded_file).sheet_names
        except Exception:
            return []


def load_and_sanitize(uploaded_file, sheet_name: str) -> pd.DataFrame:
    uploaded_file.seek(0)
    h_idx = svc.find_best_header(uploaded_file, sheet_name)
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, engine="calamine", sheet_name=sheet_name, header=h_idx)
    return svc.sanitize(df), h_idx


# ----------------------------------------------------------------------
# 세션 상태 초기화
# ----------------------------------------------------------------------
if "df_a" not in st.session_state:
    st.session_state.df_a = None
if "df_a_name" not in st.session_state:
    st.session_state.df_a_name = None
if "ref_files" not in st.session_state:
    st.session_state.ref_files = []  # [{'name', 'df', 'keys_b', 'cols'}]
if "ref_uploader_key" not in st.session_state:
    st.session_state.ref_uploader_key = 0

st.title("다중 통합 매칭 시스템 v6.9")
st.caption("자동 헤더 탐지 · 다중 파일 병합 · 불일치 사유 분석")

# ----------------------------------------------------------------------
# 1. 기준 파일 (파일A)
# ----------------------------------------------------------------------
st.header("1. 기준 파일 (파일A) 설정")

file_a = st.file_uploader("기준 파일 업로드", type=["xlsx", "xls"], key="file_a_uploader")

if file_a is not None:
    sheets_a = get_sheets_from_upload(file_a)
    if sheets_a:
        sheet_a = st.selectbox("시트 선택 (파일A)", sheets_a, key="sheet_a_select")
        if st.button("파일A 불러오기", type="primary"):
            df_a, h_idx = load_and_sanitize(file_a, sheet_a)
            st.session_state.df_a = df_a
            st.session_state.df_a_name = file_a.name
            st.session_state.keys_a = []
            st.success(f"'{file_a.name}' 로딩 완료 (헤더 탐지: {h_idx + 1}행)")
    else:
        st.error("시트를 읽을 수 없습니다. 파일 형식을 확인해주세요.")

if st.session_state.df_a is not None:
    st.write(f"현재 로딩된 파일A: **{st.session_state.df_a_name}** ({len(st.session_state.df_a)}행)")
    with st.expander("미리보기", expanded=False):
        st.dataframe(st.session_state.df_a.head(20), use_container_width=True)

    keys_a = st.multiselect(
        "매칭 기준열 선택 (선택한 순서대로 비교 파일의 매칭기준과 1:1 대응됩니다)",
        options=list(st.session_state.df_a.columns),
        default=st.session_state.get("keys_a", []),
        key="keys_a_select",
    )
    st.session_state.keys_a = keys_a

st.divider()

# ----------------------------------------------------------------------
# 2. 비교 파일 추가 및 설정
# ----------------------------------------------------------------------
st.header("2. 비교 파일 추가 및 설정")

new_ref_file = st.file_uploader(
    "비교 파일 추가",
    type=["xlsx", "xls"],
    key=f"ref_uploader_{st.session_state.ref_uploader_key}",
)

if new_ref_file is not None:
    sheets_b = get_sheets_from_upload(new_ref_file)
    if sheets_b:
        sheet_b = st.selectbox("시트 선택 (비교 파일)", sheets_b, key="sheet_b_select")
        if st.button("이 파일 추가(+)"):
            df_b, h_idx = load_and_sanitize(new_ref_file, sheet_b)
            name = new_ref_file.name.rsplit(".", 1)[0]
            st.session_state.ref_files.append(
                {"name": name, "df": df_b, "keys_b": [], "cols": []}
            )
            st.session_state.ref_uploader_key += 1  # 업로더 위젯 초기화(다음 파일을 위해)
            st.success(f"'{name}' 추가 완료 (헤더 탐지: {h_idx + 1}행)")
            st.rerun()
    else:
        st.error("시트를 읽을 수 없습니다. 파일 형식을 확인해주세요.")

if not st.session_state.ref_files:
    st.info("아직 추가된 비교 파일이 없습니다.")

remove_idx = None
for idx, item in enumerate(st.session_state.ref_files):
    ready = "✅" if (item["keys_b"] and item["cols"]) else "⚠️ 설정 필요"
    with st.expander(f"[{idx + 1}] {item['name']}  ({len(item['df'])}행)  {ready}", expanded=not item["keys_b"]):
        cols_all = list(item["df"].columns)

        keys_b = st.multiselect(
            "매칭 기준열 (파일A와 같은 순서로 선택)",
            options=cols_all,
            default=item["keys_b"],
            key=f"keys_b_{idx}",
        )
        item["keys_b"] = keys_b

        cols_selected = st.multiselect(
            "가져올 열 선택",
            options=cols_all,
            default=item["cols"],
            key=f"cols_{idx}",
        )
        item["cols"] = cols_selected

        if st.button("이 파일 삭제", key=f"remove_{idx}"):
            remove_idx = idx

if remove_idx is not None:
    st.session_state.ref_files.pop(remove_idx)
    st.rerun()

st.divider()

# ----------------------------------------------------------------------
# 3. 실행
# ----------------------------------------------------------------------
st.header("3. 다중 매칭 실행 및 결과 다운로드")

keys_a = st.session_state.get("keys_a", [])
df_a = st.session_state.df_a
ref_files = st.session_state.ref_files

can_run = df_a is not None and len(keys_a) > 0 and len(ref_files) > 0 and all(
    item["keys_b"] and item["cols"] for item in ref_files
)

if df_a is None:
    st.warning("먼저 파일A를 불러오세요.")
elif not keys_a:
    st.warning("파일A의 매칭 기준열을 선택하세요.")
elif not ref_files:
    st.warning("비교할 파일을 하나 이상 추가하세요.")
elif not can_run:
    st.warning("모든 비교 파일의 매칭 기준열/가져올 열 설정을 완료하세요.")

if st.button("다중 매칭 시작 및 통합 저장", type="primary", disabled=not can_run):
    with st.spinner("다중 파일 병합 및 사유 분석 중..."):
        try:
            result_df = svc.process_multi_merge(df_a, ref_files, keys_a)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                result_df.to_excel(writer, index=False, sheet_name="통합결과")
            buffer.seek(0)

            st.success(f"완료되었습니다! 총 데이터: {len(result_df):,}건")
            st.dataframe(result_df.head(50), use_container_width=True)

            st.download_button(
                label="결과 파일 다운로드 (.xlsx)",
                data=buffer,
                file_name="다중통합매칭_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"처리 중 오류 발생: {e}")
