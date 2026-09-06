import io
import pandas as pd
import streamlit as st

# 비즈니스 로직 서비스 (동일 유지)
from excel_service import MultiDataMatcherService

st.set_page_config(page_title="다중 통합 매칭 시스템 v7.3", layout="wide")

svc = MultiDataMatcherService()


# ----------------------------------------------------------------------
# 유틸 함수
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


def load_and_sanitize(uploaded_file, sheet_name: str) -> tuple[pd.DataFrame, int]:
    uploaded_file.seek(0)
    h_idx = svc.find_best_header(uploaded_file, sheet_name)
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, engine="calamine", sheet_name=sheet_name, header=h_idx)
    return svc.sanitize(df), h_idx


def ordered_keys_from_editor(edited_df: pd.DataFrame, order_col: str, name_col: str = "컬럼명") -> list[str]:
    """순번(order_col)이 1 이상인 행만 골라 순번 오름차순으로 정렬한 컬럼명 리스트를 반환.
    순번이 중복되면 원래 행 순서를 2차 정렬 기준으로 사용해 안정적으로 처리한다."""
    picked = edited_df[edited_df[order_col].fillna(0) > 0].copy()
    if picked.empty:
        return []
    picked["_orig_idx"] = picked.index
    picked = picked.sort_values(by=[order_col, "_orig_idx"])
    return picked[name_col].tolist()


# ----------------------------------------------------------------------
# 세션 상태 초기화
# ----------------------------------------------------------------------
if "df_a" not in st.session_state:
    st.session_state.df_a = None
if "df_a_name" not in st.session_state:
    st.session_state.df_a_name = None
if "sheet_a_name" not in st.session_state:
    st.session_state.sheet_a_name = None
if "keys_a" not in st.session_state:
    st.session_state.keys_a = []
if "ref_files" not in st.session_state:
    st.session_state.ref_files = []  # [{'name', 'sheet', 'df', 'keys_b', 'cols'}]
if "ref_uploader_key" not in st.session_state:
    st.session_state.ref_uploader_key = 0

# ----------------------------------------------------------------------
# 헤더
# ----------------------------------------------------------------------
st.title("다중 통합 매칭 시스템 v7.3 (UI 개선판)")
st.caption("좌/우 비교 화면 · 체크박스+순번 기반 컬럼 선택 · 자동 헤더 탐지")

# ----------------------------------------------------------------------
# 0. 진행 상태 스테퍼
# ----------------------------------------------------------------------
keys_a_preview = st.session_state.get("keys_a", [])
df_a_preview = st.session_state.df_a
ref_files_preview = st.session_state.ref_files

step1_done = df_a_preview is not None
step2_done = len(keys_a_preview) > 0
step3_done = len(ref_files_preview) > 0
step4_done = (
    step1_done and step2_done and step3_done and
    all(item["keys_b"] and item["cols"] for item in ref_files_preview)
)

step_cols = st.columns(4)
step_labels = [
    ("① 파일A 로딩", step1_done),
    ("② 기준열 선택", step2_done),
    ("③ 비교파일 추가", step3_done),
    ("④ 실행 준비 완료", step4_done),
]
for col, (label, done) in zip(step_cols, step_labels):
    icon = ":material/check_circle:" if done else ":material/radio_button_unchecked:"
    col.markdown(f"**{icon} {label}**")

st.divider()

# ----------------------------------------------------------------------
# 좌/우 2단 레이아웃 구성
# 비교 파일이 늘어날수록 탭 제목이 잘려 보이지 않는 문제를 완화하기 위해,
# 3개째부터는 우측(비교 파일) 컬럼 비중을 점진적으로 늘려준다 (최대 1:2.5).
# ----------------------------------------------------------------------
n_ref = len(st.session_state.ref_files)
right_weight = min(1 + max(n_ref - 2, 0) * 0.3, 2.5) if n_ref >= 3 else 1
col_left, col_right = st.columns([1, right_weight], gap="large")

# ======================================================================
# [좌측 화면] 1. 기준 파일 (파일A) 설정
# ======================================================================
with col_left:
    st.subheader(":material/push_pin: 1. 기준 파일 (파일A) 설정")

    file_a = st.file_uploader("기준 파일 업로드", type=["xlsx", "xls"], key="file_a_uploader")

    if file_a is not None:
        sheets_a = get_sheets_from_upload(file_a)
        if sheets_a:
            sheet_a = st.selectbox("시트 선택 (파일A)", sheets_a, key="sheet_a_select")
            if st.button("파일A 불러오기", type="primary", use_container_width=True):
                df_a, h_idx = load_and_sanitize(file_a, sheet_a)
                st.session_state.df_a = df_a
                st.session_state.df_a_name = file_a.name
                st.session_state.sheet_a_name = sheet_a
                st.session_state.keys_a = []
                st.success(f"'{file_a.name}' ({sheet_a}) 로딩 완료 (헤더: {h_idx + 1}행)")
        else:
            st.error("시트를 읽을 수 없습니다.")

    if st.session_state.df_a is not None:
        st.markdown(
            f"**현재 파일A:** `{st.session_state.df_a_name}` "
            f"(시트: `{st.session_state.sheet_a_name}`, {len(st.session_state.df_a):,}행)"
        )

        # --- 체크박스 + 순번 방식의 기준열 선택 ---
        st.write("▼ **매칭 기준열 선택** (순번: 매칭에 사용할 순서를 1, 2, 3...으로 입력. 미선택은 공란)")
        cols_a = list(st.session_state.df_a.columns)

        prev_keys_a = st.session_state.keys_a
        df_a_config = pd.DataFrame({
            "컬럼명": cols_a,
            "순번": [prev_keys_a.index(c) + 1 if c in prev_keys_a else None for c in cols_a],
        })
        df_a_config["순번"] = df_a_config["순번"].astype("Int64")  # nullable int → 미선택은 공란으로 표시

        edited_a = st.data_editor(
            df_a_config,
            column_config={
                "컬럼명": st.column_config.TextColumn("컬럼명", disabled=True),
                "순번": st.column_config.NumberColumn(
                    "매칭 순번", min_value=1, step=1,
                    help="매칭 기준으로 쓸 순서대로 1, 2, 3...을 입력하세요. 비워두면 미선택입니다.",
                ),
            },
            hide_index=True,
            use_container_width=True,
            key=f"editor_keys_a_{st.session_state.df_a_name}_{st.session_state.sheet_a_name}",
        )

        selected_keys_a = ordered_keys_from_editor(edited_a, "순번")
        st.session_state.keys_a = selected_keys_a

        if selected_keys_a:
            st.info(f"선택된 기준열 (순서): **{' → '.join(selected_keys_a)}**")

        with st.expander("데이터 미리보기 (상위 10행)", expanded=False):
            st.dataframe(st.session_state.df_a.head(10), use_container_width=True)


# ======================================================================
# [우측 화면] 2. 비교 파일 추가 및 설정
# ======================================================================
with col_right:
    st.subheader(":material/folder_open: 2. 비교 파일 추가 및 설정")

    new_ref_file = st.file_uploader(
        "비교 파일 추가",
        type=["xlsx", "xls"],
        key=f"ref_uploader_{st.session_state.ref_uploader_key}",
    )

    if new_ref_file is not None:
        sheets_b = get_sheets_from_upload(new_ref_file)
        if sheets_b:
            sheet_b = st.selectbox("시트 선택 (비교 파일)", sheets_b, key="sheet_b_select")
            if st.button("이 파일 추가 (+)", type="secondary", use_container_width=True):
                df_b, h_idx = load_and_sanitize(new_ref_file, sheet_b)
                name = new_ref_file.name.rsplit(".", 1)[0]
                st.session_state.ref_files.append(
                    {"name": name, "sheet": sheet_b, "df": df_b, "keys_b": [], "cols": list(df_b.columns)}
                )
                st.session_state.ref_uploader_key += 1
                st.success(f"'{name}' ({sheet_b}) 추가 완료")
                st.rerun()
        else:
            st.error("시트를 읽을 수 없습니다.")

    if not st.session_state.ref_files:
        st.info("비교할 파일을 등록해주세요.")

    remove_idx = None

    if st.session_state.ref_files:
        tab_labels = [
            f"{':material/check_circle:' if (item['keys_b'] and item['cols']) else ':material/warning:'} "
            f"[{i + 1}] {item['name']} ({item.get('sheet', '')})"
            for i, item in enumerate(st.session_state.ref_files)
        ]
        tabs = st.tabs(tab_labels)

        for idx, (tab, item) in enumerate(zip(tabs, st.session_state.ref_files)):
            with tab:
                st.caption(f"시트: {item.get('sheet', '-')} · {len(item['df']):,}행")
                cols_all = list(item["df"].columns)

                st.write("▼ **매칭 기준열 (순번)** / **가져올 열 (체크)**")

                prev_keys_b = item["keys_b"]
                ref_config_df = pd.DataFrame({
                    "컬럼명": cols_all,
                    "매칭 순번": [prev_keys_b.index(c) + 1 if c in prev_keys_b else None for c in cols_all],
                    "가져올 열": [c in item["cols"] for c in cols_all],
                })
                ref_config_df["매칭 순번"] = ref_config_df["매칭 순번"].astype("Int64")

                edited_ref = st.data_editor(
                    ref_config_df,
                    column_config={
                        "컬럼명": st.column_config.TextColumn("컬럼명", disabled=True),
                        "매칭 순번": st.column_config.NumberColumn(
                            "매칭 순번", min_value=1, step=1,
                            help="파일A와 같은 순서로 1, 2, 3...을 입력하세요. 비워두면 미선택입니다.",
                        ),
                        "가져올 열": st.column_config.CheckboxColumn("결과에 포함", default=True),
                    },
                    hide_index=True,
                    use_container_width=True,
                    key=f"editor_ref_{idx}_{item['name']}_{item.get('sheet', '')}",
                )

                # 변경사항 동기화
                item["keys_b"] = ordered_keys_from_editor(edited_ref, "매칭 순번")
                item["cols"] = edited_ref[edited_ref["가져올 열"]]["컬럼명"].tolist()

                if item["keys_b"]:
                    st.info(f"선택된 기준열 (순서): **{' → '.join(item['keys_b'])}**")

                if st.button("이 파일 삭제", key=f"remove_{idx}", icon=":material/close:", use_container_width=True):
                    remove_idx = idx

    if remove_idx is not None:
        st.session_state.ref_files.pop(remove_idx)
        st.rerun()

st.divider()

# ----------------------------------------------------------------------
# 3. 매칭열 요약 (실행 전 최종 확인용)
# ----------------------------------------------------------------------
keys_a = st.session_state.get("keys_a", [])
df_a = st.session_state.df_a
ref_files = st.session_state.ref_files

if keys_a and ref_files:
    st.subheader(":material/search: 매칭 기준열 요약 (실행 전 확인)")

    max_len = len(keys_a)

    file_a_label = "파일A"
    if st.session_state.sheet_a_name:
        file_a_label = f"파일A ({st.session_state.sheet_a_name})"

    summary = {"순번": list(range(1, max_len + 1)), file_a_label: keys_a}
    length_mismatch = False
    seen_labels = {}
    for item in ref_files:
        base_label = item["name"]
        if item.get("sheet"):
            base_label = f"{item['name']} ({item['sheet']})"
        # 같은 이름+시트 조합이 중복될 경우를 대비한 안전장치
        if base_label in seen_labels:
            seen_labels[base_label] += 1
            base_label = f"{base_label} #{seen_labels[base_label]}"
        else:
            seen_labels[base_label] = 1

        col_vals = item["keys_b"] + ["-"] * (max_len - len(item["keys_b"]))
        col_vals = col_vals[:max_len]
        if len(item["keys_b"]) != max_len:
            length_mismatch = True
        summary[base_label] = col_vals

    st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

    if length_mismatch:
        st.warning(
            "일부 비교 파일의 매칭 기준열 개수가 파일A와 다릅니다. 순번을 다시 확인하세요.",
            icon=":material/warning:",
        )

    st.divider()

# ----------------------------------------------------------------------
# 4. 매칭 실행 및 결과
# ----------------------------------------------------------------------
st.subheader(":material/rocket_launch: 4. 다중 매칭 실행 및 결과 다운로드")

can_run = df_a is not None and len(keys_a) > 0 and len(ref_files) > 0 and all(
    item["keys_b"] and item["cols"] and len(item["keys_b"]) == len(keys_a) for item in ref_files
)

# 상태 체크 메시지
if df_a is None:
    st.warning("좌측에서 기준 파일(파일A)을 불러오세요.", icon=":material/arrow_back:")
elif not keys_a:
    st.warning("좌측에서 파일A의 매칭 기준열을 순번으로 지정하세요.", icon=":material/arrow_back:")
elif not ref_files:
    st.warning("우측에서 비교할 파일을 하나 이상 추가하세요.", icon=":material/arrow_forward:")
elif not can_run:
    st.warning(
        "우측 비교 파일의 매칭 순번/가져올 열 설정을 완료하세요 (개수가 파일A와 일치해야 합니다).",
        icon=":material/arrow_forward:",
    )

# 실행 버튼
if st.button("다중 매칭 시작 및 통합 저장", type="primary", disabled=not can_run, use_container_width=True):
    with st.spinner("다중 파일 병합 및 불일치 사유 분석 중..."):
        try:
            result_df = svc.process_multi_merge(df_a, ref_files, keys_a)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                result_df.to_excel(writer, index=False, sheet_name="통합결과")
            buffer.seek(0)

            st.success(
                f"처리가 완료되었습니다! (총 데이터: {len(result_df):,}건)",
                icon=":material/celebration:",
            )
            st.dataframe(result_df.head(50), use_container_width=True)

            st.download_button(
                label="결과 파일 다운로드 (.xlsx)",
                data=buffer,
                file_name="다중통합매칭_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"처리 중 오류 발생: {e}", icon=":material/error:")
