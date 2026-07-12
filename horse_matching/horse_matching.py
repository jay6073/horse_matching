import reflex as rx
import pandas as pd
import io
from dataclasses import dataclass, field

# 기존 엑셀 서비스 레이어 가져오기
from excel_service import MultiDataMatcherService

svc = MultiDataMatcherService()


# ----------------------------------------------------------------------
# [모델] 비교 파일(파일B) 1개에 대한 설정 정보
# ----------------------------------------------------------------------
# Reflex 0.9부터 rx.Base가 제거되었습니다. 커스텀 var 타입은
# dataclass(권장) 또는 Pydantic V2 모델을 사용해야 합니다.
@dataclass
class RefFileItem:
    """비교 파일 카드 하나에 대응되는 데이터 구조."""
    name: str
    cols_all: list[str] = field(default_factory=list)
    keys_b: list[str] = field(default_factory=list)
    cols: list[str] = field(default_factory=list)


# ----------------------------------------------------------------------
# [백엔드] 상태 및 비즈니스 로직 관리소
# ----------------------------------------------------------------------
class MatcherState(rx.State):
    # 1. 파일A 관련 상태
    df_a_name: str = "선택된 파일 없음"
    df_a_loaded: bool = False
    columns_a: list[str] = []
    keys_a: list[str] = []  # 파일A에서 선택한 매칭 기준열들

    # 2. 비교 파일들(파일B들) 목록 주머니
    ref_files: list[RefFileItem] = []

    # 3. 진행 상태 및 결과 제어 변수
    is_processing: bool = False
    result_ready: bool = False
    result_message: str = ""
    result_row_count: int = 0

    # 내부 연산용 데이터프레임 임시 저장 주머니 (화면 노출 X, 백엔드 전용 var)
    _df_a: pd.DataFrame = None
    _result_df: pd.DataFrame = None
    # ref_files(dataclass, 프론트엔드용 메타데이터)와 같은 인덱스로 병렬 관리되는
    # 실제 DataFrame 저장소. DataFrame은 JSON 직렬화가 안 되므로 프론트엔드 var인
    # ref_files에는 넣을 수 없어 백엔드 전용(밑줄 프리픽스) var로 별도 보관합니다.
    _ref_dfs: list = []

    async def handle_upload_a(self, files: list[rx.UploadFile]):
        """기준 파일 A 업로드 및 정제"""
        if not files:
            return
        file = files[0]
        try:
            file_bytes = await file.read()
            uploaded_file = io.BytesIO(file_bytes)

            # 자동 헤더 탐지 및 정제 (기본 Sheet1 기준)
            h_idx = svc.find_best_header(uploaded_file, sheet_name="Sheet1")
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, header=h_idx)

            self._df_a = svc.sanitize(df)
            self.df_a_name = file.filename
            self.columns_a = list(self._df_a.columns)
            self.keys_a = []
            self.df_a_loaded = True
            # 업로드 완료 후 드롭존에 남아있던 선택 파일 표시를 정리
            yield rx.clear_selected_files("upload_a")
        except Exception as e:
            self.df_a_name = f"❌ 파일A 로딩 에러: {str(e)}"

    async def handle_add_ref_file(self, files: list[rx.UploadFile]):
        """비교 파일 추가 (+) 버튼 클릭 시 대시보드에 카드 추가"""
        if not files:
            return
        file = files[0]
        try:
            file_bytes = await file.read()
            uploaded_file = io.BytesIO(file_bytes)

            h_idx = svc.find_best_header(uploaded_file, sheet_name="Sheet1")
            uploaded_file.seek(0)
            df_b = pd.read_excel(uploaded_file, header=h_idx)
            sanitized_df_b = svc.sanitize(df_b)

            name = file.filename.rsplit(".", 1)[0]

            self.ref_files.append(
                RefFileItem(
                    name=name,
                    cols_all=list(sanitized_df_b.columns),
                    keys_b=[],
                    cols=[],
                )
            )
            # ref_files와 같은 인덱스로 실제 DataFrame을 백엔드 전용 리스트에 보관
            self._ref_dfs.append(sanitized_df_b)
            # 카드에 추가되었으니 드롭존의 선택 파일 표시는 정리
            yield rx.clear_selected_files("upload_b")
        except Exception as e:
            self.result_message = f"❌ 비교 파일 추가 실패: {str(e)}"

    def toggle_key_a(self, col: str, checked: bool):
        """파일A 매칭 기준열 체크박스 토글"""
        current = list(self.keys_a)
        if checked and col not in current:
            current.append(col)
        elif not checked and col in current:
            current.remove(col)
        self.keys_a = current

    def toggle_ref_key_b(self, idx: int, col: str, checked: bool):
        """특정 인덱스 비교 파일의 매칭 기준열 토글"""
        item = self.ref_files[idx]
        current = list(item.keys_b)
        if checked and col not in current:
            current.append(col)
        elif not checked and col in current:
            current.remove(col)
        item.keys_b = current
        self.ref_files[idx] = item

    def toggle_ref_col(self, idx: int, col: str, checked: bool):
        """특정 인덱스 비교 파일에서 가져올 열 토글"""
        item = self.ref_files[idx]
        current = list(item.cols)
        if checked and col not in current:
            current.append(col)
        elif not checked and col in current:
            current.remove(col)
        item.cols = current
        self.ref_files[idx] = item

    def remove_ref_file(self, idx: int):
        """추가된 비교 파일 카드 삭제 (st.rerun 없이 즉시 반영)"""
        self.ref_files.pop(idx)
        # 병렬로 관리되는 실제 DataFrame 저장소도 같은 인덱스에서 제거
        if 0 <= idx < len(self._ref_dfs):
            self._ref_dfs.pop(idx)

    def run_multi_match(self):
        """다중 매칭 통합 연산 가동 버튼"""
        self.is_processing = True
        self.result_ready = False
        yield  # 화면에 로딩 스피너 작동 유도

        if self._df_a is None:
            self.result_message = "❌ 먼저 기준 파일(파일A)을 업로드해주세요."
            self.is_processing = False
            return

        if not self.ref_files:
            self.result_message = "❌ 비교할 파일을 최소 1개 이상 추가해주세요."
            self.is_processing = False
            return

        if not self.keys_a:
            self.result_message = "❌ 파일A의 매칭 기준열을 선택해주세요."
            self.is_processing = False
            return

        # process_multi_merge는 각 비교 파일마다 current_keys_a = keys_a[:len(keys_b)]
        # 방식으로 파일A 기준열을 앞에서부터 잘라 사용하므로, 각 비교 파일의 keys_b
        # 개수만큼 파일A 기준열이 선택되어 있어야 합니다.
        for item in self.ref_files:
            if not item.keys_b:
                self.result_message = f"❌ '{item.name}' 파일의 매칭 기준열을 선택해주세요."
                self.is_processing = False
                return
            if len(self.keys_a) < len(item.keys_b):
                self.result_message = (
                    f"❌ '{item.name}'의 기준열({len(item.keys_b)}개)만큼 "
                    f"파일A 기준열을 최소 {len(item.keys_b)}개 이상 선택해주세요."
                )
                self.is_processing = False
                return

        try:
            # process_multi_merge가 기대하는 형태로 변환:
            # [{"df": 실제DataFrame, "keys_b": [...], "cols": [...], "name": "..."}]
            ref_files_for_merge = [
                {
                    "df": self._ref_dfs[i],
                    "keys_b": item.keys_b,
                    "cols": item.cols,
                    "name": item.name,
                }
                for i, item in enumerate(self.ref_files)
            ]

            result_df = svc.process_multi_merge(self._df_a, ref_files_for_merge, self.keys_a)
            self._result_df = result_df
            self.result_row_count = len(result_df)
            self.result_message = f"🎉 통합 매칭 완료! 총 {self.result_row_count}건 처리되었습니다."
            self.result_ready = True

        except Exception as e:
            self.result_message = f"❌ 매칭 연산 중 에러: {str(e)}"

        self.is_processing = False

    def download_result(self):
        """결과 데이터프레임을 엑셀 바이트로 변환해 다운로드"""
        if self._result_df is None or self._result_df.empty:
            return rx.toast.error("다운로드할 결과가 없습니다. 먼저 매칭을 실행해주세요.")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            self._result_df.to_excel(writer, index=False, sheet_name="결과")
        output.seek(0)

        return rx.download(
            data=output.getvalue(),
            filename="다중통합매칭_결과.xlsx",
        )


# ----------------------------------------------------------------------
# [프론트엔드] 대시보드 화면 뷰 (View) 조립
# ----------------------------------------------------------------------
def index() -> rx.Component:
    return rx.center(
        rx.vstack(
            rx.heading("📊 다중 통합 매칭 시스템 v6.9", size="8"),
            rx.text("자동 헤더 탐지 · 다중 파일 병합 · 불일치 사유 분석", color="gray"),
            rx.divider(),

            # ---- Grid 레이아웃 ----
            # 데스크톱: 왼쪽 컬럼에 파일A 카드(위) + 실행/다운로드 카드(아래),
            #           오른쪽 컬럼에 비교 파일 카드(세로로 길게, 두 행 모두 차지)
            # 모바일(좁은 화면): 자동으로 1열 세로 배치로 전환
            rx.box(
                # ---- [영역 a] 파일A 설정 카드 ----
                rx.card(
                    rx.vstack(
                        rx.heading("1. 기준 파일 (파일A) 설정", size="4"),
                        rx.upload(
                            rx.vstack(
                                rx.text("📁 여기에 기준 파일 A(xlsx)를 드래그하거나 클릭하세요."),
                                rx.foreach(
                                    rx.selected_files("upload_a"),
                                    lambda f: rx.text(f"📎 {f}", size="2", color="blue", font_weight="bold"),
                                ),
                                spacing="1",
                            ),
                            id="upload_a",
                            border="2px dashed #ccc",
                            padding="4",
                            width="100%",
                        ),
                        rx.button(
                            "파일A 불러오기",
                            on_click=MatcherState.handle_upload_a(rx.upload_files(upload_id="upload_a")),
                            color_scheme="blue",
                            width="100%"
                        ),
                        rx.cond(
                            MatcherState.df_a_loaded,
                            rx.vstack(
                                rx.box(
                                    rx.text(f"🟢 로드 완료: {MatcherState.df_a_name}", font_weight="bold", color="green"),
                                    padding="2", bg="#f0fff4", border_radius="md", width="100%"
                                ),
                                rx.text("👉 매칭 기준열 선택 (다중 선택 가능):", font_weight="bold", size="2"),
                                rx.vstack(
                                    rx.foreach(
                                        MatcherState.columns_a,
                                        lambda col: rx.checkbox(
                                            col,
                                            checked=MatcherState.keys_a.contains(col),
                                            on_change=lambda checked: MatcherState.toggle_key_a(col, checked),
                                        )
                                    ),
                                    spacing="1",
                                    align_items="start",
                                    max_height="150px",
                                    overflow_y="auto",
                                    width="100%",
                                    border="1px solid #e2e8f0",
                                    border_radius="md",
                                    padding="2",
                                ),
                                width="100%", spacing="2"
                            )
                        ),
                    ),
                    width="100%", padding="5",
                    style={"gridArea": "a"},
                ),

                # ---- [영역 c] 통합 매칭 가동 및 다운로드 구역 ----
                rx.card(
                    rx.vstack(
                        rx.heading("3. 다중 매칭 실행 및 통합 저장", size="4"),
                        rx.button(
                            "🔍 다중 매칭 시작 및 통합 저장",
                            on_click=MatcherState.run_multi_match,
                            loading=MatcherState.is_processing,
                            color_scheme="iris",
                            size="3",
                            width="100%"
                        ),
                        rx.cond(
                            MatcherState.result_ready,
                            rx.box(
                                rx.text(MatcherState.result_message, font_weight="bold", color="blue"),
                                rx.button(
                                    "📥 결과 파일 다운로드 (.xlsx)",
                                    on_click=MatcherState.download_result,
                                    color_scheme="green",
                                    margin_top="2",
                                    width="100%"
                                ),
                                padding="3", bg="#ebf8ff", border_radius="md", width="100%"
                            )
                        )
                    ),
                    width="100%", padding="5",
                    style={"gridArea": "c"},
                ),

                # ---- [영역 b] 비교 파일 추가 및 동적 카드 배열 구역 ----
                rx.card(
                    rx.vstack(
                        rx.heading("2. 비교 파일 추가 및 설정", size="4"),
                        rx.upload(
                            rx.vstack(
                                rx.text("➕ 추가할 비교 파일들을 여기에 드래그앤드롭 하세요."),
                                rx.foreach(
                                    rx.selected_files("upload_b"),
                                    lambda f: rx.text(f"📎 {f}", size="2", color="teal", font_weight="bold"),
                                ),
                                spacing="1",
                            ),
                            id="upload_b",
                            border="2px dashed #aaa",
                            padding="4",
                            width="100%",
                        ),
                        rx.button(
                            "이 파일 추가(+)",
                            on_click=MatcherState.handle_add_ref_file(rx.upload_files(upload_id="upload_b")),
                            color_scheme="cyan",
                            width="100%"
                        ),
                        rx.foreach(
                            MatcherState.ref_files,
                            lambda item, idx: rx.card(
                                rx.vstack(
                                    rx.hstack(
                                        rx.text(f"📋 {item.name}", font_weight="bold", color="indigo"),
                                        rx.spacer(),
                                        rx.button(
                                            "🗑️ 삭제",
                                            on_click=lambda: MatcherState.remove_ref_file(idx),
                                            color_scheme="red",
                                            size="1"
                                        )
                                    ),
                                    rx.text("• 매칭 기준열 (파일A와 같은 순서 대응):", size="2"),
                                    rx.vstack(
                                        rx.foreach(
                                            item.cols_all,
                                            lambda c: rx.checkbox(
                                                c,
                                                checked=item.keys_b.contains(c),
                                                on_change=lambda checked: MatcherState.toggle_ref_key_b(idx, c, checked),
                                            )
                                        ),
                                        spacing="1",
                                        align_items="start",
                                        max_height="120px",
                                        overflow_y="auto",
                                        width="100%",
                                        border="1px solid #e2e8f0",
                                        border_radius="md",
                                        padding="2",
                                    ),
                                    rx.text("• 가져올 열 선택:", size="2"),
                                    rx.vstack(
                                        rx.foreach(
                                            item.cols_all,
                                            lambda c: rx.checkbox(
                                                c,
                                                checked=item.cols.contains(c),
                                                on_change=lambda checked: MatcherState.toggle_ref_col(idx, c, checked),
                                            )
                                        ),
                                        spacing="1",
                                        align_items="start",
                                        max_height="120px",
                                        overflow_y="auto",
                                        width="100%",
                                        border="1px solid #e2e8f0",
                                        border_radius="md",
                                        padding="2",
                                    ),
                                    spacing="2", width="100%"
                                ),
                                width="100%", margin_top="2", background_color="#f8fafc"
                            )
                        )
                    ),
                    width="100%", padding="5",
                    style={"gridArea": "b"},
                ),

                display="grid",
                grid_template_columns="1fr 1fr",
                grid_template_areas="'a b' 'c b'",
                gap="24px",
                width="100%",
                align_items="start",
            ),
            spacing="5",
            width="1100px",
            max_width="95vw",
        ),
        padding_top="3%", padding_bottom="5%"
    )


app = rx.App()
app.add_page(index)
