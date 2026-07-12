import pandas as pd

class MultiDataMatcherService:
    @staticmethod
    def get_sheets(path: str) -> list[str]:
        try:
            from python_calamine import CalamineWorkbook
            return CalamineWorkbook.from_path(path).sheet_names
        except Exception:
            try: return pd.ExcelFile(path).sheet_names
            except: return []

    @staticmethod
    def find_best_header(path: str, sheet_name: str) -> int:
        """상위 10개 행을 분석하여 데이터가 가장 많이 차 있는 행을 헤더로 자동 선택"""
        try:
            temp_df = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=10, engine="calamine")
            best_row = 0
            max_non_empty = -1

            for i, row in temp_df.iterrows():
                non_empty_count = row.count()
                if non_empty_count > max_non_empty:
                    max_non_empty = non_empty_count
                    best_row = i
                if non_empty_count >= (len(temp_df.columns) * 0.8):
                    return i
            return best_row
        except:
            return 0

    @staticmethod
    def sanitize(df: pd.DataFrame) -> pd.DataFrame:
        """데이터 정제: 문자열 변환, 공백 제거 및 무의미한 기호(-, . 등) 처리"""
        df = df.copy()
        ignore_values = ['-', '.', 'nan', 'null', 'N/A', 'none', '?', '미기재']

        for col in df.columns:
            df[col] = df[col].apply(lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else ('' if pd.isna(x) else str(x)))
            df[col] = df[col].str.strip()
            df[col] = df[col].apply(lambda x: '' if x.lower() in ignore_values else x)
        return df

    @classmethod
    def process_multi_merge(cls, df_a: pd.DataFrame, ref_files: list, keys_a_all: list) -> pd.DataFrame:
        """다중 파일 병합 및 상세 사유 분석 로직 (불일치 항목 추적 기능 포함)"""
        final_df = df_a.copy()
        final_df['_origin_order'] = range(len(final_df))

        def is_empty(val):
            return str(val).strip() == '' or pd.isna(val)

        for item in ref_files:
            df_b = item['df'].copy()
            keys_b = item['keys_b']
            selected_cols = item['cols']
            file_label = item['name']

            current_keys_a = keys_a_all[:len(keys_b)]

            # 비교 대상 파일(B)에서 기준열이 모두 비어있는 행 제외
            mask_b = df_b[keys_b].apply(lambda row: row.map(is_empty).all(), axis=1)
            df_b_filtered = df_b[~mask_b].copy()

            cols_to_keep = list(dict.fromkeys(keys_b + selected_cols))
            df_b_prep = df_b_filtered[cols_to_keep].drop_duplicates(subset=keys_b, keep='first')

            rename_map = {c: f"{c}_{file_label}" for c in cols_to_keep if c in final_df.columns}
            df_b_prep = df_b_prep.rename(columns=rename_map)
            target_keys_b = [rename_map.get(k, k) for k in keys_b]

            final_df = pd.merge(
                final_df,
                df_b_prep,
                left_on=current_keys_a,
                right_on=target_keys_b,
                how='left',
                indicator='_merge_flag'
            )

            # 상세 사유 분석을 위해 첫 번째 키를 기준으로 한 set 준비
            key1_b_name = keys_b[0]
            b_key1_values = set(df_b_filtered[key1_b_name].unique())

            def get_reason(row):
                # 1. 마스터 파일(A)의 기준열 정보가 없는 경우
                if all(is_empty(row[k]) for k in current_keys_a):
                    return '실패', '마스터파일에 데이터 없음'

                # 2. 매칭 성공 시
                if row['_merge_flag'] == 'both':
                    return '성공', ''

                # 3. 매칭 실패 시 상세 분석 로직
                val_a1 = str(row[current_keys_a[0]]).strip()
                if val_a1 not in b_key1_values:
                    # 첫 번째 기준값 자체가 비교 파일에 없는 경우
                    return '실패', f'[{file_label}] 데이터 없음'
                else:
                    # 첫 번째 값은 일치하지만 다른 기준열들 중 불일치가 있는 경우
                    # 비교 파일(B)에서 동일한 첫 번째 키를 가진 행을 찾아 대조
                    row_b_candidates = df_b_filtered[df_b_filtered[key1_b_name] == val_a1]
                    if not row_b_candidates.empty:
                        row_b = row_b_candidates.iloc[0]
                        mismatched_cols = []

                        for i in range(len(current_keys_a)):
                            val_a = str(row[current_keys_a[i]]).strip()
                            val_b = str(row_b[keys_b[i]]).strip()
                            if val_a != val_b:
                                mismatched_cols.append(keys_b[i])

                        if mismatched_cols:
                            col_str = ", ".join(mismatched_cols)
                            return '실패', f'({file_label}) {col_str} 데이터가 다름'

                    return '실패', f'[{file_label}] 기준열 정보 불일치'

            analysis = final_df.apply(get_reason, axis=1, result_type='expand')
            final_df[f'매칭_{file_label}'] = analysis[0]
            final_df[f'사유_{file_label}'] = analysis[1]

            if '_merge_flag' in final_df.columns:
                final_df.drop(columns=['_merge_flag'], inplace=True)

        return final_df.sort_values('_origin_order').drop(columns=['_origin_order'])

    @staticmethod
    def save_excel(df: pd.DataFrame, path: str) -> None:
        try:
            with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='통합결과')
        except PermissionError:
            raise PermissionError("파일이 이미 열려있습니다. 엑셀을 닫고 다시 시도해주세요.")
