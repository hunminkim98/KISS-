#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
연구비 처리 자동화 프로그램 - GUI 인터페이스

작성자: 차세대지원팀 데이터 김훈민
작성일자: 2025-07-22
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox
import logging
from research_core import ExcelFileLoader, DataClassifier, ExcelExporter
from config import (
    WINDOW_TITLE, WINDOW_SIZE, MAIN_FONT, NORMAL_FONT, BUTTON_FONT,
    BUTTON_COLOR, BUTTON_TEXT_COLOR, FILE_DIALOG_TYPES,
    TEXT_AREA_HEIGHT, TEXT_AREA_WIDTH, BUTTON_WIDTH,
    CLASSIFY_BUTTON_WIDTH, CLASSIFY_BUTTON_HEIGHT, FILE_PATH_LABEL_WIDTH,
    SAMPLE_DATA_COUNT, INFO_SAMPLE_COUNT, SAVE_FILE_TYPES, DEFAULT_OUTPUT_FILENAME
)


class ResearchFundGUI:
    '''연구비 처리 GUI 클래스'''

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)

        self.file_loader = ExcelFileLoader()
        self.data_classifier = DataClassifier()
        self.excel_exporter = ExcelExporter()
        self.selected_file = None
        self.classification_result = None

        self._create_widgets()
        self._setup_layout()

    def _create_widgets(self):
        '''위젯 생성'''
        # 제목
        self.title_label = tk.Label(self.root, text=WINDOW_TITLE, font=MAIN_FONT)

        # 파일 선택 프레임
        self.file_frame = tk.Frame(self.root)
        self.file_label = tk.Label(self.file_frame, text="Excel 파일:", font=NORMAL_FONT)
        self.file_path_label = tk.Label(
            self.file_frame, text="파일을 선택해주세요", bg="white",
            relief="sunken", width=FILE_PATH_LABEL_WIDTH, anchor="w"
        )
        self.browse_button = tk.Button(
            self.file_frame, text="파일 선택", command=self._select_file, width=BUTTON_WIDTH
        )

        # 버튼 프레임
        self.button_frame = tk.Frame(self.root)

        # 분류 실행 버튼
        self.classify_button = tk.Button(
            self.button_frame, text="데이터 분류 실행", command=self._classify_data,
            width=CLASSIFY_BUTTON_WIDTH, height=CLASSIFY_BUTTON_HEIGHT,
            font=BUTTON_FONT, bg=BUTTON_COLOR, fg=BUTTON_TEXT_COLOR, state=tk.DISABLED
        )

        # 저장 버튼
        self.save_button = tk.Button(
            self.button_frame, text="Excel 파일 저장", command=self._save_excel,
            width=CLASSIFY_BUTTON_WIDTH, height=CLASSIFY_BUTTON_HEIGHT,
            font=BUTTON_FONT, bg="#2196F3", fg=BUTTON_TEXT_COLOR, state=tk.DISABLED
        )

        # 정보 표시 영역
        self._create_text_areas()

    def _create_text_areas(self):
        '''텍스트 영역 생성'''
        # 파일 정보 영역
        self.info_frame = tk.Frame(self.root)
        self.info_label = tk.Label(self.info_frame, text="파일 정보:", font=BUTTON_FONT)
        self.info_text = tk.Text(
            self.info_frame, height=TEXT_AREA_HEIGHT, width=TEXT_AREA_WIDTH,
            wrap=tk.WORD, state=tk.DISABLED
        )
        self.info_scrollbar = tk.Scrollbar(
            self.info_frame, orient="vertical", command=self.info_text.yview
        )
        self.info_text.configure(yscrollcommand=self.info_scrollbar.set)

        # 분류 결과 영역
        self.result_frame = tk.Frame(self.root)
        self.result_label = tk.Label(self.result_frame, text="분류 결과:", font=BUTTON_FONT)
        self.result_text = tk.Text(
            self.result_frame, height=TEXT_AREA_HEIGHT, width=TEXT_AREA_WIDTH,
            wrap=tk.WORD, state=tk.DISABLED
        )
        self.result_scrollbar = tk.Scrollbar(
            self.result_frame, orient="vertical", command=self.result_text.yview
        )
        self.result_text.configure(yscrollcommand=self.result_scrollbar.set)

    def _setup_layout(self):
        '''레이아웃 설정'''
        self.title_label.pack(pady=20)

        # 파일 선택 영역
        self.file_frame.pack(pady=10, padx=20, fill="x")
        self.file_label.pack(side=tk.LEFT)
        self.file_path_label.pack(side=tk.LEFT, padx=10, fill="x", expand=True)
        self.browse_button.pack(side=tk.RIGHT)

        # 버튼 영역
        self.button_frame.pack(pady=10)
        self.classify_button.pack(side=tk.LEFT, padx=5)
        self.save_button.pack(side=tk.LEFT, padx=5)

        # 정보 표시 영역들
        self._setup_text_areas()

    def _setup_text_areas(self):
        '''텍스트 영역 레이아웃 설정'''
        # 파일 정보 영역
        self.info_frame.pack(pady=5, padx=20, fill="both", expand=True)
        self.info_label.pack(anchor="w")
        self.info_text.pack(side="left", fill="both", expand=True)
        self.info_scrollbar.pack(side="right", fill="y")

        # 분류 결과 영역
        self.result_frame.pack(pady=5, padx=20, fill="both", expand=True)
        self.result_label.pack(anchor="w")
        self.result_text.pack(side="left", fill="both", expand=True)
        self.result_scrollbar.pack(side="right", fill="y")

    def _select_file(self):
        '''파일 선택 다이얼로그'''
        file_path = filedialog.askopenfilename(
            title="지출요구 조회 Excel 파일 선택", filetypes=FILE_DIALOG_TYPES
        )

        if file_path:
            self.selected_file = file_path
            self.file_path_label.config(text=os.path.basename(file_path))
            self._load_and_display_file(file_path)

    def _load_and_display_file(self, file_path: str):
        '''파일을 로드하고 정보를 표시'''
        success = self.file_loader.load_file(file_path)
        
        if success:
            self._display_file_info()
            self.classify_button.config(state=tk.NORMAL)
        else:
            messagebox.showerror("오류", "파일을 로드할 수 없습니다.")
            self.classify_button.config(state=tk.DISABLED)

    def _display_file_info(self):
        '''파일 정보를 텍스트 영역에 표시'''
        info = self.file_loader.get_data_info()
        if info is None:
            return

        self._update_text_widget(self.info_text, self._format_file_info(info))

    def _format_file_info(self, info: dict) -> str:
        '''파일 정보를 포맷팅'''
        info_text = f"""파일 경로: {info['file_path']}
데이터 크기: {info['shape'][0]}행 × {info['shape'][1]}열

컬럼 목록:
"""
        for i, col in enumerate(info['columns'], 1):
            info_text += f"  {i}. {col}\n"

        info_text += "\n샘플 데이터 (상위 5행):\n" + "=" * 50 + "\n"

        for i, row in enumerate(info['sample_data'][:INFO_SAMPLE_COUNT], 1):
            info_text += f"\n[행 {i}]\n"
            for col, value in row.items():
                info_text += f"  {col}: {value}\n"

        return info_text

    def _classify_data(self):
        '''데이터 분류를 실행합니다.'''
        if self.file_loader.data is None:
            messagebox.showerror("오류", "먼저 Excel 파일을 로드해주세요.")
            return

        try:
            self._set_classify_button_state(False, "분류 중...")
            
            if '적요' not in self.file_loader.data.columns:
                self._show_column_error()
                return

            logging.info("데이터 분류 시작")
            self.classification_result = self.data_classifier.classify_data(self.file_loader.data)
            
            self._display_classification_result()
            self._show_completion_message()

            # 분류 완료 후 저장 버튼 활성화
            self.save_button.config(state=tk.NORMAL)
            logging.info("데이터 분류 완료")

        except ValueError as e:
            logging.error(f"데이터 검증 오류: {str(e)}")
            messagebox.showerror("데이터 오류", str(e))
        except Exception as e:
            logging.error(f"예상치 못한 오류: {str(e)}")
            messagebox.showerror("오류", f"데이터 분류 중 오류가 발생했습니다:\n{str(e)}")
        finally:
            self._set_classify_button_state(True, "데이터 분류 실행")

    def _save_excel(self):
        '''Excel 파일로 저장합니다.'''
        if self.classification_result is None:
            messagebox.showerror("오류", "먼저 데이터 분류를 실행해주세요.")
            return

        # 파일 저장 다이얼로그
        output_path = filedialog.asksaveasfilename(
            title="Excel 파일 저장",
            defaultextension=".xlsx",
            filetypes=SAVE_FILE_TYPES,
            initialfile=DEFAULT_OUTPUT_FILENAME
        )

        if not output_path:
            return

        try:
            # 저장 버튼 비활성화
            self.save_button.config(state=tk.DISABLED, text="저장 중...")
            self.root.update()

            # Excel 파일 생성
            business_data = self.data_classifier.get_business_data()
            research_data = self.data_classifier.get_research_data()

            success = self.excel_exporter.export_to_excel(
                business_data, research_data, output_path
            )

            if success:
                # 저장 완료 메시지
                stats = self.data_classifier.get_classification_stats()
                completion_msg = (
                    f"Excel 파일이 성공적으로 저장되었습니다!\n\n"
                    f"저장 경로: {output_path}\n\n"
                    f"저장된 데이터:\n"
                    f"• 사업비: {stats['business_count']:,}건\n"
                    f"• 연구비: {stats['research_count']:,}건\n\n"
                    f"생성된 시트:\n"
                    f"• 대시보드\n"
                    f"• 집행관리(사업비)\n"
                    f"• 집행관리(연구비)\n"
                    f"• 사업비 (요약 시트)\n"
                    f"• 연구비 (요약 시트)"
                )
                messagebox.showinfo("저장 완료", completion_msg)
                logging.info(f"Excel 파일 저장 완료: {output_path}")
            else:
                messagebox.showerror("오류", "Excel 파일 저장에 실패했습니다.")

        except Exception as e:
            logging.error(f"Excel 저장 중 오류: {str(e)}")
            messagebox.showerror("오류", f"Excel 파일 저장 중 오류가 발생했습니다:\n{str(e)}")
        finally:
            # 저장 버튼 다시 활성화
            self.save_button.config(state=tk.NORMAL, text="Excel 파일 저장")

    def _set_classify_button_state(self, enabled: bool, text: str):
        '''분류 버튼 상태 설정'''
        state = tk.NORMAL if enabled else tk.DISABLED
        self.classify_button.config(state=state, text=text)
        self.root.update()

    def _show_column_error(self):
        '''컬럼 오류 메시지 표시'''
        available_columns = '\n'.join([f"• {col}" for col in self.file_loader.data.columns])
        messagebox.showerror(
            "오류",
            f"'적요' 컬럼을 찾을 수 없습니다.\n\n"
            f"사용 가능한 컬럼:\n{available_columns}\n\n"
            f"Excel 파일의 컬럼명을 확인해주세요."
        )

    def _show_completion_message(self):
        '''완료 메시지 표시'''
        stats = self.data_classifier.get_classification_stats()
        completion_msg = (
            f"데이터 분류가 완료되었습니다!\n\n"
            f"분류 결과:\n"
            f"• 전체: {stats['total']:,}건\n"
            f"• 사업비: {stats['business_count']:,}건\n"
            f"• 연구비: {stats['research_count']:,}건\n"
            f"• 미분류: {stats['unclassified_count']:,}건"
        )

        if stats['unclassified_count'] > 0:
            completion_msg += f"\n\n경고: {stats['unclassified_count']}건의 미분류 데이터가 있습니다."

        messagebox.showinfo("완료", completion_msg)

    def _display_classification_result(self):
        '''분류 결과를 텍스트 영역에 표시'''
        if self.classification_result is None:
            return

        result_text = self._format_classification_result()
        self._update_text_widget(self.result_text, result_text)

    def _format_classification_result(self) -> str:
        '''분류 결과를 포맷팅'''
        stats = self.data_classifier.get_classification_stats()
        
        result_text = f"""=== 데이터 분류 결과 ===

전체 데이터: {stats['total']:,}건

분류 통계:
• 사업비 (25 차세대): {stats['business_count']:,}건 ({stats['business_percentage']:.1f}%)
• 연구비 (25 연구비): {stats['research_count']:,}건 ({stats['research_percentage']:.1f}%)
• 미분류: {stats['unclassified_count']:,}건 ({stats['unclassified_percentage']:.1f}%)

=== 사업비 데이터 샘플 (상위 3건) ===
"""

        # 사업비 데이터 샘플
        business_data = self.data_classifier.get_business_data()
        result_text += self._format_data_sample(business_data, "사업비")

        result_text += "\n=== 연구비 데이터 샘플 (상위 3건) ===\n"

        # 연구비 데이터 샘플
        research_data = self.data_classifier.get_research_data()
        result_text += self._format_data_sample(research_data, "연구비")

        # 미분류 데이터 경고
        if stats['unclassified_count'] > 0:
            result_text += f"\n경고: {stats['unclassified_count']}건의 미분류 데이터가 있습니다.\n"
            result_text += "미분류 데이터의 '적요' 컬럼을 확인해주세요.\n"

        return result_text

    def _format_data_sample(self, data, data_type: str) -> str:
        '''데이터 샘플을 포맷팅'''
        if data is None or data.empty:
            return f"\n{data_type} 데이터가 없습니다.\n"

        sample_text = ""
        for i, (_, row) in enumerate(data.head(SAMPLE_DATA_COUNT).iterrows(), 1):
            sample_text += f"\n[{data_type} {i}]\n"
            for col, value in row.items():
                sample_text += f"  {col}: {value}\n"
        
        return sample_text

    def _update_text_widget(self, widget, text: str):
        '''텍스트 위젯 업데이트'''
        widget.config(state=tk.NORMAL)
        widget.delete(1.0, tk.END)
        widget.insert(tk.END, text)
        widget.config(state=tk.DISABLED)

    def run(self): 
        '''GUI 실행 from main.py'''
        self.root.mainloop()
