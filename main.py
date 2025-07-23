#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
연구비 처리 자동화 프로그램 - 메인 실행 파일

작성자: 차세대지원팀 데이터 김훈민
작성일자: 2025-07-22
"""

import logging
from research_gui import ResearchFundGUI
from config import LOG_FILE, LOG_FORMAT


def setup_logging():
    '''로깅 설정'''
    logging.basicConfig(
        level=logging.INFO,
        format=LOG_FORMAT,
        handlers=[
            logging.FileHandler(LOG_FILE, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )


def main():
    '''메인 함수'''
    setup_logging()
    logging.info("연구비 처리 자동화 프로그램 시작")
    
    try:
        app = ResearchFundGUI()
        app.run() # In research_gui.py
    except Exception as e:
        logging.error(f"프로그램 실행 중 오류 발생: {str(e)}")
        raise
    finally:
        logging.info("연구비 처리 자동화 프로그램 종료")


if __name__ == "__main__":
    main()
