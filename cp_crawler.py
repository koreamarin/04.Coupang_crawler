from re import S
from unicodedata import category
from PyQt5.QtWidgets import *
from PyQt5 import uic
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from posixpath import split
from ntpath import join
from collections import Counter
from webdriver_manager.chrome import ChromeDriverManager
import sys, os, time, requests, webbrowser, pyautogui, threading, getmac, subprocess, ast, pyperclip
from pprint import pprint
import datetime as dt
from cryptography.fernet import Fernet
from tkinter import filedialog

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UI_PATH = "cp_crawler.ui"

class MainDialog(QDialog) :
    def __init__(self) :
        QDialog.__init__(self, None)
        uic.loadUi(os.path.join(BASE_DIR, UI_PATH), self)

        self.chrome_fname = 'C:\Program Files\Google\Chrome Beta\Application\chrome.exe'
        self.chrome_driver_fname = 'chromedriver.exe'

        # tablewidget 행 열 개수 설정
        self.crawling_status_tableWidget.setRowCount(1000)
        self.crawling_status_tableWidget.setColumnCount(8)

        self.excel_status_tableWidget.setRowCount(1000)
        self.excel_status_tableWidget.setColumnCount(100)

        # tablewidget 첫 행 헤더 텍스트 설정
        self.crawling_status_tableWidget.setHorizontalHeaderLabels(["로켓종류", "상품명", "가격", "품절여부", "썸네일URL", "상품 URL"])
        self.excel_status_tableWidget.setHorizontalHeaderLabels(["상품상태","카테고리ID","상품명","판매가","재고수량","A/S 안내내용","A/S 전화번호","대표 이미지 파일명","추가 이미지 파일명","상품 상세정보","판매자 상품코드","판매자 바코드","제조사","브랜드","제조일자","유효일자","부가세","미성년자 구매","구매평 노출여부","원산지 코드","수입사","복수원산지 여부","원산지 직접 입력","베송방법","배송비 유형","기본배송비","배송비 결제방식","조건부무료-상품판매가합계","수량별부과-수량","반품배송비","교환배송비","지역별 차등배송비 정보","별도설치비","판매자 특이사항","즉시할인 값","즉시할인 단위","복수구매할인 조건 값","복수구매할인 조건 단위","복수구매할인 값","복수구매할인 단위","상품구매시 포인트 지급 값","상품구매시 포인트 지급 단위","텍스트리뷰 작성시 지급 포인트","포토/동영상 리뷰 작성시 지급 포인트","한달사용 텍스트리뷰 작성시 지급 포인트","한달사용 포토/동영상 리뷰 작성시 지급 포인트","톡톡친구/스토어찜고객 리뷰 작성시 지급 포인트","무이자 할부 개월","사은품","옵션형태","옵션명","옵션값","옵션가","옵션 재고수량","추가상품명","추가상품값","추가상품가","추가상품 재고수량","상품정보제공고시 품명","상품정보제공고시 모델명","상품정보제공고시 인증 허가사항","상품정보제공고시 제조자","스토어찜회원 전용여부","문화비 소득공제","ISBN","독립출판","ISSN","출간일","출판사","글작가","그림작가","번역자명"])                                                                                                                                                                                                                                        

        self.single_product_save_radioButton.setChecked(True)

        self.cp_crawling_start_pushButton.setText('현재 URL\n수집 시작')

        # self.cp_browser_start()
        self.forbidden_keyword_read()

        self.cp_crawling_start_pushButton.clicked.connect(self.cp_collect_start)
        self.forbidden_keyword_save_pushButton.clicked.connect(self.forbidden_keyword_save)
        

    # 쿠팡 브라우저 열기 함수.
    def cp_browser_start(self) :
        # 브라우저 꺼짐 방지
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)

        # 불필요한 에러 메시지 없애기 (꼭 필요한 코드는 아님. 없어도됨. 다만 이상한 오류코드가 발생하기에 그거 없애려고 추가하는 것)
        chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

        try :
            # 크롬 베타버전으로 열기, 베타버전
            chrome_options.binary_location = f'{self.chrome_fname}'

        except : 
            self.status_label1.setText('')
            self.status_label2.setText('크롬 베타버전이 선택되지 않았습니다.')
            self.status_label3.setText('')
            return 0

        try :    
            # 크롬 드라이버 104 버전 열기, 베타버전
            self.cp_crawling_driver = webdriver.Chrome(rf'{self.chrome_driver_fname}', chrome_options=chrome_options)

        except :
            self.status_label1.setText('')
            self.status_label2.setText('같은 폴더내에 크롬드라이버 version 104 (베타버전)이 없습니다.')
            self.status_label3.setText('')
            return 0

        # 쿠팡 access denied 안나게 하는 코드
        self.cp_crawling_driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": """ Object.defineProperty(navigator, 'webdriver', { get: () => undefined }) """})

        # 웹페이지가 로딩 될때까지 5초는 기다림
        self.cp_crawling_driver.implicitly_wait(5)

        # 화면 최대화
        self.cp_crawling_driver.maximize_window()

        # 쿠팡 한번 접속 해줌. access denied 안뜨게 하기 위함.
        self.cp_crawling_driver.get('https://www.coupang.com/')

    # 쿠팡 수집 함수
    def cp_collect_start(self) :
        cp_thumbnail_collect = self.cp_thumbnail_collect_start()
        if cp_thumbnail_collect != False :
            self.cp_detailpage_collect_start(cp_thumbnail_collect)
        if cp_thumbnail_collect == False :
            return

    # 쿠팡 썸네일 수집 함수
    def cp_thumbnail_collect_start(self) :
        self.crawling_status_tableWidget.clear()
        self.crawling_status_tableWidget.setHorizontalHeaderLabels(["로켓종류", "상품명", "가격", "품절여부", "썸네일URL", "상품 URL"])

        if self.rocket_delivery_checkBox.isChecked() == False and self.rocket_merchant_checkBox.isChecked() == False and self.rocket_global_checkBox.isChecked() == False and self.rocket_fresh_checkBox.isChecked() == False :
            self.status_label.setText('수집항목설정을 한 개 이상 체크하여주세요.')
            return False

        if self.product_collect_amount_lineEdit.text() == '' :
            self.status_label.setText('수집개수를 입력해주세요.')
            return False

        total_collect_product = 0
        page_num = 0
        delivery_kind_list = []
        product_name_list = []
        price_list = []
        product_url_list = []
        sold_out_list= []
        thumbnail_img_list = []

        # 수집할 개수 확인
        product_collect_amount = int(self.product_collect_amount_lineEdit.text())

        # 현재 URL 추출
        current_url = self.cp_crawling_driver.current_url
        current_url_list = current_url.split('?')

        while total_collect_product+1 < product_collect_amount :
            page_num += 1
            current_page_collect_product_amount = 1

            # URL에 페이지 정보 넣기
            page_url = f'{current_url_list[0]}?page={page_num}'

            # page_url로 이동
            self.cp_crawling_driver.get(page_url)

            # baby_product 수집
            baby_product_infos = self.cp_crawling_driver.find_elements(By.CSS_SELECTOR, '.baby-product.renew-badge ')

            for baby_product_info in baby_product_infos :
                delivery_url = baby_product_info.find_element(By.CSS_SELECTOR, '.badge > img').get_attribute('src')
                if delivery_url == 'https://image6.coupangcdn.com/image/cmg/icon/ios/logo_rocket_large@3x.png' :
                    delivery_kind = '로켓배송'
                elif delivery_url == 'https://image9.coupangcdn.com/image/badges/rocket-plus2/default/pc/rocket_plus_16_bi@2x.png' :
                    delivery_kind = '로켓+2일'
                elif delivery_url == 'https://image8.coupangcdn.com/image/badges/falcon/v1/web/rocketwow-bi-16@2x.png' :
                    delivery_kind = '로켓와우'
                elif delivery_url == 'https://image6.coupangcdn.com/image/badges/falcon/v1/web/rocket-fresh@2x.png' :
                    delivery_kind = '로켓프레시'
                elif delivery_url == 'https://image6.coupangcdn.com/image/delivery_badge/default/pc/global_b/global_b.png' :
                    delivery_kind = '로켓직구'
                elif delivery_url == 'https://image10.coupangcdn.com/image/delivery_badge/default/ios/rocket_merchant/consignment_v3@2x.png' :
                    delivery_kind = '제트배송'
                else :
                    delivery_kind = f'등록되지 않은 배송 방식 : {delivery_url}'

                thumbnail_img = baby_product_info.find_element(By.CSS_SELECTOR, '.image > img').get_attribute('src')
                product_name = baby_product_info.find_element(By.CSS_SELECTOR, '.name').text
                price = baby_product_info.find_element(By.CSS_SELECTOR, '.price-value').text
                product_url = baby_product_info.find_element(By.CSS_SELECTOR, '.baby-product-link').get_attribute('href')
                try :
                    if baby_product_info.find_element(By.CSS_SELECTOR, '.delivery').is_displayed() == True :
                        sold_out = '판매중'
                except :
                    sold_out = '판매중지'
                

                if delivery_kind == '로켓배송' or delivery_kind == '로켓+2일' or delivery_kind == '로켓와우':
                    if self.rocket_delivery_checkBox.isChecked() == False : 
                        continue
                elif delivery_kind == '로켓프레시' :
                    if self.rocket_fresh_checkBox.isChecked() == False :
                        continue
                elif delivery_kind == '로켓직구' :
                    if self.rocket_global_checkBox.isChecked() == False :
                        continue
                elif delivery_kind == '제트배송' :
                    if self.rocket_merchant_checkBox.isChecked() == False :
                        continue

                delivery_kind_list.append(delivery_kind)
                product_name_list.append(product_name)
                price_list.append(price)
                product_url_list.append(product_url)
                sold_out_list.append(sold_out)
                thumbnail_img_list.append(thumbnail_img)

                self.crawling_status_tableWidget.setItem(total_collect_product, 0, QTableWidgetItem(delivery_kind_list[total_collect_product]))
                self.crawling_status_tableWidget.setItem(total_collect_product, 1, QTableWidgetItem(product_name_list[total_collect_product]))
                self.crawling_status_tableWidget.setItem(total_collect_product, 2, QTableWidgetItem(price_list[total_collect_product]))
                self.crawling_status_tableWidget.setItem(total_collect_product, 3, QTableWidgetItem(sold_out_list[total_collect_product]))
                self.crawling_status_tableWidget.setItem(total_collect_product, 4, QTableWidgetItem(thumbnail_img_list[total_collect_product]))
                self.crawling_status_tableWidget.setItem(total_collect_product, 5, QTableWidgetItem(product_url_list[total_collect_product]))

                QApplication.processEvents()

                current_page_collect_product_amount += 1
                total_collect_product += 1

                if product_collect_amount == total_collect_product :
                    break

            self.crawling_status_listWidget.insertItem(-1,f'{page_num}페이지 {current_page_collect_product_amount-1}개 수집, 총 {total_collect_product}개 수집')

        return total_collect_product

    # 쿠팡 상세페이지 수집 함수.
    def cp_detailpage_collect_start(self, total_collect_product) :
        for i in range(total_collect_product) :
            self.cp_crawling_driver.get(self.crawling_status_tableWidget.item(i, 5).text())













    # 키워드 금지어 설정 함수
    def forbidden_keyword_save(self) :
        forbidden_keyword_txt = self.forbidden_keyword_plainTextEdit.toPlainText()
        if forbidden_keyword_txt == None :
            return 0
        with open(f'forbidden_keyword.txt', mode='w') as file :            # txt 파일 생성 및 저장
            file.write(forbidden_keyword_txt)
        self.forbidden_keyword_plainTextEdit.setPlainText(forbidden_keyword_txt)
        self.status_label.setText('금지어 저장 완료')

    # 키워드 금지어 읽기 함수
    def forbidden_keyword_read(self) :
        forbidden_keyword_file_existence = os.path.isfile(f'forbidden_keyword.txt')      # 라이센스 파일 존재 유무 확인.    존재할 경우 True, 비존재할경우 False 반환.

        if forbidden_keyword_file_existence == True :                              # 라이센스 파일이 존재할 경우 실행.
            with open(f'forbidden_keyword.txt', mode='r') as file :                       # txt 파일 읽어오기
                forbidden_keyword = file.read()
            self.forbidden_keyword_plainTextEdit.setPlainText(forbidden_keyword)
        elif forbidden_keyword_file_existence == False :                                       # 라이센스 파일이 존재하지 않을 경우 실행.
            return 0

QApplication.setStyle("fusion")
app = QApplication(sys.argv)
main_dialog = MainDialog()
main_dialog.show()
sys.exit(app.exec_())


# 바