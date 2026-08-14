"""
YES24 베스트셀러 전체 데이터 수집 스크립트

이 모듈은 YES24 웹사이트의 베스트셀러 카테고리 AJAX API(BestSellerContents)를 활용하여
1페이지부터 마지막 페이지까지 전체 도서 목록을 수집하고 CSV 파일로 저장합니다.

주요 특징:
- 1페이지부터 데이터가 더 이상 나오지 않을 때(마지막 페이지)까지 자동 반복 수집
- 서버 부하 방지 및 차단 회피를 위해 요청 간 0.1초~0.5초 사이의 랜덤 딜레이(지연시간) 적용
- 결과를 CSV 파일(yes24/data/yes24_bestseller_all.csv)로 저장
"""

import os
import re
import json
import csv
import time
import random
import requests
from bs4 import BeautifulSoup


def scrape_yes24_bestseller_page(page_number=1, category_number="001001003"):
    """
    YES24 베스트셀러 특정 페이지의 상품 목록 데이터를 스크래핑합니다.
    
    :param page_number: 수집할 페이지 번호
    :param category_number: 카테고리 번호 (기본값: 001001003 - IT 모바일)
    :return: 해당 페이지의 상품 정보 딕셔너리 리스트
    """
    # 1. 요청 URL 지정
    url = "https://www.yes24.com/product/category/BestSellerContents"

    # 2. HTTP 요청 파라미터 설정 (scraping_prompt.md 기반)
    params = {
        "categoryNumber": category_number,
        "sumGb": "06",
        "sex": "A",
        "age": "255",
        "goodsTp": "0",
        "addOptionTp": "0",
        "excludeTp": "2",
        "pageNumber": str(page_number),
        "pageSize": "24",
        "goodsStatGb": "06",
        "eBookTp": "0",
        "bestType": "YES24_BESTSELLER",
        "type": "",
        "saleYear": "0",
        "saleMonth": "0",
        "weekNo": "0",
        "saleDts": "",
        "viewMode": "",
        "freeYn": ""
    }

    # 3. HTTP 요청 헤더 설정 (차단 방지 및 서버 응답 유도)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36",
        "Referer": f"https://www.yes24.com/product/category/bestseller?categoryNumber={category_number}&pageNumber={page_number}&pageSize=24",
        "X-Requested-With": "XMLHttpRequest"
    }

    # 4. HTTP GET 요청 보내기
    try:
        response = requests.get(url, headers=headers, params=params, timeout=10)
        if response.status_code != 200:
            print(f"[!] {page_number}페이지 요청 실패 (상태 코드: {response.status_code})")
            return []
    except Exception as e:
        print(f"[!] {page_number}페이지 요청 중 에러 발생: {e}")
        return []

    # 5. BeautifulSoup을 이용한 HTML 파싱
    soup = BeautifulSoup(response.text, "html.parser")
    
    # data-goods-no 속성을 가진 li 요소 추출 (상품 항목)
    goods_lis = [li for li in soup.find_all("li") if li.has_attr("data-goods-no")]

    if not goods_lis:
        # 더 이상 수집할 상품 목록이 없는 경우
        return []

    book_list = []

    # 6. 각 상품 요소에서 상세 정보 추출
    for li in goods_lis:
        goods_no = li.get("data-goods-no", "")
        
        # 순위 추출
        rank_elem = li.select_one("em.ico.rank")
        rank = rank_elem.get_text(strip=True) if rank_elem else ""

        # 도서명 추출
        title_elem = li.select_one("a.gd_name")
        title = title_elem.get_text(strip=True) if title_elem else ""

        # 상세페이지 링크 추출
        link = f"https://www.yes24.com{title_elem['href']}" if title_elem and title_elem.has_attr("href") else ""

        # 부제목 / 설명 추출
        sub_title_elem = li.select_one("span.gd_nameE")
        sub_title = sub_title_elem.get_text(strip=True) if sub_title_elem else ""

        # 이미지 URL 추출
        img_elem = li.select_one("img.lazy") or li.select_one("img")
        img_url = ""
        if img_elem:
            img_url = img_elem.get("data-original") or img_elem.get("src") or ""

        # 저자 추출
        author_elem = li.select_one("span.info_auth")
        author = author_elem.get_text(strip=True) if author_elem else ""

        # 출판사 추출
        pub_elem = li.select_one("span.info_pub")
        pub = pub_elem.get_text(strip=True) if pub_elem else ""

        # 출간일 추출
        date_elem = li.select_one("span.info_date")
        pub_date = date_elem.get_text(strip=True) if date_elem else ""

        # 할인율 추출
        discount_elem = li.select_one("span.txt_sale em.num")
        discount_rate = discount_elem.get_text(strip=True) if discount_elem else ""

        # 판매가 추출
        price_elem = li.select_one("strong.txt_num em.yes_b")
        sale_price = price_elem.get_text(strip=True) if price_elem else ""

        # 정가 추출
        orig_price_elem = li.select_one("span.txt_num.dash em.yes_m")
        orig_price = orig_price_elem.get_text(strip=True) if orig_price_elem else ""

        # 판매지수 추출
        sale_num_elem = li.select_one("span.saleNum")
        sale_index = sale_num_elem.get_text(strip=True) if sale_num_elem else ""

        # 평점 추출
        rating_elem = li.select_one("span.rating_grade em.yes_b")
        rating = rating_elem.get_text(strip=True) if rating_elem else ""

        # 리뷰 수 추출
        rv_count_elem = li.select_one("span.rating_rvCount em.txC_blue")
        rv_count = rv_count_elem.get_text(strip=True) if rv_count_elem else ""

        # 데이터 딕셔너리 생성
        book_info = {
            "페이지": page_number,
            "순위": rank,
            "상품번호": goods_no,
            "도서명": title,
            "부제목": sub_title,
            "저자": author,
            "출판사": pub,
            "출간일": pub_date,
            "할인율": discount_rate,
            "판매가": sale_price,
            "정가": orig_price,
            "판매지수": sale_index,
            "평점": rating,
            "리뷰수": rv_count,
            "상세링크": link,
            "이미지URL": img_url
        }

        book_list.append(book_info)

    return book_list


def scrape_all_pages(category_number="001001003", start_page=1, max_pages=500):
    """
    1페이지부터 마지막 페이지까지 연속으로 데이터 수집을 진행합니다.
    페이지 요청 간에 0.1~0.5초 사이의 랜덤 딜레이를 줍니다.
    """
    all_books = []
    current_page = start_page

    print(f"[*] 1페이지부터 마지막 페이지까지 데이터 수집을 시작합니다. (카테고리: {category_number})")

    while current_page <= max_pages:
        print(f"[*] {current_page}페이지 수집 중...", end="", flush=True)
        
        books = scrape_yes24_bestseller_page(page_number=current_page, category_number=category_number)
        
        # 상품 데이터가 없으면 마지막 페이지에 도달한 것이므로 종료
        if not books:
            print("\n[*] 더 이상 수집할 데이터가 없습니다. 수집을 종료합니다.")
            break
            
        all_books.extend(books)
        print(f" 완료 ({len(books)}개 수집, 누적: {len(all_books)}개)")

        # 0.1초 ~ 0.5초 사이의 랜덤 지연 시간(Sleep) 적용 (서버 차단 방지)
        sleep_time = random.uniform(0.1, 0.5)
        time.sleep(sleep_time)

        current_page += 1

    return all_books


def save_to_csv(data, file_path):
    """
    수집한 전체 데이터를 CSV 파일로 저장합니다.
    """
    if not data:
        print("[!] 저장할 데이터가 없습니다.")
        return

    # 저장 경로 디렉터리 확인 및 생성
    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    # UTF-8-SIG 인코딩으로 저장하여 엑셀(Excel)에서 한글 깨짐 방지
    headers = list(data[0].keys())
    with open(file_path, mode="w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(data)

    print(f"[+] 성공적으로 CSV 파일 저장 완료: {file_path}")


if __name__ == "__main__":
    # 전체 페이지 수집 실행
    all_scraped_data = scrape_all_pages(category_number="001001003")

    # 기존 1페이지 수집 CSV 경로 및 전체 데이터 CSV 경로 설정
    data_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "data"))
    
    # 1)기존 csv 파일(yes24_bestseller_page1.csv)에 전체 수집 데이터 저장
    existing_csv = os.path.join(data_dir, "yes24_bestseller_page1.csv")
    save_to_csv(all_scraped_data, existing_csv)

    # 2) 전체 데이터 파일(yes24_bestseller_all.csv)로도 저장
    all_csv = os.path.join(data_dir, "yes24_bestseller_all.csv")
    save_to_csv(all_scraped_data, all_csv)
