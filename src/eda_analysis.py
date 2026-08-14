# -*- coding: utf-8 -*-
"""
YES24 베스트셀러 데이터셋 탐색적 데이터 분석(EDA) 파이썬 스크립트
작성 목적: 데이터 기본 정보 검증, 기술통계 산출, 11종 시각화 이미지 생성, 
           교차표/피봇테이블 생성 및 데이터 분석 결과 정리.
"""

import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import koreanize_matplotlib  # 한글 폰트 깨짐 방지용 패키지
from sklearn.feature_extraction.text import TfidfVectorizer

# ---------------------------------------------------------
# 1. 파일 경로 및 출력 디렉터리 설정 (상대 경로 준수)
# ---------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(BASE_DIR, "data", "yes24_bestseller_all.csv")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
REPORTS_DIR = os.path.join(BASE_DIR, "reports")

# 이미지 및 리포트 저장 폴더 생성 (# 설명: 디렉토리 존재 여부 확인 후 자동 생성)
os.makedirs(IMAGES_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

print(f"[정보] 데이터 로드 경로: {DATA_PATH}")

# ---------------------------------------------------------
# 2. 데이터 로드 및 전처리
# ---------------------------------------------------------
# 데이터 읽기 (# 설명: CSV 파일 읽기)
df = pd.read_csv(DATA_PATH)

# 숫자형 컬럼 변환 함수 (# 설명: 쉼표 및 문자가 포함된 숫자 전처리)
def clean_numeric(val):
    if pd.isna(val):
        return np.nan
    val_str = str(val).replace(',', '').strip()
    numbers = re.findall(r'\d+\.?\d*', val_str)
    if numbers:
        return float(numbers[0])
    return np.nan

# 판매가, 정가, 판매지수 전처리 (# 설명: 문자열 형태의 가격/지수를 숫자로 변환)
df['판매가_num'] = df['판매가'].apply(clean_numeric)
df['정가_num'] = df['정가'].apply(clean_numeric)
df['판매지수_num'] = df['판매지수'].apply(clean_numeric)
df['할인금액'] = df['정가_num'] - df['판매가_num']
df['리뷰수_num'] = df['리뷰수'].apply(clean_numeric)
df['평점_num'] = df['평점'].apply(clean_numeric)

# 출간일 전처리 (# 설명: 출간 연도 및 월 추출)
def extract_year_month(date_str):
    if pd.isna(date_str):
        return np.nan, np.nan, '미상'
    match = re.search(r'(\d{4})년\s*(\d{1,2})월', str(date_str))
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        ym = f"{year}-{month:02d}"
        return year, month, ym
    return np.nan, np.nan, '미상'

df[['출간년도', '출간월', '출간연월']] = df['출간일'].apply(extract_year_month).tolist()

# 대표 저자 추출 (# 설명: 저자 필드에서 첫번째 주요 저자 이름 정리)
def clean_author(author_str):
    if pd.isna(author_str):
        return '미상'
    author_str = str(author_str).replace('저', '').replace('글', '').strip()
    authors = re.split(r'[,·/]', author_str)
    return authors[0].strip() if authors else '미상'

df['대표저자'] = df['저자'].apply(clean_author)

print(f"[정보] 데이터 로드 완료 - 전체 행: {df.shape[0]}, 전체 열: {df.shape[1]}")

# ---------------------------------------------------------
# 3. 데이터 검증 지표 계산
# ---------------------------------------------------------
head_5 = df.head(5)
tail_5 = df.tail(5)
total_rows, total_cols = df.shape
duplicated_count = df.duplicated().sum()

# 수치형 및 범주형 기술통계 산출 (# 설명: 수치형/범주형 변수별 기술통계 표 생성)
num_cols = ['정가_num', '판매가_num', '할인율', '할인금액', '판매지수_num', '평점_num', '리뷰수_num']
cat_cols = ['출판사', '대표저자', '출간연월']

desc_num = df[num_cols].describe().T
desc_num['median'] = df[num_cols].median()
desc_num['skew'] = df[num_cols].skew()

desc_cat = df[cat_cols].describe().T

# ---------------------------------------------------------
# 4. 시각화 생성 (11종 차트, Seaborn 전역 스타일 금지 준수)
# ---------------------------------------------------------
# Chart 1: 판매지수 분포 (일변량 히스토그램)
fig, ax = plt.subplots(figsize=(10, 5))
ax.hist(df['판매지수_num'].dropna(), bins=30, color='#1f77b4', edgecolor='black', alpha=0.8)
ax.set_title("1. YES24 베스트셀러 판매지수 분포 (Univariate)", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("판매지수", fontsize=11)
ax.set_ylabel("도서 수 (권)", fontsize=11)
ax.grid(True, linestyle='--', alpha=0.5)
fig.tight_layout()
chart1_path = os.path.join(IMAGES_DIR, "01_sales_index_dist.png")
fig.savefig(chart1_path, dpi=300)
plt.close(fig)

# Chart 2: 도서 정가 및 판매가 분포 (일변량/이변량 Boxplot)
fig, ax = plt.subplots(figsize=(9, 5))
price_data = [df['정가_num'].dropna(), df['판매가_num'].dropna()]
bp = ax.boxplot(price_data, patch_artist=True, tick_labels=['정가', '판매가'])
colors = ['#ff7f0e', '#2ca02c']
for patch, color in zip(bp['boxes'], colors):
    patch.set_facecolor(color)
    patch.set_alpha(0.7)
ax.set_title("2. 도서 정가 및 판매가 분포 비교 (Bivariate Boxplot)", fontsize=14, fontweight='bold', pad=15)
ax.set_ylabel("가격 (원)", fontsize=11)
ax.grid(True, linestyle='--', alpha=0.5)
fig.tight_layout()
chart2_path = os.path.join(IMAGES_DIR, "02_price_distribution.png")
fig.savefig(chart2_path, dpi=300)
plt.close(fig)

# Chart 3: 할인율 분포 (일변량 바 차트)
fig, ax = plt.subplots(figsize=(9, 5))
disc_counts = df['할인율'].value_counts().sort_index()
ax.bar(disc_counts.index.astype(str) + '%', disc_counts.values, color='#9467bd', edgecolor='black', alpha=0.8)
ax.set_title("3. 도서 할인율 빈도 분포 (Univariate Bar Chart)", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("할인율 (%)", fontsize=11)
ax.set_ylabel("도서 수 (권)", fontsize=11)
ax.grid(True, linestyle='--', alpha=0.5)
fig.tight_layout()
chart3_path = os.path.join(IMAGES_DIR, "03_discount_rate_dist.png")
fig.savefig(chart3_path, dpi=300)
plt.close(fig)

# Chart 4: 상위 30개 출판사 빈도 (범주형 일변량)
fig, ax = plt.subplots(figsize=(10, 8))
top30_pub = df['출판사'].value_counts().head(30).sort_values(ascending=True)
ax.barh(top30_pub.index, top30_pub.values, color='#8c564b', edgecolor='black', alpha=0.85)
ax.set_title("4. 베스트셀러 도서 출판사 상위 30개 (Categorical Top 30)", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("출간 권수 (권)", fontsize=11)
ax.grid(True, linestyle='--', alpha=0.5)
fig.tight_layout()
chart4_path = os.path.join(IMAGES_DIR, "04_top30_publishers.png")
fig.savefig(chart4_path, dpi=300)
plt.close(fig)

# Chart 5: 상위 30개 저자 빈도 (범주형 일변량)
fig, ax = plt.subplots(figsize=(10, 8))
top30_author = df['대표저자'].value_counts().head(30).sort_values(ascending=True)
ax.barh(top30_author.index, top30_author.values, color='#e377c2', edgecolor='black', alpha=0.85)
ax.set_title("5. 베스트셀러 저자 상위 30명 (Categorical Top 30)", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("출간 권수 (권)", fontsize=11)
ax.grid(True, linestyle='--', alpha=0.5)
fig.tight_layout()
chart5_path = os.path.join(IMAGES_DIR, "05_top30_authors.png")
fig.savefig(chart5_path, dpi=300)
plt.close(fig)

# Chart 6: 출판사별 평균 판매지수 상위 15개 (이변량)
fig, ax = plt.subplots(figsize=(10, 6))
pub_sales = df.groupby('출판사')['판매지수_num'].agg(['mean', 'count'])
pub_sales_filter = pub_sales[pub_sales['count'] >= 2].sort_values(by='mean', ascending=False).head(15)
ax.barh(pub_sales_filter.index[::-1], pub_sales_filter['mean'][::-1], color='#bcbd22', edgecolor='black', alpha=0.85)
ax.set_title("6. 출판사별 평균 판매지수 상위 15개 (최소 2권 이상 출간)", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("평균 판매지수", fontsize=11)
ax.grid(True, linestyle='--', alpha=0.5)
fig.tight_layout()
chart6_path = os.path.join(IMAGES_DIR, "06_publisher_avg_sales_index.png")
fig.savefig(chart6_path, dpi=300)
plt.close(fig)

# Chart 7: 정가 vs 판매지수 관계 (이변량 산점도)
fig, ax = plt.subplots(figsize=(9, 6))
ax.scatter(df['정가_num'], df['판매지수_num'], color='#17becf', alpha=0.6, edgecolors='none')
ax.set_title("7. 도서 정가와 판매지수 간의 상관관계 (Bivariate Scatter)", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("정가 (원)", fontsize=11)
ax.set_ylabel("판매지수", fontsize=11)
ax.grid(True, linestyle='--', alpha=0.5)
fig.tight_layout()
chart7_path = os.path.join(IMAGES_DIR, "07_price_vs_sales_index.png")
fig.savefig(chart7_path, dpi=300)
plt.close(fig)

# Chart 8: 평점 vs 리뷰수 관계 (이변량 산점도)
fig, ax = plt.subplots(figsize=(9, 6))
ax.scatter(df['평점_num'], df['리뷰수_num'], color='#d62728', alpha=0.6, edgecolors='k', linewidths=0.5)
ax.set_title("8. 도서 평점과 리뷰수 간의 분포 (Bivariate Scatter)", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("평점 (10점 만점)", fontsize=11)
ax.set_ylabel("리뷰수 (건)", fontsize=11)
ax.grid(True, linestyle='--', alpha=0.5)
fig.tight_layout()
chart8_path = os.path.join(IMAGES_DIR, "08_rating_vs_reviews.png")
fig.savefig(chart8_path, dpi=300)
plt.close(fig)

# Chart 9: 출간연월별 베스트셀러 출간권수 추이 (시간 시계열 이변량)
fig, ax = plt.subplots(figsize=(11, 5))
ym_counts = df[df['출간연월'] != '미상']['출간연월'].value_counts().sort_index()
ax.plot(ym_counts.index, ym_counts.values, marker='o', color='#2ca02c', linewidth=2, markersize=6)
ax.set_title("9. 출간연월별 베스트셀러 등록 추이 (Time Series Trend)", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("출간연월", fontsize=11)
ax.set_ylabel("등록 도서 수 (권)", fontsize=11)
plt.xticks(rotation=45)
ax.grid(True, linestyle='--', alpha=0.5)
fig.tight_layout()
chart9_path = os.path.join(IMAGES_DIR, "09_pub_date_trend.png")
fig.savefig(chart9_path, dpi=300)
plt.close(fig)

# Chart 10: 수치형 변수 상관관계 히트맵 (다변량)
fig, ax = plt.subplots(figsize=(8, 7))
corr = df[['정가_num', '판매가_num', '할인율', '판매지수_num', '평점_num', '리뷰수_num']].corr()
cax = ax.matshow(corr, cmap='coolwarm', vmin=-1, vmax=1)
fig.colorbar(cax)
ticks = np.arange(len(corr.columns))
ax.set_xticks(ticks)
ax.set_yticks(ticks)
ax.set_xticklabels(['정가', '판매가', '할인율', '판매지수', '평점', '리뷰수'], rotation=45, ha='left')
ax.set_yticklabels(['정가', '판매가', '할인율', '판매지수', '평점', '리뷰수'])
for i in range(len(corr.columns)):
    for j in range(len(corr.columns)):
        ax.text(j, i, f"{corr.iloc[i, j]:.2f}", ha='center', va='center', color='black', fontsize=10)
ax.set_title("10. 수치형 변수 상관관계 히트맵 (Multivariate Correlation)", fontsize=14, fontweight='bold', pad=25)
fig.tight_layout()
chart10_path = os.path.join(IMAGES_DIR, "10_correlation_heatmap.png")
fig.savefig(chart10_path, dpi=300)
plt.close(fig)

# Chart 11: 텍스트 TF-IDF 키워드 상위 30개 (텍스트 키워드 분석)
text_corpus = (df['도서명'].fillna('') + ' ' + df['부제목'].fillna('')).tolist()
# 한글 및 영문 단어 단위 토큰화 (# 설명: 형태소 분석기 없이 정규식 단어 토큰화 후 TF-IDF 도출)
vectorizer = TfidfVectorizer(token_pattern=r'(?u)\b\w\w+\b', max_features=100)
tfidf_matrix = vectorizer.fit_transform(text_corpus)
feature_names = vectorizer.get_feature_names_out()
tfidf_scores = tfidf_matrix.sum(axis=0).A1

tfidf_df = pd.DataFrame({'keyword': feature_names, 'score': tfidf_scores})
tfidf_top30 = tfidf_df.sort_values(by='score', ascending=False).head(30)

fig, ax = plt.subplots(figsize=(10, 8))
ax.barh(tfidf_top30['keyword'][::-1], tfidf_top30['score'][::-1], color='#393b79', alpha=0.85, edgecolor='black')
ax.set_title("11. 도서명 및 부제목 TF-IDF 중요 키워드 상위 30개", fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel("TF-IDF 중요도 총점", fontsize=11)
ax.grid(True, linestyle='--', alpha=0.5)
fig.tight_layout()
chart11_path = os.path.join(IMAGES_DIR, "11_tfidf_top30_keywords.png")
fig.savefig(chart11_path, dpi=300)
plt.close(fig)

print("[성공] 11개 시각화 이미지 및 전처리 결과 계산 완료!")
