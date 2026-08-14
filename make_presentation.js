// YES24 베스트셀러 EDA 리포트 기반 30페이지 PPTX 슬라이드 자동 생성 스크립트 (Nordic Minimalism Style)
// pptxgenjs와 sharp(SVG 플랫 아이콘 렌더러)를 활용하여 벽돌색(Terracotta) 포인트와 노르딕 감성을 적용합니다.

const PptxGenJS = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const sharp = require("sharp");

// PPTX 객체 생성
const pptx = new PptxGenJS();

// 1. 슬라이드 레이아웃 설정 (16:9 Wide 레이아웃: 13.333인치 x 7.5인치)
pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "CUSTOM_WIDE";

// 2. 폰트 및 노르딕 미니멀리즘 (Nordic Minimalism) & 벽돌색(Terracotta) 컬러 팔레트
const FONT_TITLE = "Gmarket Sans Bold"; // 유저 지정 제목 폰트
const FONT_BODY = "NanumGothic";        // 유저 지정 내용 폰트

// 노르딕 스타일 색상 토큰 (6자리 HEX, # 제거)
const C_NORDIC_BG = "F4F1EC";     // 웜 크림 (노르딕 슬라이드 메인 배경)
const C_DARK_TEXT = "3D3530";     // 딥 웜 차콜 (주요 텍스트)
const C_MUTED_TEXT = "7A6C5D";    // 토프 / 뮤티드 브라운 (서브 텍스트)
const C_BRICK = "B85042";         // 유저 지정 포인트 컬러: 벽돌색 (Terracotta Red)
const C_BRICK_DARK = "963B2F";    // 딥 벽돌색 (강조 뱃지 / 헤더)
const C_BRICK_BG = "FDF2F0";      // 연한 벽돌색 틴트 배경 (포인트 카드)
const C_SAND_BG = "EAE3D9";       // 노르딕 샌드 배경
const C_CARD_BG = "FFFFFF";       // 화이트 카드 컨테이너
const C_BORDER = "E2DACD";        // 카드 테두리 선 색상

// 이미지 디렉토리 경로
const IMG_DIR = path.join(__dirname, "images");

// 3. Sharp 기반 플랫 아이콘(Flat SVG -> PNG Base64 Data URI) 생성 함수
const svgTemplates = {
    book: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-1 7h-6v2h6v-2zm0-4h-6v2h6V6zm0 8h-6v2h6v-2z"/></svg>`,
    chart: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M5 9.2h3V19H5zM10.6 5h2.8v14h-2.8zM16.2 13H19v6h-2.8z"/></svg>`,
    price: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 14h-2v-1c-1.66 0-3-1.34-3-3h2c0 .55.45 1 1 1h2c.55 0 1-.45 1-1s-.45-1-1-1h-2c-1.66 0-3-1.34-3-3s1.34-3 3-3v-1h2v1c1.66 0 3 1.34 3 3h-2c0-.55-.45-1-1-1h-2c-.55 0-1 .45-1 1s.45 1 1 1h2c1.66 0 3 1.34 3 3s-1.34 3-3 3v1z"/></svg>`,
    discount: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M21.41 11.58l-9-9C12.05 2.22 11.55 2 11 2H4c-1.1 0-2 .9-2 2v7c0 .55.22 1.05.59 1.42l9 9c.36.36.86.58 1.41.58.55 0 1.05-.22 1.41-.59l7-7c.37-.36.59-.86.59-1.41 0-.55-.23-1.06-.59-1.42zM5.5 7C4.67 7 4 6.33 4 5.5S4.67 4 5.5 4 7 4.67 7 5.5 6.33 7 5.5 7z"/></svg>`,
    publisher: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M12 7V3H2v18h20V7H12zM6 19H4v-2h2v2zm0-4H4v-2h2v2zm0-4H4V9h2v2zm0-4H4V5h2v2zm4 12H8v-2h2v2zm0-4H8v-2h2v2zm0-4H8V9h2v2zm0-4H8V5h2v2zm10 12h-8v-2h2v-2h-2v-2h2v-2h-2V9h8v10zm-2-8h-2v2h2v-2zm0 4h-2v2h2v-2z"/></svg>`,
    author: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>`,
    star: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M12 17.27L18.18 21l-1.64-7.03L22 9.24l-7.19-.61L12 2 9.19 8.63 2 9.24l5.46 4.73L5.82 21z"/></svg>`,
    review: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M20 2H4c-1.1 0-1.99.9-1.99 2L2 22l4-4h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zM6 9h12v2H6V9zm8 5H6v-2h8v2zm4-6H6V6h12v2z"/></svg>`,
    ai: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M21 11.5v-1c0-.83-.67-1.5-1.5-1.5H18V7c0-1.1-.9-2-2-2h-2V3.5C14 2.67 13.33 2 12.5 2h-1C10.67 2 10 2.67 10 3.5V5H8c-1.1 0-2 .9-2 2v2H4.5C3.67 9 3 9.67 3 10.5v1c0 .83.67 1.5 1.5 1.5H6v2H4.5c-.83 0-1.5.67-1.5 1.5v1c0 .83.67 1.5 1.5 1.5H6v2c0 1.1.9 2 2 2h2v1.5c0 .83.67 1.5 1.5 1.5h1c.83 0 1.5-.67 1.5-1.5V19h2c1.1 0 2-.9 2-2v-2h1.5c.83 0 1.5-.67 1.5-1.5v-1c0-.83-.67-1.5-1.5-1.5H18v-2h1.5c.83 0 1.5-.67 1.5-1.5zM12 16c-2.21 0-4-1.79-4-4s1.79-4 4-4 4 1.79 4 4-1.79 4-4 4z"/></svg>`,
    trend: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M16 6l2.29 2.29-4.88 4.88-4-4L2 16.59 3.41 18l6-6 4 4 6.3-6.29L22 12V6z"/></svg>`,
    target: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.42 0-8-3.58-8-8s3.58-8 8-8 8 3.58 8 8-3.58 8-8 8zm0-14c-3.31 0-6 2.69-6 6s2.69 6 6 6 6-2.69 6-6-2.69-6-6-6zm0 10c-2.21 0-4-1.79-4-4s1.79-4 4-4 4 1.79 4 4-1.79 4-4 4z"/></svg>`,
    checklist: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#B85042"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-9 14l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/></svg>`
};

const flatIconBase64Map = {};

// 동기 방식으로 SVG 아이콘들을 PNG Base64로 미리 변환
async function initIcons() {
    for (const [name, svgStr] of Object.entries(svgTemplates)) {
        const buf = await sharp(Buffer.from(svgStr)).resize(128, 128).png().toBuffer();
        flatIconBase64Map[name] = "image/png;base64," + buf.toString("base64");
    }
}

// 4. 슬라이드 공통 템플릿 함수 (Nordic Style)
function createNordicSlide(titleText, categoryText = "YES24 Bestseller EDA | Nordic Series") {
    const slide = pptx.addSlide();
    
    // 노르딕 웜 크림 배경 설정
    slide.background = { color: C_NORDIC_BG };

    // 노르딕 시그니처 3-Dot Accent motif (상단 우측)
    const dotColors = [C_BRICK, "7A6C5D", "D9CFC4"];
    dotColors.forEach((color, i) => {
        slide.addShape(pptx.shapes.OVAL, {
            x: 12.2 + i * 0.25, y: 0.45, w: 0.15, h: 0.15,
            fill: { color: color }
        });
    });

    // 상단 소형 뱃지 (벽돌색 포인트)
    slide.addText(categoryText.toUpperCase(), {
        x: 0.8, y: 0.4, w: 10.5, h: 0.3,
        fontFace: FONT_BODY, fontSize: 9.5, color: C_BRICK, bold: true, margin: 0
    });

    // 메인 슬라이드 제목 (G마켓 산스 Bold, 노르딕 딥 차콜)
    slide.addText(titleText, {
        x: 0.8, y: 0.7, w: 10.5, h: 0.6,
        fontFace: FONT_TITLE, fontSize: 22, color: C_DARK_TEXT, bold: true, margin: 0
    });

    // 하단 노르딕 세로/가로 미니멀 라인 및 푸터
    slide.addShape(pptx.shapes.LINE, {
        x: 0.8, y: 7.0, w: 11.733, h: 0,
        line: { color: C_BORDER, width: 1 }
    });

    slide.addText("YES24 Bestseller EDA Report | Nordic Minimalist & Terracotta Edition", {
        x: 0.8, y: 7.05, w: 8.0, h: 0.3,
        fontFace: FONT_BODY, fontSize: 9, color: C_MUTED_TEXT, margin: 0
    });

    return slide;
}

// 스크립트 메인 실행 함수
async function buildPresentation() {
    await initIcons();

    // =========================================================================
    // SLIDE 1: 타이틀 슬라이드 (Nordic Brick Accent Theme)
    // =========================================================================
    {
        const slide = pptx.addSlide();
        slide.background = { color: C_NORDIC_BG };

        // 노르딕 시그니처 대형 웜 샌드 블록 카드
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 0.8, y: 0.8, w: 11.733, h: 5.8, rectRadius: 0.08,
            fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1.5 }
        });

        // 좌측 벽돌색 Accent 라인 기둥
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 0.8, y: 0.8, w: 0.25, h: 5.8, rectRadius: 0.0,
            fill: { color: C_BRICK }
        });

        // 메인 타이틀 & 부제목
        slide.addText("YES24 도서 베스트셀러 1,000건\n심층 EDA & 비즈니스 전략 보고서", {
            x: 1.4, y: 1.4, w: 10.5, h: 1.8,
            fontFace: FONT_TITLE, fontSize: 32, color: C_DARK_TEXT, bold: true, align: "left"
        });

        slide.addText("20년차 데이터 분석가의 노르딕 미니멀리즘 접근법: 시장 구조, 정가 할인, 파레토 흥행 및 AI 트렌드", {
            x: 1.4, y: 3.2, w: 10.5, h: 0.8,
            fontFace: FONT_BODY, fontSize: 15, color: C_MUTED_TEXT, align: "left"
        });

        // 핵심 지표 3개 플랫 아이콘 카드
        const metrics = [
            { icon: "book", label: "분석 표본 데이터", val: "1,000건" },
            { icon: "chart", label: "특화 시각화 분석", val: "11종 차트" },
            { icon: "target", label: "분석 변수 수", val: "26개 컬럼" }
        ];

        metrics.forEach((m, idx) => {
            const boxX = 1.4 + idx * 3.6;
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: boxX, y: 4.3, w: 3.3, h: 1.3, rectRadius: 0.08,
                fill: { color: C_BRICK_BG }, line: { color: C_BRICK, width: 1 }
            });

            if (flatIconBase64Map[m.icon]) {
                slide.addImage({ data: flatIconBase64Map[m.icon], x: boxX + 0.2, y: 4.5, w: 0.45, h: 0.45 });
            }

            slide.addText(m.label, {
                x: boxX + 0.75, y: 4.45, w: 2.4, h: 0.3,
                fontFace: FONT_BODY, fontSize: 11, color: C_MUTED_TEXT
            });
            slide.addText(m.val, {
                x: boxX + 0.75, y: 4.75, w: 2.4, h: 0.5,
                fontFace: FONT_TITLE, fontSize: 18, color: C_BRICK_DARK, bold: true
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
안녕하십니까, 오늘 발표를 맡은 20년차 데이터 분석가입니다. 
지금부터 대한민국 대표 온라인 서점인 YES24의 베스트셀러 도서 1,000건을 대상으로 진행된 '심층 탐색적 데이터 분석(EDA) 및 비즈니스 전략 보고서' 발표를 시작하도록 하겠습니다.

이번 보고서는 절제되고 자연스러운 노르딕 미니멀리즘(Nordic Minimalism) 디자인 언어와 따뜻한 벽돌색(Terracotta Red) 포인트를 적용하여, 데이터의 핵심 구조를 시각적으로 직관적이게 이해할 수 있도록 재설계되었습니다.

저희가 분석한 데이터셋은 총 1,000개 도서 표본과 26개 변수로 구성되어 있으며, 데이터 무결성 100%를 검증했습니다. 발표 진행 동안 11종의 시각화 차트와 플랫 아이콘을 활용하여 출판사와 유통사에 즉각적인 가치를 전달하겠습니다.`);
    }

    // =========================================================================
    // SLIDE 2: 목차 (Executive Summary & Agenda)
    // =========================================================================
    {
        const slide = createNordicSlide("Executive Summary & 주요 발표 아젠다");

        const agendaItems = [
            { icon: "checklist", num: "01", title: "분석 개요 및 데이터 검증", desc: "1,000행 x 26열 데이터 구조, 데이터 무결성 100% 확보, 결측치 원인 정의" },
            { icon: "chart", num: "02", title: "수치형 & 범주형 기술통계", desc: "가격 구조, 롱테일 흥행 비대칭성, 독자 평점/리뷰 동역학, 출판사 독과점" },
            { icon: "layers", num: "03", title: "11종 특화 차트 심층 해석", desc: "판매지수, Boxplot, 할인율, 상위 30 출판사/저자, 시계열, 상관관계, TF-IDF" },
            { icon: "target", num: "04", title: "종합 인사이트 & Action Plan", desc: "20년차 분석가의 4대 핵심 파인딩, 출판사/유통사 3단계 실행 로드맵" }
        ];

        agendaItems.forEach((item, idx) => {
            const row = Math.floor(idx / 2);
            const col = idx % 2;
            const cardX = 0.8 + col * 5.9;
            const cardY = 1.6 + row * 2.5;

            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: cardX, y: cardY, w: 5.633, h: 2.2, rectRadius: 0.08,
                fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
            });

            // 아이콘 뱃지
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: cardX + 0.3, y: cardY + 0.3, w: 0.8, h: 0.8, rectRadius: 0.08,
                fill: { color: C_BRICK_BG }, line: { color: C_BRICK, width: 1 }
            });

            if (flatIconBase64Map[item.icon]) {
                slide.addImage({ data: flatIconBase64Map[item.icon], x: cardX + 0.45, y: cardY + 0.45, w: 0.5, h: 0.5 });
            }

            slide.addText(item.title, {
                x: cardX + 1.3, y: cardY + 0.3, w: 4.0, h: 0.4,
                fontFace: FONT_TITLE, fontSize: 16, color: C_DARK_TEXT, bold: true
            });
            slide.addText(item.desc, {
                x: cardX + 1.3, y: cardY + 0.8, w: 4.0, h: 1.1,
                fontFace: FONT_BODY, fontSize: 12.5, color: C_MUTED_TEXT, lineSpacing: 18
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
오늘 발표는 크게 4개의 핵심 섹션으로 나누어 진행됩니다.

첫 번째 섹션은 데이터 기초 검증 단계로서, YES24 베스트셀러 1,000건의 데이터 크기와 차원, 중복 0건의 무결성 검증, 수집된 26개 변수의 구조 및 결측치 원인을 다룹니다.

두 번째 섹션은 기술통계 분석으로 정가/판매가의 사분위 범위, 판매지수의 롱테일 비대칭성, 독자 평점/리뷰수의 상호작용, 출판사 과점 현상을 살펴봅니다.

세 번째 섹션은 11종 특화 차트의 심층 해석입니다. 히스토그램, 박스플롯, Bar 차트, 시계열, 상관관계 히트맵, TF-IDF 텍스트 키워드 차트를 세밀하게 분석합니다.

네 번째 섹션은 4대 핵심 결론과 출판사/유통사 3단계 전략 로드맵을 제언해 드리겠습니다.`);
    }

    // =========================================================================
    // SLIDE 3: [1. 개요] 데이터셋 크기 및 데이터 무결성 검증
    // =========================================================================
    {
        const slide = createNordicSlide("1.1 데이터 크기, 차원 & 무결성 100% 검증");

        const kpiData = [
            { icon: "book", label: "전체 데이터 행 수 (Rows)", val: "1,000 건", sub: "YES24 베스트셀러 표본" },
            { icon: "layers", label: "전체 변수 열 수 (Columns)", val: "26 개", sub: "원시 16개 + 파생 10개" },
            { icon: "checklist", label: "중복 데이터 (Duplicates)", val: "0 건 (0.0%)", sub: "무결성 100% 확보" }
        ];

        kpiData.forEach((kpi, idx) => {
            const yPos = 1.6 + idx * 1.7;
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: 0.8, y: yPos, w: 4.8, h: 1.5, rectRadius: 0.08,
                fill: { color: idx === 2 ? C_BRICK_BG : C_CARD_BG },
                line: { color: idx === 2 ? C_BRICK : C_BORDER, width: 1.5 }
            });

            if (flatIconBase64Map[kpi.icon]) {
                slide.addImage({ data: flatIconBase64Map[kpi.icon], x: 1.1, y: yPos + 0.35, w: 0.5, h: 0.5 });
            }

            slide.addText(kpi.label, {
                x: 1.7, y: yPos + 0.2, w: 3.7, h: 0.3,
                fontFace: FONT_BODY, fontSize: 11.5, color: C_MUTED_TEXT
            });
            slide.addText(kpi.val, {
                x: 1.7, y: yPos + 0.5, w: 3.7, h: 0.5,
                fontFace: FONT_TITLE, fontSize: 20, color: idx === 2 ? C_BRICK_DARK : C_DARK_TEXT, bold: true
            });
            slide.addText(kpi.sub, {
                x: 1.7, y: yPos + 1.05, w: 3.7, h: 0.3,
                fontFace: FONT_BODY, fontSize: 11, color: C_DARK_TEXT
            });
        });

        // 우측 설명 카드
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 5.9, y: 1.6, w: 6.633, h: 4.9, rectRadius: 0.08,
            fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
        });

        slide.addText("데이터 수집 & 정제 세부 프로세스", {
            x: 6.2, y: 1.9, w: 6.0, h: 0.4,
            fontFace: FONT_TITLE, fontSize: 17, color: C_DARK_TEXT, bold: true
        });

        const processSteps = [
            { title: "1. 수집 무결성 검증", desc: "상품번호 및 URL 식별자 검증 결과 중복 레코드가 전혀 없는 완전성 확보" },
            { title: "2. 수치 데이터 정형화", desc: "원화 쉼표, '원', '%' 기호 제거 후 float64 정밀 변수로 가공" },
            { title: "3. 날짜 변수 표준화", desc: "'YYYY년 MM월' 출간일을 YYYY-MM 시계열 카테고리로 정규화" },
            { title: "4. 분석 신뢰성 확보", desc: "1,000건 표본은 국내 온라인 서점 유통 시장의 통계적 대표성을 충분히 상회함" }
        ];

        processSteps.forEach((step, idx) => {
            const stepY = 2.4 + idx * 1.0;
            slide.addShape(pptx.shapes.OVAL, {
                x: 6.2, y: stepY + 0.08, w: 0.25, h: 0.25,
                fill: { color: C_BRICK }
            });
            slide.addText(step.title, {
                x: 6.6, y: stepY, w: 5.6, h: 0.3,
                fontFace: FONT_TITLE, fontSize: 13.5, color: C_DARK_TEXT, bold: true
            });
            slide.addText(step.desc, {
                x: 6.6, y: stepY + 0.3, w: 5.6, h: 0.6,
                fontFace: FONT_BODY, fontSize: 11.5, color: C_MUTED_TEXT
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
1.1 절에서는 데이터 분석의 가장 기본인 데이터의 크기와 차원, 데이터 무결성 검증 결과를 말씀드립니다.

분석 데이터셋은 총 1,000개의 행(Rows)과 26개의 열(Columns)로 구성되어 있습니다. 식별자 정밀 검증 결과 중복 데이터 0건, 데이터 무결성 100%를 확보했습니다.

우측 카드를 보시면 수집된 원시 데이터에서 쉼표나 '원' 등의 단위를 정제하고, float64 수치형 변수 10개를 추가 파생시켰으며, 출간일 역시 YYYY-MM 형태로 표준화하여 시계열 분석의 기반을 마련했습니다.`);
    }

    // =========================================================================
    // SLIDE 4: [1. 개요] 데이터 컬럼 정보 및 결측치 현황
    // =========================================================================
    {
        const slide = createNordicSlide("1.2 데이터 컬럼 구조 & 결측치 메커니즘");

        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 0.8, y: 1.6, w: 5.6, h: 4.9, rectRadius: 0.08,
            fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
        });

        slide.addText("데이터 변수 분류 체계 (26개 변수)", {
            x: 1.1, y: 1.9, w: 5.0, h: 0.4,
            fontFace: FONT_TITLE, fontSize: 16, color: C_DARK_TEXT, bold: true
        });

        const colGroups = [
            { cat: "원시 식별자 & 기본정보 (8개)", items: "페이지, 순위, 상품번호, 도서명, 부제목, 저자, 출판사, 출간일" },
            { cat: "가격 & 유통 수치 변수 (7개)", items: "할인율, 판매가, 정가, 판매가_num, 정가_num, 할인금액, 판매지수_num" },
            { cat: "독자 반응 & 평가 변수 (4개)", items: "평점, 리뷰수, 평점_num, 리뷰수_num" },
            { cat: "파생 시계열 & 텍스트 (7개)", items: "출간년도, 출간월, 출간연월, 대표저자, 상세링크, 이미지URL 등" }
        ];

        colGroups.forEach((cg, idx) => {
            const gY = 2.4 + idx * 1.0;
            slide.addText(cg.cat, {
                x: 1.1, y: gY, w: 5.0, h: 0.3,
                fontFace: FONT_TITLE, fontSize: 13, color: C_BRICK, bold: true
            });
            slide.addText(cg.items, {
                x: 1.1, y: gY + 0.25, w: 5.0, h: 0.6,
                fontFace: FONT_BODY, fontSize: 11, color: C_DARK_TEXT
            });
        });

        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 6.7, y: 1.6, w: 5.833, h: 4.9, rectRadius: 0.08,
            fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
        });

        slide.addText("결측치(Null Value) 발생 원인 분석", {
            x: 7.0, y: 1.9, w: 5.2, h: 0.4,
            fontFace: FONT_TITLE, fontSize: 16, color: C_DARK_TEXT, bold: true
        });

        const nullCards = [
            { varName: "부제목 (Subtitle)", count: "288건 (28.8%)", reason: "단권 도서의 부제목 미입력 기획으로 자연스럽게 발생하는 결측" },
            { varName: "할인율 (Discount Rate)", count: "111건 (11.1%)", reason: "정가 그대로 판매되는 0% 할인 도서이거나 무할인 정책 반영" },
            { varName: "평점 & 리뷰수 (Rating/Review)", count: "199건 (19.9%)", reason: "출간 직후 아직 독자 평가가 누적되지 않은 초기 신간 베스트셀러" }
        ];

        nullCards.forEach((nc, idx) => {
            const nY = 2.4 + idx * 1.4;
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: 7.0, y: nY, w: 5.233, h: 1.25, rectRadius: 0.05,
                fill: { color: C_BRICK_BG }, line: { color: C_BRICK, width: 0.8 }
            });
            slide.addText(`${nc.varName} : ${nc.count}`, {
                x: 7.2, y: nY + 0.15, w: 4.8, h: 0.3,
                fontFace: FONT_TITLE, fontSize: 13, color: C_BRICK_DARK, bold: true
            });
            slide.addText(nc.reason, {
                x: 7.2, y: nY + 0.45, w: 4.8, h: 0.7,
                fontFace: FONT_BODY, fontSize: 11, color: C_MUTED_TEXT, lineSpacing: 16
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
1.2 절에서는 변수 구조 및 수집 과정에서 관찰된 결측치(Missing Values) 현황을 말씀드립니다.

부제목(28.8%), 할인율(11.1%), 평점 및 리뷰수(19.9%) 3개 주요 변수에서 결측치가 관찰되었습니다. 평점/리뷰수의 199건 결측은 데이터 오류가 아닌, 출간된 지 얼마 되지 않아 독자 서평이 누적되지 않은 신규 진입 신간들이므로 정밀 처리하여 데이터 손실을 최소화했습니다.`);
    }

    // =========================================================================
    // SLIDE 5: [1. 개요] Top 5 vs Bottom 5 샘플 프리뷰
    // =========================================================================
    {
        const slide = createNordicSlide("1.3 베스트셀러 최상위 Top 5 vs 하위 Bottom 5 비교");

        slide.addText("TOP 5 메가 베스트셀러 (순위 1위 ~ 5위)", {
            x: 0.8, y: 1.5, w: 11.733, h: 0.35,
            fontFace: FONT_TITLE, fontSize: 14, color: C_BRICK, bold: true
        });

        const topHeaders = ["순위", "도서명", "저자", "출판사", "정가", "판매가", "판매지수", "평점"];
        const topRows = [
            ["1위", "바로바로 클로드 with 코워크, 스킬...", "차진우저", "골든래빗", "28,000원", "25,200원", "45,369", "9.7"],
            ["2위", "뚝딱 바로 써먹는 AI 3대장 챗GPT...", "코리아교육그룹", "안경다리BOOKS", "22,000원", "19,800원", "24,654", "9.9"],
            ["3위", "혼자 공부하는 바이브 코딩 with...", "조태호저", "한빛미디어", "30,000원", "27,000원", "77,061", "9.9"],
            ["4위", "클로드 에이전트 협업의 기술", "조쉬 (김승권)저", "한빛미디어", "30,000원", "27,000원", "5,376", "9.0"],
            ["5위", "클로드 코드 제대로 시작하기", "주홍철,황진성저", "길벗", "32,000원", "28,800원", "3,960", "10.0"]
        ];

        slide.addTable([topHeaders, ...topRows], {
            x: 0.8, y: 1.9, w: 11.733, colW: [0.8, 3.8, 1.6, 1.6, 1.1, 1.1, 1.1, 0.6],
            fontFace: FONT_BODY, fontSize: 10, color: C_DARK_TEXT,
            headerFill: C_BRICK, headerColor: "FFFFFF",
            border: { pt: 0.5, color: C_BORDER }, align: "center", fill: C_CARD_BG
        });

        slide.addText("BOTTOM 5 턱걸이 베스트셀러 (순위 996위 ~ 1000위)", {
            x: 0.8, y: 4.1, w: 11.733, h: 0.35,
            fontFace: FONT_TITLE, fontSize: 14, color: C_MUTED_TEXT, bold: true
        });

        const botRows = [
            ["996위", "Node.js 백엔드 개발자 되기", "박승규저", "골든래빗", "38,000원", "34,200원", "708", "9.9"],
            ["997위", "챗GPT를 활용한 40가지 파이썬...", "장문철저", "앤써북", "17,700원", "15,930원", "354", "9.3"],
            ["998위", "마인크래프트 놀라운 발명품 만들기", "Mojang Studio", "영진닷컴", "16,000원", "14,400원", "246", "6.0"],
            ["999위", "SPOTFIRE 응용편 1", "김성기저", "한나래아카데미", "19,000원", "19,000원", "120", "-"],
            ["1000위", "개념탑재 퓨전 360 디자인 모델링", "이예진저", "피앤피북", "22,000원", "19,800원", "450", "7.5"]
        ];

        slide.addTable([topHeaders, ...botRows], {
            x: 0.8, y: 4.5, w: 11.733, colW: [0.8, 3.8, 1.6, 1.6, 1.1, 1.1, 1.1, 0.6],
            fontFace: FONT_BODY, fontSize: 10, color: C_DARK_TEXT,
            headerFill: C_MUTED_TEXT, headerColor: "FFFFFF",
            border: { pt: 0.5, color: C_BORDER }, align: "center", fill: C_CARD_BG
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
1.3 절에서는 상위 1위~5위 도서와 하위 996위~1000위 도서를 비교한 샘플 프리뷰입니다.

Top 5 도서명에는 '클로드', '챗GPT', 'AI 3대장', '바이브 코딩' 등 인공지능 관련 키워드가 일률적으로 진입해 있으며, 판매지수 역시 77,061에 달합니다. 하위 5개 도서는 판매지수가 120~700 수준으로 상하위 도서 간 판매지수 격차가 최대 600배 이상 벌어지는 흥행 비대칭성을 볼 수 있습니다.`);
    }

    // =========================================================================
    // SLIDE 6: [2. 기술통계] 주요 수치형 변수 기술통계 요약
    // =========================================================================
    {
        const slide = createNordicSlide("2.1 수치형 변수 종합 기술통계 (Descriptive Stats)");

        const statsSummary = [
            { icon: "price", label: "정가 평균 (Mean Price)", val: "26,128원", sub: "중앙값 25,000원" },
            { icon: "discount", label: "판매가 평균 (Sales Price)", val: "23,845원", sub: "중앙값 22,500원" },
            { icon: "checklist", label: "평균 할인율 (Discount)", val: "9.86%", sub: "75% 사분위 10.0%" },
            { icon: "chart", label: "평균 판매지수 (Sales Index)", val: "2,768 pt", sub: "최대 77,061 pt" }
        ];

        statsSummary.forEach((st, idx) => {
            const xPos = 0.8 + idx * 2.95;
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: xPos, y: 1.5, w: 2.8, h: 1.35, rectRadius: 0.08,
                fill: { color: C_CARD_BG }, line: { color: C_BRICK, width: 1 }
            });

            if (flatIconBase64Map[st.icon]) {
                slide.addImage({ data: flatIconBase64Map[st.icon], x: xPos + 0.15, y: 1.65, w: 0.35, h: 0.35 });
            }

            slide.addText(st.label, {
                x: xPos + 0.55, y: 1.6, w: 2.15, h: 0.25,
                fontFace: FONT_BODY, fontSize: 10.5, color: C_MUTED_TEXT
            });
            slide.addText(st.val, {
                x: xPos + 0.1, y: 1.95, w: 2.6, h: 0.45,
                fontFace: FONT_TITLE, fontSize: 17, color: C_BRICK_DARK, bold: true, align: "center"
            });
            slide.addText(st.sub, {
                x: xPos + 0.1, y: 2.4, w: 2.6, h: 0.35,
                fontFace: FONT_BODY, fontSize: 9.5, color: C_DARK_TEXT, align: "center"
            });
        });

        const statsHeaders = ["변수명", "데이터수", "평균 (Mean)", "표준편차 (Std)", "최소 (Min)", "25% (Q1)", "중앙값 (Q2)", "75% (Q3)", "최대 (Max)"];
        const statsTableData = [
            ["정가_num", "1,000", "26,128원", "9,550.5", "10,000원", "20,000원", "25,000원", "30,000원", "127,100원"],
            ["판매가_num", "1,000", "23,845원", "8,923.6", "9,000원", "18,000원", "22,500원", "28,000원", "127,100원"],
            ["할인율 (%)", "889", "9.86%", "0.84%", "3.0%", "10.0%", "10.0%", "10.0%", "10.0%"],
            ["할인금액", "1,000", "2,282원", "1,156.2", "0원", "1,800원", "2,300원", "2,825원", "8,000원"],
            ["판매지수_num", "1,000", "2,768.4", "6,878.7", "60", "450", "1,144.5", "2,446.5", "77,061"],
            ["평점_num", "801", "9.65점", "0.70", "2.0점", "9.6점", "9.9점", "10.0점", "10.0점"],
            ["리뷰수_num", "801", "26.94건", "77.98", "1건", "6건", "14건", "28건", "1,991건"]
        ];

        slide.addTable([statsHeaders, ...statsTableData], {
            x: 0.8, y: 3.15, w: 11.733, colW: [1.333, 1.1, 1.4, 1.3, 1.2, 1.3, 1.3, 1.4, 1.4],
            fontFace: FONT_BODY, fontSize: 10, color: C_DARK_TEXT,
            headerFill: C_DARK_TEXT, headerColor: "FFFFFF",
            border: { pt: 0.5, color: C_BORDER }, align: "center", fill: C_CARD_BG
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
2.1 절에서는 7개 수치형 변수의 기술통계 요약입니다. 정가 평균 26,128원, 판매가 평균 23,845원, 중앙값은 각각 25,000원과 22,500원입니다.

판매지수의 평균은 2,768인 반면 중간값은 1,144.5에 불과하여 2.4배 이상의 격차를 보이는 우측 꼬리 롱테일 분포를 정밀하게 입증해 줍니다.`);
    }

    // =========================================================================
    // SLIDE 7: [2. 기술통계] 핵심 범주형 변수 요약
    // =========================================================================
    {
        const slide = createNordicSlide("2.2 범주형 변수 요약 & 생태계 집중도");

        const catCards = [
            { icon: "publisher", title: "출판사 (Publisher)", count: "190 개 고유사", top: "한빛미디어", topCount: "149 권 (14.9%)", desc: "상위 10개 대형 전문 출판사가 전체의 40% 이상을 과점 점유" },
            { icon: "author", title: "대표저자 (Author)", count: "814 명 고유저자", top: "Mojang Studio", topCount: "16 권 (1.6%)", desc: "마인크래프트 및 IT 전문 저자군의 연쇄 흥행 팬덤 구축" },
            { icon: "trend", title: "출간연월 (Pub Date)", count: "92 개 월별구분", top: "2026년 07월", topCount: "87 권 (8.7%)", desc: "출간 직후 1~3개월 이내 신간 집중 효과가 베스트셀러 진입 결정" }
        ];

        catCards.forEach((card, idx) => {
            const xPos = 0.8 + idx * 3.95;
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: xPos, y: 1.6, w: 3.8, h: 4.9, rectRadius: 0.08,
                fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
            });

            if (flatIconBase64Map[card.icon]) {
                slide.addImage({ data: flatIconBase64Map[card.icon], x: xPos + 0.3, y: 1.85, w: 0.5, h: 0.5 });
            }

            slide.addText(card.title, {
                x: xPos + 0.9, y: 1.9, w: 2.7, h: 0.35,
                fontFace: FONT_TITLE, fontSize: 15, color: C_DARK_TEXT, bold: true
            });

            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: xPos + 0.2, y: 2.45, w: 3.4, h: 0.55, rectRadius: 0.05,
                fill: { color: C_BRICK_BG }, line: { color: C_BRICK, width: 0.8 }
            });
            slide.addText("고유값(Unique)", {
                x: xPos + 0.3, y: 2.5, w: 1.5, h: 0.45,
                fontFace: FONT_BODY, fontSize: 10, color: C_MUTED_TEXT
            });
            slide.addText(card.count, {
                x: xPos + 1.8, y: 2.5, w: 1.7, h: 0.45,
                fontFace: FONT_TITLE, fontSize: 12, color: C_BRICK_DARK, bold: true, align: "right"
            });

            slide.addText("최빈값 (Top Mode)", {
                x: xPos + 0.2, y: 3.2, w: 3.4, h: 0.25,
                fontFace: FONT_BODY, fontSize: 11, color: C_MUTED_TEXT
            });
            slide.addText(card.top, {
                x: xPos + 0.2, y: 3.45, w: 3.4, h: 0.4,
                fontFace: FONT_TITLE, fontSize: 15, color: C_DARK_TEXT, bold: true
            });
            slide.addText(`빈도수: ${card.topCount}`, {
                x: xPos + 0.2, y: 3.85, w: 3.4, h: 0.3,
                fontFace: FONT_BODY, fontSize: 12, color: C_BRICK, bold: true
            });

            slide.addText(card.desc, {
                x: xPos + 0.2, y: 4.3, w: 3.4, h: 1.9,
                fontFace: FONT_BODY, fontSize: 12, color: C_MUTED_TEXT, lineSpacing: 18
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
2.2 절에서는 범주형 변수의 생태계 집중도입니다. 한빛미디어가 149권(14.9%)으로 독보적 1위이며, 저자는 Mojang Studio(16권) 등 기술서 전문 저자의 팬덤 연쇄 흥행 현상이 관찰됩니다.`);
    }

    // =========================================================================
    // SLIDE 8: [2. 심층 분석 1] 도서 가격 구조 & 10% 정률 할인 전략
    // =========================================================================
    {
        const slide = createNordicSlide("2.3 심층 분석 ① : 도서 가격 구조 & 10% 할인 메커니즘");

        const pricePipeline = [
            { icon: "price", step: "STEP 1. 정가 책정", val: "26,128원", detail: "중앙값 25,000원\n사분위 2만~3만원" },
            { icon: "discount", step: "STEP 2. 도서정가제", val: "- 10.0% 정률", detail: "법적 최대 할인 한도\n88.9% 도서 일괄 적용" },
            { icon: "price", step: "STEP 3. 평균 할인액", val: "- 2,283원", detail: "체감 할인 혜택액\n표준편차 1,156원" },
            { icon: "book", step: "STEP 4. 최종 판매가", val: "23,845원", detail: "중앙값 22,500원\n실제 결제 금액" }
        ];

        pricePipeline.forEach((p, idx) => {
            const xPos = 0.8 + idx * 2.95;
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: xPos, y: 1.6, w: 2.8, h: 2.2, rectRadius: 0.08,
                fill: { color: idx === 1 ? C_BRICK_BG : C_CARD_BG },
                line: { color: idx === 1 ? C_BRICK : C_BORDER, width: 1.5 }
            });

            if (flatIconBase64Map[p.icon]) {
                slide.addImage({ data: flatIconBase64Map[p.icon], x: xPos + 1.2, y: 1.75, w: 0.4, h: 0.4 });
            }

            slide.addText(p.step, {
                x: xPos + 0.1, y: 2.2, w: 2.6, h: 0.25,
                fontFace: FONT_BODY, fontSize: 10.5, color: C_MUTED_TEXT, align: "center"
            });
            slide.addText(p.val, {
                x: xPos + 0.1, y: 2.45, w: 2.6, h: 0.4,
                fontFace: FONT_TITLE, fontSize: 17, color: idx === 1 ? C_BRICK_DARK : C_DARK_TEXT, bold: true, align: "center"
            });
            slide.addText(p.detail, {
                x: xPos + 0.1, y: 2.9, w: 2.6, h: 0.7,
                fontFace: FONT_BODY, fontSize: 11, color: C_DARK_TEXT, align: "center", lineSpacing: 15
            });
        });

        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 0.8, y: 4.1, w: 11.733, h: 2.4, rectRadius: 0.08,
            fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
        });

        slide.addText("도서 가격 및 할인 메커니즘 3대 시사점", {
            x: 1.1, y: 4.3, w: 11.0, h: 0.35,
            fontFace: FONT_TITLE, fontSize: 15, color: C_DARK_TEXT, bold: true
        });

        const priceInsights = [
            "① 가격 밀집도: 베스트셀러 도서의 50% 이상이 정가 20,000원~30,000원 범위 내에 집약적으로 분포하는 정규 구조 형성",
            "② 할인 고착화: 한국 도서정가제 규제(최대 10% 가격 할인)에 따라, 전체 도서의 88.9%가 정확히 10% 정률 할인을 적용함",
            "③ 마케팅 경쟁 축 이동: 할인율 자체로는 가격 경쟁력을 확보할 수 없으므로, 사은품, 마일리지 적립, 독점 리워드가 본질적 마케팅 차별화 요소임"
        ];

        priceInsights.forEach((pi, idx) => {
            slide.addText(pi, {
                x: 1.1, y: 4.75 + idx * 0.5, w: 11.1, h: 0.45,
                fontFace: FONT_BODY, fontSize: 12, color: C_DARK_TEXT
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
2.3 절에서는 가격 구조와 10% 할인 메커니즘입니다. 베스트셀러 정가 평균 2.6만 원에 10% 할인이 일률 적용되어 최종 2.3만 원에 구매됩니다. 할인율이 10%로 고착화되었으므로 사은품, 포인트 등 비가격 마케팅이 핵심 승부처입니다.`);
    }

    // =========================================================================
    // SLIDE 9: [2. 심층 분석 2] 판매지수 흥행 비대칭성 & 롱테일 파레토
    // =========================================================================
    {
        const slide = createNordicSlide("2.4 심층 분석 ② : 판매지수 흥행 비대칭성 & 롱테일 법칙");

        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 0.8, y: 1.6, w: 5.6, h: 4.9, rectRadius: 0.08,
            fill: { color: C_BRICK_BG }, line: { color: C_BRICK, width: 1.5 }
        });

        if (flatIconBase64Map["star"]) {
            slide.addImage({ data: flatIconBase64Map["star"], x: 1.1, y: 1.85, w: 0.45, h: 0.45 });
        }

        slide.addText("상위 5% 메가 베스트셀러 (Top 5%)", {
            x: 1.65, y: 1.9, w: 4.5, h: 0.4,
            fontFace: FONT_TITLE, fontSize: 17, color: C_BRICK_DARK, bold: true
        });
        slide.addText("판매지수: 10,000 ~ 77,061 pt", {
            x: 1.1, y: 2.35, w: 5.0, h: 0.35,
            fontFace: FONT_TITLE, fontSize: 15, color: C_DARK_TEXT, bold: true
        });
        const topPoints = [
            "• 전체 베스트셀러 총 판매량의 절대 다수를 압도적으로 견인",
            "• 주요 서적: '혼자 공부하는 바이브 코딩', '바로바로 클로드' 등 AI 킬러 콘텐츠",
            "• 출판사 재고 관리 및 물류 리드타임 집중 수립 필요"
        ];
        topPoints.forEach((tp, i) => {
            slide.addText(tp, {
                x: 1.1, y: 2.8 + i * 0.7, w: 5.0, h: 0.6,
                fontFace: FONT_BODY, fontSize: 12, color: C_DARK_TEXT, lineSpacing: 16
            });
        });

        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 6.7, y: 1.6, w: 5.833, h: 4.9, rectRadius: 0.08,
            fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
        });

        slide.addText("하위 95% 일반 베스트셀러 (Bottom 95%)", {
            x: 7.0, y: 1.9, w: 5.2, h: 0.4,
            fontFace: FONT_TITLE, fontSize: 17, color: C_DARK_TEXT, bold: true
        });
        slide.addText("판매지수: 60 ~ 2,446 pt (중앙값 1,144 pt)", {
            x: 7.0, y: 2.35, w: 5.2, h: 0.35,
            fontFace: FONT_TITLE, fontSize: 15, color: C_MUTED_TEXT, bold: true
        });
        const botPoints = [
            "• 베스트셀러 목록에 등재되었으나 순위권 턱걸이 수준 유지",
            "• 75% 사분위수 도서조차 판매지수 2,446 수준으로 메가작과 30배 차이",
            "• 롱테일(Long-tail) 우측 왜도(Skewness > 3.0) 구조의 명확한 증명"
        ];
        botPoints.forEach((bp, i) => {
            slide.addText(bp, {
                x: 7.0, y: 2.8 + i * 0.7, w: 5.2, h: 0.6,
                fontFace: FONT_BODY, fontSize: 12, color: C_DARK_TEXT, lineSpacing: 16
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
2.4 절에서는 상위 5% 메가 베스트셀러와 하위 95% 일반 베스트셀러 간의 파레토 롱테일 양극화 현상입니다. 메가 베스트셀러의 재고 공급망 완벽 유지가 필수적입니다.`);
    }

    // =========================================================================
    // SLIDE 10: [2. 심층 분석 3] 독자 평점 및 리뷰 반응성 분석
    // =========================================================================
    {
        const slide = createNordicSlide("2.5 심층 분석 ③ : 독자 평점 & 리뷰 반응성 동역학");

        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 0.8, y: 1.6, w: 5.6, h: 4.9, rectRadius: 0.08,
            fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
        });

        if (flatIconBase64Map["star"]) {
            slide.addImage({ data: flatIconBase64Map["star"], x: 1.1, y: 1.85, w: 0.45, h: 0.45 });
        }

        slide.addText("독자 평점 (Rating) : 기본 조건", {
            x: 1.65, y: 1.9, w: 4.5, h: 0.4,
            fontFace: FONT_TITLE, fontSize: 15, color: C_DARK_TEXT, bold: true
        });
        slide.addText("평균 9.65점 | 중앙값 9.9점", {
            x: 1.1, y: 2.3, w: 5.0, h: 0.4,
            fontFace: FONT_TITLE, fontSize: 19, color: C_BRICK, bold: true
        });
        const ratingDetails = [
            "• 25% 사분위수가 9.6점에 달하여 대부분 도서가 극히 높은 평점 확보",
            "• 표준편차가 0.70으로 변동성이 매우 적음 (평준화 현상)",
            "• 결론: 9.5점 이상의 높은 평점은 베스트셀러 진입의 최소 전제 조건(Baseline)임"
        ];
        ratingDetails.forEach((rd, i) => {
            slide.addText(rd, {
                x: 1.1, y: 2.85 + i * 0.7, w: 5.0, h: 0.6,
                fontFace: FONT_BODY, fontSize: 12, color: C_DARK_TEXT, lineSpacing: 16
            });
        });

        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 6.7, y: 1.6, w: 5.833, h: 4.9, rectRadius: 0.08,
            fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
        });

        if (flatIconBase64Map["review"]) {
            slide.addImage({ data: flatIconBase64Map["review"], x: 7.0, y: 1.85, w: 0.45, h: 0.45 });
        }

        slide.addText("리뷰 수 (Review Count) : 흥행 지속 파워", {
            x: 7.55, y: 1.9, w: 4.8, h: 0.4,
            fontFace: FONT_TITLE, fontSize: 15, color: C_DARK_TEXT, bold: true
        });
        slide.addText("평균 26.9건 | 최댓값 1,991건", {
            x: 7.0, y: 2.3, w: 5.2, h: 0.4,
            fontFace: FONT_TITLE, fontSize: 19, color: C_BRICK_DARK, bold: true
        });
        const reviewDetails = [
            "• 중앙값은 14건 수준이나, 스테디셀러의 경우 1,000건 이상 폭발적 축적",
            "• 독자 리뷰의 절대 볼륨(Absolute Volume)이 구전 효과(WOM) 창출",
            "• 결론: 리뷰 볼륨 생성이 베스트셀러의 수명(Lifetime) 확장 핵심 동력임"
        ];
        reviewDetails.forEach((rv, i) => {
            slide.addText(rv, {
                x: 7.0, y: 2.85 + i * 0.7, w: 5.2, h: 0.6,
                fontFace: FONT_BODY, fontSize: 12, color: C_DARK_TEXT, lineSpacing: 16
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
2.5 절에서는 평점 9.5점 이상 확보는 기본 조건(Baseline)이며, 신속한 30건 이상의 리뷰 볼륨 생성이 흥행 수명을 확장한다는 분석입니다.`);
    }

    // =========================================================================
    // SLIDE 11: [2. 심층 분석 4] 출판사 & 저자 파워 생태계 집중도
    // =========================================================================
    {
        const slide = createNordicSlide("2.6 심층 분석 ④ : 출판사 & 저자 파워 지형도");

        const ecoCards = [
            { icon: "publisher", title: "출판사 과점 생태계", main: "Top 10 출판사 점유율 40%+", sub: "한빛미디어(149권), 길벗(79권), 이지스(55권), 골든래빗(50권) 등 대형 IT/전문서 브랜드가 기획력과 유통 독점" },
            { icon: "author", title: "저자 브랜딩 파워", main: "검증된 베스트셀러 저자의 연쇄 흥행", sub: "Mojang Studio(16권), 조태호, 차진우 등 AI/코딩 저자의 팬덤 및 신뢰도가 차기작 판매량 즉각 견인" },
            { icon: "trend", title: "신간 집중 마케팅 골든타임", main: "출간 후 1~3개월 신간 효과 피크", sub: "2026년 7월 출간작 87권 진입. 출간 초기 1개월 이내 마케팅 자원 집중 투입이 베스트셀러 안착의 승패 결정" }
        ];

        ecoCards.forEach((ec, idx) => {
            const xPos = 0.8 + idx * 3.95;
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: xPos, y: 1.6, w: 3.8, h: 4.9, rectRadius: 0.08,
                fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
            });

            if (flatIconBase64Map[ec.icon]) {
                slide.addImage({ data: flatIconBase64Map[ec.icon], x: xPos + 0.3, y: 1.85, w: 0.45, h: 0.45 });
            }

            slide.addText(ec.title, {
                x: xPos + 0.85, y: 1.9, w: 2.7, h: 0.4,
                fontFace: FONT_TITLE, fontSize: 15, color: C_DARK_TEXT, bold: true
            });

            slide.addText(ec.main, {
                x: xPos + 0.3, y: 2.5, w: 3.2, h: 0.6,
                fontFace: FONT_TITLE, fontSize: 13, color: C_BRICK, bold: true, lineSpacing: 16
            });

            slide.addText(ec.sub, {
                x: xPos + 0.3, y: 3.2, w: 3.2, h: 3.0,
                fontFace: FONT_BODY, fontSize: 12, color: C_MUTED_TEXT, lineSpacing: 18
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
2.6 절에서는 대형 출판사의 과점화, 저자 브랜딩 파워, 신간 1~3개월 초기 마케팅 골든타임이라는 3대 생태계 구조를 정리합니다.`);
    }

    // =========================================================================
    // SLIDE 12 ~ 27: 11종 특화 시각화 차트 슬라이드 (Nordic Style)
    // =========================================================================
    
    // SLIDE 12: Chart 1
    {
        const slide = createNordicSlide("3.1 [Chart 1] YES24 베스트셀러 판매지수 분포 (Univariate)");
        const imgPath = path.join(IMG_DIR, "01_sales_index_dist.png");
        if (fs.existsSync(imgPath)) slide.addImage({ path: imgPath, x: 0.8, y: 1.6, w: 6.2, h: 4.9 });

        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 7.2, y: 1.6, w: 5.333, h: 4.9, rectRadius: 0.08,
            fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
        });

        slide.addText("연동 통계 요약표", {
            x: 7.4, y: 1.8, w: 4.9, h: 0.3,
            fontFace: FONT_TITLE, fontSize: 14, color: C_DARK_TEXT, bold: true
        });

        const c1Headers = ["지표", "Count", "Mean", "Std", "Min", "50%", "Max"];
        const c1Data = [["판매지수", "1,000", "2,768.4", "6,878.7", "60", "1,144.5", "77,061"]];
        slide.addTable([c1Headers, ...c1Data], {
            x: 7.4, y: 2.15, w: 4.933, colW: [0.9, 0.6, 0.8, 0.8, 0.5, 0.7, 0.633],
            fontFace: FONT_BODY, fontSize: 8.5, color: C_DARK_TEXT,
            headerFill: C_BRICK, headerColor: "FFFFFF", align: "center", fill: C_CARD_BG
        });

        slide.addText("비즈니스 인사이트 및 심층 해석", {
            x: 7.4, y: 2.9, w: 4.9, h: 0.3,
            fontFace: FONT_TITLE, fontSize: 14, color: C_BRICK, bold: true
        });

        const c1InsightText = "판매지수 분포는 극심한 우측 꼬리 분포(Right-skewed)를 보입니다. 대부분의 베스트셀러 도서는 판매지수 5,000 미만 구간에 조밀하게 밀집해 있는 반면, 상위 1%에 해당하는 초대형 흥행작들이 판매지수 50,000~77,000 구간에 위치하고 있습니다.\n\n이는 베스트셀러 순위에 등재된 책이라 할지라도 상위권 메가 베스트셀러와 하위권 베스트셀러 간의 실질 판매 격차가 수십 배에 달함을 의미합니다. 상위 5% 메가 베스트셀러의 재고 및 리드타임 집중 관리 전략이 필수적입니다.";
        slide.addText(c1InsightText, {
            x: 7.4, y: 3.25, w: 4.933, h: 3.1,
            fontFace: FONT_BODY, fontSize: 11, color: C_DARK_TEXT, lineSpacing: 16
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
3.1 절에서는 판매지수의 우측 꼬리 롱테일 분포 차트를 통한 상위 5% 메가 베스트셀러 수량 집중 관리 필요성을 다룹니다.`);
    }

    // SLIDE 13: Chart 1 S-A-B
    {
        const slide = createNordicSlide("3.1 [Chart 1] 판매지수 구간별 대응 프로세스");

        const steps = [
            { level: "구간 1: S-Grade (메가작)", range: "판매지수 10,000 pt 이상", action: "안전재고 상시 확보 & 물류 리드타임 제로화. 마케팅 예산 추가 투입으로 스테디셀러화 추진" },
            { level: "구간 2: A-Grade (주요작)", range: "판매지수 2,500 ~ 9,999 pt", action: "리뷰 수 30건 조기 확보 및 대표저자 팬덤 서평단 운영. 10% 할인 외 독점 굿즈 제공" },
            { level: "구간 3: B-Grade (일반작)", range: "판매지수 2,500 pt 미만", action: "체과 재고 방지를 위한 주문형 생산(POD) 및 온라인 기획전 연계로 순위 방어" }
        ];

        steps.forEach((st, idx) => {
            const yPos = 1.6 + idx * 1.7;
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: 0.8, y: yPos, w: 11.733, h: 1.45, rectRadius: 0.08,
                fill: { color: idx === 0 ? C_BRICK_BG : C_CARD_BG },
                line: { color: idx === 0 ? C_BRICK : C_BORDER, width: 1.5 }
            });

            slide.addText(st.level, {
                x: 1.1, y: yPos + 0.2, w: 4.5, h: 0.35,
                fontFace: FONT_TITLE, fontSize: 16, color: idx === 0 ? C_BRICK_DARK : C_DARK_TEXT, bold: true
            });
            slide.addText(st.range, {
                x: 1.1, y: yPos + 0.6, w: 4.5, h: 0.3,
                fontFace: FONT_BODY, fontSize: 12, color: C_MUTED_TEXT
            });
            slide.addText(st.action, {
                x: 5.5, y: yPos + 0.2, w: 6.7, h: 1.0,
                fontFace: FONT_BODY, fontSize: 12, color: C_DARK_TEXT, lineSpacing: 18
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
3.1 절 연계로서 판매지수 10,000pt 이상 S등급, A등급, B등급별 3단계 차별화 마케팅/재고 대응 가이드입니다.`);
    }

    // SLIDE 14: Chart 2 Boxplot
    {
        const slide = createNordicSlide("3.2 [Chart 2] 도서 정가 및 판매가 분포 비교 (Boxplot)");
        const imgPath = path.join(IMG_DIR, "02_price_distribution.png");
        if (fs.existsSync(imgPath)) slide.addImage({ path: imgPath, x: 0.8, y: 1.6, w: 6.2, h: 4.9 });

        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 7.2, y: 1.6, w: 5.333, h: 4.9, rectRadius: 0.08,
            fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
        });

        slide.addText("연동 통계 요약표", {
            x: 7.4, y: 1.8, w: 4.9, h: 0.3,
            fontFace: FONT_TITLE, fontSize: 14, color: C_DARK_TEXT, bold: true
        });

        const c2Headers = ["변수명", "Mean", "Std", "Min", "50%", "Max"];
        const c2Data = [
            ["정가_num", "26,128", "9,550.5", "10,000", "25,000", "127,100"],
            ["판매가_num", "23,845", "8,923.6", "9,000", "22,500", "127,100"],
            ["할인금액", "2,282.8", "1,156.2", "0", "2,300", "8,000"]
        ];
        slide.addTable([c2Headers, ...c2Data], {
            x: 7.4, y: 2.15, w: 4.933, colW: [1.2, 0.8, 0.8, 0.7, 0.7, 0.733],
            fontFace: FONT_BODY, fontSize: 8.5, color: C_DARK_TEXT,
            headerFill: C_BRICK, headerColor: "FFFFFF", align: "center", fill: C_CARD_BG
        });

        slide.addText("비즈니스 인사이트 및 심층 해석", {
            x: 7.4, y: 3.3, w: 4.9, h: 0.3,
            fontFace: FONT_TITLE, fontSize: 14, color: C_BRICK, bold: true
        });

        const c2InsightText = "도서 정가와 판매가의 박스플롯 비교 결과, 두 가격 변수의 사분위 범위(IQR)가 완벽히 평행하게 형성되어 있습니다. 정가의 상위 25%~75% 구간은 20,000원에서 30,000원 사이이며, 판매가는 18,000원에서 27,000원 사이입니다. 정가 대비 판매가의 일정한 차이는 10% 가격 할인이 베스트셀러 전반에 거쳐 엄격하게 적용되고 있음을 증명합니다.";
        slide.addText(c2InsightText, {
            x: 7.4, y: 3.65, w: 4.933, h: 2.7,
            fontFace: FONT_BODY, fontSize: 11, color: C_DARK_TEXT, lineSpacing: 16
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
3.2 절 박스플롯 분석에서는 정가와 판매가 사분위 범위의 완벽한 평행 이동을 통해 10% 정률 할인의 준수 실태를 증명합니다.`);
    }

    // SLIDE 15 ~ 27 추가 특화 차트 슬라이드들도 완벽한 노르딕 스타일과 벽돌색으로 순차 생성
    const chartSlidesData = [
        {
            num: "3.2", chartNo: "2", title: "가격대별 분포 & 고가 서적 이상치 해석",
            card1Title: "표준 베스트셀러 가격대", card1Val: "정가 20,000원 ~ 30,000원 구간", card1Points: ["• 체감 가격 저항선 25,000원 안팎 형성", "• IT 실용서 및 AI 활용서 대다수 포함", "• 정가 2.5만 원 기준 10% 할인 유도"],
            card2Title: "고가 서적 이상치 (Outlier)", card2Val: "정가 50,000원 ~ 127,100원", card2Points: ["• 전문 수험서, 학술서, 전집류 포함", "• 구매층 한정으로 최고 판매지수 한계", "• 권당 단가가 높아 매출액(Revenue) 기여"],
            notes: "3.2 절 연계로서 표준 가격대와 10만 원 이상 고가 서적 이상치의 비즈니스 특성 비교입니다."
        },
        {
            num: "3.3", chartNo: "3", title: "[Chart 3] 도서 할인율 빈도 분포 (Univariate)",
            img: "03_discount_rate_dist.png",
            headers: ["할인율 (%)", "도서 수 (Count)", "비율 (%)"],
            tableData: [["10% 할인", "864 권", "86.4 %"], ["5% 할인", "24 권", "2.4 %"], ["3% 할인", "1 권", "0.1 %"], ["0% (무할인)", "111 권", "11.1 %"]],
            insight: "할인율 빈도 분석 결과 88.9%에 달하는 도서가 정확히 10% 할인율을 적용받고 있습니다. 10% 할인은 베스트셀러 진입의 디폴트 수단으로 고착화되었으므로 비가격적 마케팅 혜택이 핵심 차별화 요소입니다.",
            notes: "3.3 절 할인율 빈도 차트로 10% 정률 할인율의 디폴트 고착화 현상을 입증합니다."
        },
        {
            num: "3.3", chartNo: "3", title: "[Chart 3] 도서정가제 하 3대 비가격 마케팅 축",
            axes: [
                { title: "1. 독점 굿즈 / 사은품 (Goods)", desc: "도서 컨셉 연계 독점 굿즈(노트, 데스크용품, AI 프롬프트 카드) 제공" },
                { title: "2. 마일리지 & 포인트 (Loyalty)", desc: "유통 플랫폼 마일리지 추가 적립(최대 5%)으로 결제 혜택 극대화" },
                { title: "3. 저자 팬덤 콘텐츠 (Exclusive)", desc: "저자 친필 사인본, VOD 강좌 쿠폰, 독점 부록 PDF 연계" }
            ],
            notes: "3.3 절 연계 비가격 마케팅 3대 차별화 축입니다."
        },
        {
            num: "3.4", chartNo: "4", title: "[Chart 4] 베스트셀러 출판사 상위 30개 (Categorical)",
            img: "04_top30_publishers.png",
            headers: ["순위", "출판사명", "베스트셀러 권수", "점유율 (%)"],
            tableData: [["1위", "한빛미디어", "149 권", "14.9 %"], ["2위", "길벗", "79 권", "7.9 %"], ["3위", "이지스퍼블리싱", "55 권", "5.5 %"], ["4위", "골든래빗", "50 권", "5.0 %"], ["5위", "제이펍", "48 권", "4.8 %"], ["6위", "영진닷컴", "38 권", "3.8 %"]],
            insight: "상위 10개 출판사가 베스트셀러의 52.9%를 과점하고 있습니다. IT 및 전문서적 브랜드들의 탄탄한 기획력이 시장을 선도하고 있습니다.",
            notes: "3.4 절 상위 출판사 지형도로 대형 출판사의 과점화 현상을 설명합니다."
        },
        {
            num: "3.4", chartNo: "4", title: "[Chart 4] 대형 전문 출판사 4대 핵심 성공 패턴",
            patterns: [
                { title: "1. 브랜드 시리즈 구축", desc: "'혼자 공부하는~', 'Do it!~' 등 시그니처 입문서 브랜드 시리즈 구축" },
                { title: "2. 트렌드 키워드 빠른 선점", desc: "생성형 AI, 클로드 코드 등 기술 변곡점 발생 시 수개월 내 출간" },
                { title: "3. 탄탄한 저자 네트워크", desc: "검증된 전문 저자진 및 커뮤니티 네트워킹을 통한 원고 수급" },
                { title: "4. 초기 마케팅 자본 투입", desc: "출간 초기 서평단 및 서점 메인 배너 프로모션 집중 투입" }
            ],
            notes: "3.4 절 연계 대형 출판사의 4대 핵심 성공 패턴 분석입니다."
        },
        {
            num: "3.5", chartNo: "5", title: "[Chart 5] 베스트셀러 저자 상위 30명 (Categorical)",
            img: "05_top30_authors.png",
            headers: ["순위", "대표저자명", "진입 권수", "대표 분야"],
            tableData: [["1위", "Mojang Studio", "16 권", "마인크래프트 게임/코딩"], ["2위", "코리아교육그룹", "5 권", "IT 자격증/수험서"], ["3위", "박현규", "5 권", "컴퓨터 입문/실용서"], ["4위", "장문철", "5 권", "파이썬/챗GPT"], ["5위", "서지영", "4 권", "딥러닝/파이토치"]],
            insight: "1위 Mojang Studio(16권)를 비롯하여 검증된 베스트셀러 저자의 연쇄 흥행 파워 현상이 두드러집니다.",
            notes: "3.5 절 저자 상위 30명 분석으로 저자 브랜드 파워를 다룹니다."
        },
        {
            num: "3.6", chartNo: "6", title: "[Chart 6] 출판사별 평균 판매지수 상위 15개 (Bivariate)",
            img: "06_publisher_avg_sales_index.png",
            headers: ["출판사명", "출간권수", "평균 판매지수", "중앙값"],
            tableData: [["코리아닷컴", "1 권", "76,152 pt", "76,152"], ["사회평론아카데미", "1 권", "17,370 pt", "17,370"], ["메가스터디북스", "1 권", "14,322 pt", "14,322"], ["안경다리BOOKS", "2 권", "13,662 pt", "13,662"]],
            insight: "단순 양적 다작보다 시장 적합성(PMF)이 높은 킬러 콘텐츠 한 권 기획의 중요성을 입증합니다.",
            notes: "3.6 절 권당 평균 판매지수 상위 출판사 분석입니다."
        },
        {
            num: "3.7", chartNo: "7", title: "[Chart 7] 도서 정가 vs 판매지수 상관관계 (Bivariate)",
            img: "07_price_vs_sales_index.png",
            headers: ["가격대 구간", "도서 수", "평균 판매지수", "중앙값"],
            tableData: [["1.5만 원 이하", "87 권", "632.4 pt", "396 pt"], ["1.5만 ~ 2.5만", "464 권", "3,343.4 pt", "1,353 pt"], ["2.5만 ~ 3.5만", "317 권", "3,004.3 pt", "1,176 pt"], ["5.0만 원 이상", "16 권", "974.4 pt", "513 pt"]],
            insight: "판매지수 20,000 이상의 메가작은 정가 2.0만~3.0만 원 구간에 집약되어 체감 가격 저항선을 입증합니다.",
            notes: "3.7 절 산점도 분석으로 정가와 판매지수 간 골디락스 구간을 제시합니다."
        },
        {
            num: "3.8", chartNo: "8", title: "[Chart 8] 도서 평점 vs 리뷰수 분포 (Bivariate Scatter)",
            img: "08_rating_vs_reviews.png",
            headers: ["평점 구간", "평균 리뷰수", "평균 판매지수"],
            tableData: [["8.9점 이하", "12.83 건", "2,284.9 pt"], ["9.0 ~ 9.4점", "23.17 건", "2,307.6 pt"], ["9.5 ~ 9.8점", "42.73 건", "5,132.7 pt"], ["9.9 ~ 10.0점", "20.00 건", "2,485.0 pt"]],
            insight: "9.5~9.8점 평점 구간 도서들의 평균 리뷰수가 42.7건으로 가장 높아 자발적 구전 효과(WOM) 선순환을 입증합니다.",
            notes: "3.8 절 평점 대 리뷰수 분석입니다."
        },
        {
            num: "3.9", chartNo: "9", title: "[Chart 9] 출간연월별 베스트셀러 추이 (Time Series)",
            img: "09_pub_date_trend.png",
            headers: ["출간연월", "등록 권수", "비중 (%)"],
            tableData: [["2026년 07월", "87 권", "8.7 % (피크)"], ["2026년 06월", "60 권", "6.0 %"], ["2026년 05월", "51 권", "5.1 %"], ["2026년 03월", "51 권", "5.1 %"]],
            insight: "2026년 5~7월 신간이 목록의 대부분을 차지하여 신간 1~3개월 집중 마케팅의 중요성을 입증합니다.",
            notes: "3.9 절 출간연월 시계열 분석입니다."
        },
        {
            num: "3.10", chartNo: "10", title: "[Chart 10] 수치형 변수 상관관계 히트맵 (Correlation)",
            img: "10_correlation_heatmap.png",
            headers: ["변수쌍", "상관계수(r)", "비즈니스 의미"],
            tableData: [["정가 ↔ 판매가", "r = +0.99", "10% 정률 할인의 완벽한 선형성"], ["판매지수 ↔ 리뷰수", "r = +0.45", "리뷰 작성이 판매량 확대에 유의미 기여"], ["정가 ↔ 평점", "r = -0.07", "가격 수준과 독자 만족도는 무관"]],
            insight: "판매지수와 리뷰수 간 +0.45 상관관계는 리뷰 확보가 실질 판매량을 증대시키는 직접적 동력임을 입증합니다.",
            notes: "3.10 절 수치형 변수 상관관계 히트맵 분석입니다."
        },
        {
            num: "3.11", chartNo: "11", title: "[Chart 11] 도서명 TF-IDF 중요 키워드 Top 30",
            img: "11_tfidf_top30_keywords.png",
            headers: ["순위", "키워드", "TF-IDF 점수", "트렌드 분류"],
            tableData: [["1위", "AI", "109.19 pt", "생성형 AI (독보적 1위)"], ["2위", "위한", "58.06 pt", "수요 타깃 지향"], ["3위", "배우는", "39.44 pt", "입문/학습 서적"], ["4위", "가이드", "33.58 pt", "실무 가이드북"]],
            insight: "AI(109.19점) 키워드가 압도적 1위를 기록하며 인공지능 트렌드가 2026 도서 시장을 완벽 장악함을 입증합니다.",
            notes: "3.11 절 TF-IDF 키워드 상위 30개 분석입니다."
        }
    ];

    for (const data of chartSlidesData) {
        const slide = createNordicSlide(`${data.num} ${data.title}`);

        if (data.card1Title) {
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: 0.8, y: 1.6, w: 5.6, h: 4.9, rectRadius: 0.08,
                fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
            });
            slide.addText(data.card1Title, { x: 1.1, y: 1.9, w: 5.0, h: 0.4, fontFace: FONT_TITLE, fontSize: 16, color: C_DARK_TEXT, bold: true });
            slide.addText(data.card1Val, { x: 1.1, y: 2.35, w: 5.0, h: 0.4, fontFace: FONT_TITLE, fontSize: 17, color: C_BRICK, bold: true });
            data.card1Points.forEach((p, i) => {
                slide.addText(p, { x: 1.1, y: 2.9 + i * 0.7, w: 5.0, h: 0.6, fontFace: FONT_BODY, fontSize: 12, color: C_DARK_TEXT });
            });

            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: 6.7, y: 1.6, w: 5.833, h: 4.9, rectRadius: 0.08,
                fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
            });
            slide.addText(data.card2Title, { x: 7.0, y: 1.9, w: 5.2, h: 0.4, fontFace: FONT_TITLE, fontSize: 16, color: C_DARK_TEXT, bold: true });
            slide.addText(data.card2Val, { x: 7.0, y: 2.35, w: 5.2, h: 0.4, fontFace: FONT_TITLE, fontSize: 17, color: C_MUTED_TEXT, bold: true });
            data.card2Points.forEach((p, i) => {
                slide.addText(p, { x: 7.0, y: 2.9 + i * 0.7, w: 5.2, h: 0.6, fontFace: FONT_BODY, fontSize: 12, color: C_DARK_TEXT });
            });
        } else if (data.axes) {
            data.axes.forEach((ma, idx) => {
                const yPos = 1.6 + idx * 1.7;
                slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                    x: 0.8, y: yPos, w: 11.733, h: 1.45, rectRadius: 0.08,
                    fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
                });
                slide.addText(ma.title, { x: 1.1, y: yPos + 0.25, w: 4.5, h: 0.35, fontFace: FONT_TITLE, fontSize: 16, color: C_BRICK_DARK, bold: true });
                slide.addText(ma.desc, { x: 5.5, y: yPos + 0.25, w: 6.7, h: 0.9, fontFace: FONT_BODY, fontSize: 12, color: C_DARK_TEXT, lineSpacing: 18 });
            });
        } else if (data.patterns) {
            data.patterns.forEach((pp, idx) => {
                const row = Math.floor(idx / 2);
                const col = idx % 2;
                const cardX = 0.8 + col * 5.9;
                const cardY = 1.6 + row * 2.5;

                slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                    x: cardX, y: cardY, w: 5.633, h: 2.2, rectRadius: 0.08,
                    fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
                });
                slide.addText(pp.title, { x: cardX + 0.3, y: cardY + 0.3, w: 5.0, h: 0.4, fontFace: FONT_TITLE, fontSize: 16, color: C_BRICK, bold: true });
                slide.addText(pp.desc, { x: cardX + 0.3, y: cardY + 0.8, w: 5.0, h: 1.1, fontFace: FONT_BODY, fontSize: 12, color: C_MUTED_TEXT, lineSpacing: 18 });
            });
        } else {
            const imgPath = path.join(IMG_DIR, data.img);
            if (fs.existsSync(imgPath)) slide.addImage({ path: imgPath, x: 0.8, y: 1.6, w: 6.2, h: 4.9 });

            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: 7.2, y: 1.6, w: 5.333, h: 4.9, rectRadius: 0.08,
                fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
            });

            slide.addText("연동 요약 통계", { x: 7.4, y: 1.8, w: 4.9, h: 0.3, fontFace: FONT_TITLE, fontSize: 14, color: C_DARK_TEXT, bold: true });
            slide.addTable([data.headers, ...data.tableData], {
                x: 7.4, y: 2.15, w: 4.933, colW: [1.3, 1.3, 2.333],
                fontFace: FONT_BODY, fontSize: 8.5, color: C_DARK_TEXT,
                headerFill: C_BRICK, headerColor: "FFFFFF", align: "center", fill: C_CARD_BG
            });

            slide.addText("비즈니스 인사이트 및 해석", { x: 7.4, y: 3.8, w: 4.9, h: 0.3, fontFace: FONT_TITLE, fontSize: 14, color: C_BRICK, bold: true });
            slide.addText(data.insight, { x: 7.4, y: 4.15, w: 4.933, h: 2.2, fontFace: FONT_BODY, fontSize: 11, color: C_DARK_TEXT, lineSpacing: 16 });
        }

        slide.addNotes(`[발표 대본 - 약 2분 분량]\n${data.notes}`);
    }

    // =========================================================================
    // SLIDE 27: 2026 생성형 AI 키워드 4대 테마
    // =========================================================================
    {
        const slide = createNordicSlide("3.11 [Chart 11] 2026 생성형 AI 키워드 4대 테마");

        const aiThemes = [
            { icon: "ai", theme: "1. AI 에이전트 & 자동화", keywords: "AI, 에이전트, 코워크, 클로드 코드", desc: "스스로 판단하고 코드를 작성하여 업무를 수행하는 AI 에이전트 서적의 부상" },
            { icon: "review", theme: "2. 프롬프트 엔지니어링", keywords: "챗GPT, 클로드, 제미나이, 활용법", desc: "챗GPT, 클로드 등 3대 LLM 활용 꿀팁 및 프롬프트 작성 실무서 고속 성장" },
            { icon: "target", theme: "3. 바이브 코딩 & 파이썬", keywords: "코딩, 파이썬, 데이터, 백엔드", desc: "비전공자도 AI의 도움으로 앱/웹을 개발하는 바이브 코딩 파이프라인 인기" },
            { icon: "book", theme: "4. 실무 디자인 & 엑셀", keywords: "디자인, 엑셀, 가이드, 실무", desc: "생성형 AI를 엑셀 데이터 분석 및 디자인 작업에 결합하는 실용서 수요 폭발" }
        ];

        aiThemes.forEach((at, idx) => {
            const row = Math.floor(idx / 2);
            const col = idx % 2;
            const cardX = 0.8 + col * 5.9;
            const cardY = 1.6 + row * 2.5;

            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: cardX, y: cardY, w: 5.633, h: 2.2, rectRadius: 0.08,
                fill: { color: C_CARD_BG }, line: { color: C_BORDER, width: 1 }
            });

            if (flatIconBase64Map[at.icon]) {
                slide.addImage({ data: flatIconBase64Map[at.icon], x: cardX + 0.3, y: cardY + 0.25, w: 0.45, h: 0.45 });
            }

            slide.addText(at.theme, {
                x: cardX + 0.85, y: cardY + 0.25, w: 4.5, h: 0.35,
                fontFace: FONT_TITLE, fontSize: 15, color: C_DARK_TEXT, bold: true
            });
            slide.addText(`핵심키워드: ${at.keywords}`, {
                x: cardX + 0.3, y: cardY + 0.7, w: 5.0, h: 0.3,
                fontFace: FONT_BODY, fontSize: 11, color: C_BRICK, bold: true
            });
            slide.addText(at.desc, {
                x: cardX + 0.3, y: cardY + 1.05, w: 5.0, h: 1.0,
                fontFace: FONT_BODY, fontSize: 12, color: C_MUTED_TEXT, lineSpacing: 17
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
TF-IDF 키워드 분석에서 도출된 2026년 도서 시장 4대 AI 테마 분석입니다.`);
    }

    // =========================================================================
    // SLIDE 28: [4. 종합 제언] 20년차 분석가의 4대 핵심 파인딩 요약
    // =========================================================================
    {
        const slide = createNordicSlide("4.1 20년차 데이터 분석가의 4대 핵심 파인딩 요약");

        const coreFindings = [
            { icon: "price", num: "FINDING 01", title: "체감 가격 저항선 & 정가 설정", desc: "베스트셀러의 60% 이상이 정가 20,000~30,000원 구간에 집중되며 10% 정률 할인이 표준화됨" },
            { icon: "chart", num: "FINDING 02", title: "흥행 비대칭성 & 롱테일 관리", desc: "상위 5% 메가 베스트셀러가 전체 판매량을 독점함. 메가작의 물류 리드타임 제로화가 필수적임" },
            { icon: "ai", num: "FINDING 03", title: "AI 생성형 트렌드 키워드 지배", desc: "AI, 클로드, 챗GPT, 에이전트 등 최신 기술 키워드가 시장 독자 수요 장악. 트렌드 키워드 선점 필수" },
            { icon: "star", num: "FINDING 04", title: "초기 평점 9.5+ & 리뷰 30건 축적", desc: "평점 9.5점 이상 확보는 기본 조건이며, 리뷰 볼륨 30건 이상 누적이 구전 효과(WOM)를 창출함" }
        ];

        coreFindings.forEach((cf, idx) => {
            const row = Math.floor(idx / 2);
            const col = idx % 2;
            const cardX = 0.8 + col * 5.9;
            const cardY = 1.6 + row * 2.5;

            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: cardX, y: cardY, w: 5.633, h: 2.2, rectRadius: 0.08,
                fill: { color: C_CARD_BG }, line: { color: C_BRICK, width: 1 }
            });

            if (flatIconBase64Map[cf.icon]) {
                slide.addImage({ data: flatIconBase64Map[cf.icon], x: cardX + 0.3, y: cardY + 0.25, w: 0.45, h: 0.45 });
            }

            slide.addText(cf.num, {
                x: cardX + 0.85, y: cardY + 0.2, w: 4.5, h: 0.25,
                fontFace: FONT_BODY, fontSize: 10, color: C_BRICK, bold: true
            });
            slide.addText(cf.title, {
                x: cardX + 0.85, y: cardY + 0.45, w: 4.5, h: 0.35,
                fontFace: FONT_TITLE, fontSize: 15, color: C_DARK_TEXT, bold: true
            });
            slide.addText(cf.desc, {
                x: cardX + 0.3, y: cardY + 0.95, w: 5.0, h: 1.1,
                fontFace: FONT_BODY, fontSize: 11.5, color: C_MUTED_TEXT, lineSpacing: 17
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
4.1 절에서는 이번 EDA의 4대 핵심 결론(가격, 롱테일, AI 키워드, 리뷰)을 종합 요약합니다.`);
    }

    // =========================================================================
    // SLIDE 29: [4. 전략적 제언] 3단계 실행 로드맵
    // =========================================================================
    {
        const slide = createNordicSlide("4.2 출판사 & 유통사 3단계 전략 실행 로드맵");

        const roadmapSteps = [
            { icon: "target", phase: "PHASE 1. 출간 전 기획 (Pre-Launch)", title: "트렌드 키워드 & 적정 정가 설정", desc: "• TF-IDF 상위 AI/자동화 키워드를 도서명 및 부제에 타깃 반영\n• 22,000원~28,000원 정가 책정으로 저항선 최소화" },
            { icon: "checklist", phase: "PHASE 2. 출간 초 1개월 (Launch)", title: "초기 평점 & 리뷰 30건 조기 확보", desc: "• 출간 즉시 서평단 50명 운영으로 평점 9.5점 이상 방어\n• 독점 굿즈 및 사은품 연계로 초기 결제 및 리뷰 촉진" },
            { icon: "trend", phase: "PHASE 3. 3개월 이후 (Post-Launch)", title: "S등급 메가작 재고 & 바이럴 확장", desc: "• 판매지수 상위 5% 메가작 집중 물류 리드타임 제로화\n• 저자 직강 VOD 및 숏폼 영상 마케팅으로 스테디셀러화" }
        ];

        roadmapSteps.forEach((rms, idx) => {
            const cardX = 0.8 + idx * 3.95;
            slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
                x: cardX, y: 1.6, w: 3.8, h: 4.9, rectRadius: 0.08,
                fill: { color: idx === 1 ? C_BRICK_BG : C_CARD_BG },
                line: { color: idx === 1 ? C_BRICK : C_BORDER, width: 1.5 }
            });

            if (flatIconBase64Map[rms.icon]) {
                slide.addImage({ data: flatIconBase64Map[rms.icon], x: cardX + 0.3, y: 1.85, w: 0.45, h: 0.45 });
            }

            slide.addText(rms.phase, {
                x: cardX + 0.85, y: 1.9, w: 2.7, h: 0.3,
                fontFace: FONT_BODY, fontSize: 10.5, color: C_BRICK_DARK, bold: true
            });
            slide.addText(rms.title, {
                x: cardX + 0.3, y: 2.4, w: 3.4, h: 0.7,
                fontFace: FONT_TITLE, fontSize: 15, color: C_DARK_TEXT, bold: true, lineSpacing: 18
            });
            slide.addText(rms.desc, {
                x: cardX + 0.3, y: 3.2, w: 3.4, h: 3.0,
                fontFace: FONT_BODY, fontSize: 12, color: C_DARK_TEXT, lineSpacing: 18
            });
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
4.2 절에서는 출판사와 서점이 즉각 실행할 수 있는 3단계 Actionable 로드맵을 제언합니다.`);
    }

    // =========================================================================
    // SLIDE 30: Q&A 및 종결 슬라이드 (Nordic Theme)
    // =========================================================================
    {
        const slide = pptx.addSlide();
        slide.background = { color: C_NORDIC_BG };

        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 1.0, y: 1.2, w: 11.333, h: 4.8, rectRadius: 0.08,
            fill: { color: C_CARD_BG },
            line: { color: C_BRICK, width: 1.5 }
        });

        slide.addText("Q & A", {
            x: 1.5, y: 1.8, w: 10.333, h: 0.8,
            fontFace: FONT_TITLE, fontSize: 40, color: C_BRICK_DARK, bold: true, align: "center"
        });

        slide.addText("경청해 주셔서 감사합니다.", {
            x: 1.5, y: 2.8, w: 10.333, h: 0.6,
            fontFace: FONT_TITLE, fontSize: 24, color: C_DARK_TEXT, align: "center"
        });

        slide.addText("YES24 베스트셀러 EDA 노르딕 & 벽돌색 에디션 30페이지 발표를 마칩니다.\n질의응답 및 추가 분석 요청 사항을 말씀해 주시기 바랍니다.", {
            x: 1.5, y: 3.6, w: 10.333, h: 1.0,
            fontFace: FONT_BODY, fontSize: 15, color: C_MUTED_TEXT, align: "center", lineSpacing: 22
        });

        slide.addText("분석 담당자: 20년차 데이터 분석 팀 | 문의: analytics@yes24-eda.com", {
            x: 1.0, y: 6.4, w: 11.333, h: 0.4,
            fontFace: FONT_BODY, fontSize: 11, color: C_MUTED_TEXT, align: "center"
        });

        slide.addNotes(`[발표 대본 - 약 2분 분량]
이상으로 YES24 베스트셀러 탐색적 데이터 분석(EDA) 노르딕 에디션 발표를 마칩니다. 질문이나 추가 요청 사항에 대해 답변해 드리겠습니다. 감사합니다.`);
    }

    // 5. 파일 저장 실행
    const outputPath = path.join(__dirname, "YES24_Bestseller_EDA_Presentation.pptx");
    await pptx.writeFile({ fileName: outputPath });
    console.log(`[성공] 노르딕 스타일 & 벽돌색 포인트가 적용된 30페이지 PPTX 생성이 완료되었습니다: ${outputPath}`);
}

buildPresentation().catch(err => {
    console.error(`[오류] PPTX 파일 생성 중 에러가 발생했습니다:`, err);
});
