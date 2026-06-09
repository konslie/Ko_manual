# 인수인계쨩 — Work Log

## 2026-05-21 | 초기 구현 → index5 완성

### 진행 경과
- index.html (v1, 541줄): 기본 레이아웃, 애니메이션 없음
- index2.html: 레이아웃 개선, 배지·테이블 스타일 정비
- index3.html: 카드 컴포넌트 + sec-pill 도입, 파이프라인 다이어그램 추가
- index4.html: 배치 타임라인, 계산 예시 섹션 추가
- index5.html (현행): Chart.js 도넛 차트, 전체 스크롤 애니메이션, 반응형 완성

### 주요 결정사항
- 단일 HTML 파일 구조 유지 (빌드 도구 없음, 브라우저 직접 오픈)
- CDN: Tailwind CSS + Chart.js 4.4.0 + Pretendard
- 섹션 5개: 파이프라인 / 산출 로직 / 예외 규칙 / 계산 예시 / 배치 일정

### 이번 세션 변경 (index5.html)
- VRB 테이블 행 색채우기: 이모지 인디케이터 제거 → row 배경색으로 대체 (amber/blue/green/purple)
- 하단 범례 및 카드 헤더 legend pill: 이모지 → 색 사각형
- 반영열(반영개월수·반영매출·반영이익) font-weight:700 bold 강조
- 4.2 손익별 카드 마우스오버 → 4.1 테이블 연동 하이라이트
  - 비대상 행: 회색(#f1f5f9) + 텍스트 흐림
  - 대상 행: 밝은 주황(#fff7ed)
  - tbody.hover-active 클래스로 제어
- 브라우저 탭 타이틀 + Hero 타이틀 → "힘내십쇼.."
- VRB 테이블 행 색 전체 컬럼 통일 (reflect-blue/reflect-purple 제거)
- 4.2 hover 하이라이트: outline 방식 폐기 → 비대상 회색/대상 주황 배경색 방식으로 전환
- 전체 폰트 사이즈 +2px 스케일업 (11px→13px 기준, body 14→16px)
