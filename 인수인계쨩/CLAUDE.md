# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

성과관리 시스템 인수인계 문서를 단일 HTML 파일로 구현한 대화형 가이드 시리즈. 빌드 도구 없음 — 브라우저에서 직접 열어 확인한다.

```
index.html   ← 초기 버전 (541줄, 애니메이션 없음)
index2.html  ← 레이아웃 개선
index3.html  ← 카드 디자인 + 섹션 pill 도입
index4.html  ← 파이프라인 다이어그램 + 배치 타임라인 추가
index5.html  ← 현행 최신 (1767줄, Chart.js 도넛 차트 + 전체 애니메이션)
```

## 실행

별도 서버 불필요. 파일을 브라우저로 열거나:

```bash
open index5.html
```

## CDN 의존성 (index5.html 기준)

- **Tailwind CSS** — `https://cdn.tailwindcss.com` (유틸리티 클래스)
- **Chart.js 4.4.0** — `https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js` (도넛 차트)
- **Pretendard** — Google Fonts (한글 폰트)

index3 이하에는 Chart.js 없음.

## 문서 구조 (index5.html)

6개 섹션, 단일 페이지 스크롤:

| 앵커 | 섹션명 |
|------|--------|
| `#s1` | 시스템 구성 및 데이터 파이프라인 |
| `#s2` | 성과지표 산출 로직 |
| `#s3` | 특정 고객사 예외 처리 규칙 |
| `#s4` | 산출 상세 예시 — 그룹A 시뮬레이션 |
| `#s5` | 데이터 배치 및 릴리즈 일정 |
| `#s6` | 매월 기본 작업 리스트 |

## 주요 CSS 패턴

- `.card` / `.card-header` / `.card-body` — 흰색 라운드 카드 컴포넌트
- `.sec-pill` / `.step-pill` — 섹션·스텝 번호 레이블 (sec-pill은 H2급, font-size 20px font-weight 800)
- `.badge .b-{color}` — 인라인 컬러 태그 (blue/green/amber/purple/slate)
- `.formula-block` — 이익 공식 시각화 블록 (result + op + term 구조)
- `.exc-card` / `.exc-head` / `.exc-body` — 예외 규칙 카드
- `.grid-2col`, `.grid-3col` 등 — 커스텀 그리드 (반응형: 768px 이하 1열)
- `.reveal` + IntersectionObserver — 스크롤 시 fade-in
- `.tbl-row-reveal` — 테이블 행 stagger 진입 애니메이션
- `.counter` + `data-target` / `data-suffix` / `data-float` / `data-comma` — 카운트업 숫자 애니메이션
- `.pipe-node` / `.pipe-connector` / `.flow-dot` — 파이프라인 다이어그램 + 흐름 점 애니메이션
- `.batch-step` / `.batch-connector-line` — 배치 타임라인 세로 진입 애니메이션

## JavaScript 구조 (index5.html `<script>`)

인라인 스크립트 9개 블록:

1. **Scroll-reveal** — `.reveal` 요소 IntersectionObserver
2. **Counter animation** — `.counter` 요소 카운트업
3. **Pipeline hover + auto-play** — 스크롤 진입 시 파이프라인 순차 점등, 마우스오버로 설명 변경
4. **Formula highlight** — 수익성 공식 항 순차 하이라이트
5. **VRB table rows stagger** — `#vrbTableBody` 행 순차 등장
6. **Calc section reveal** — 계산 박스 라인별 등장 + counter 트리거
7. **Step3 table rows stagger** — `#step3Card` 행 순차 등장
8. **Batch timeline** — 배치 스텝 순차 등장 + connector 성장 애니메이션
9. **Nav active highlight** — 스크롤 위치에 따라 상단 네비 링크 강조

Chart.js(`donutChart`)는 `#calcSection` 진입 시 지연 생성 (`chartsCreated` 플래그로 중복 방지).

## 콘텐츠 규칙 (수정 시 주의)

- **반영개월수 계산 기준**: '26년 1~5월 중 심의가 유효한 개월 수. 🟡 26년 내 종료 / 🔵 26년 이후 종료 / 🟢 26년 2월 이후 시작 / 🟣 일회성
- **VRB 가중평균**: 반영이익 합계 ÷ 반영매출 합계 (단순평균 ❌)
- **예외 고객**: 네이버 코람코 IDC (Billing 역산), 한전 KEPCO 전력통신 (PL 이익률 강제 일괄)
- **수기 보정 필수 고객**: 에스원(M21/M2M), KGM(T35/커넥티드카), 토요타(T35/커넥티드카) — 월평균 4~5억 원
- **배치 기준일**: 매월 11일 05:00~, 공휴일 무관

## 매월 기본 작업 흐름 (s6)

1. **데이터 생성 전** — 수기테이블 정비 (DW 매출·PL 자료 수령 즉시 대응)
   - 에스원(M21/M2M)·KGM(T35/커넥티드카)·토요타(T35/커넥티드카) 보정매출 → 매출조정 수기테이블 기입
   - 기획팀 PL 자료 수령 시 → 한전 PL 수익률 수기테이블 완성
2. **데이터 생성 후 (IF 미완성)** — 수기테이블 완성 후 데이터 재생성
3. **데이터 생성 후 (ELSE 완성)** — KA 고객 매출·이익·심의내역 추출 → 네이버클라우드·한전 수익성 조정 → 특이사항 Eye-Checking (Claude로 양식 생성)

## 푸시 대상 레포

- **작업 레포**: `https://github.com/konslie/boanjiyeon` (origin)
- **인수인계 문서 배포 레포**: `https://github.com/konslie/LGU_OH_KO` (remote명: lgu) ← 이쪽으로 push할 것
