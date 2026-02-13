# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**PS-GradeM** — 경찰/소방 현장 모의고사 성적 분석 및 합격 예측 시스템. 학원 수강생이 모의고사 성적, 석차, 과목별 약점, 합격 가능성을 시각적으로 확인하는 웹 애플리케이션.

## Tech Stack

- **Backend:** Google Apps Script (GAS), JavaScript (ES5 호환)
- **Frontend:** Vanilla JavaScript + HTML5 + CSS3 (프레임워크 없음)
- **Database:** Google Sheets (시트 기반 관계형 데이터)
- **Charting:** Chart.js 3.x (CDN)
- **Font:** Noto Sans KR (Google Fonts)

## Architecture

단일 GAS 컨테이너 바운드 스크립트에 3개 파일로 구성된 SPA:

- **Code.gs** (~1,700줄) — 서버사이드 로직 전체. CONFIG 객체로 시트명/컬럼 인덱스 관리. 모든 공개 함수는 `{ success, message, data }` 형태로 응답 반환.
- **Index.html** (~2,100줄) — 로그인, 시험 선택, 대시보드(7개 분석 탭)를 포함한 SPA. `google.script.run`으로 백엔드 호출.
- **Css.html** (~1,250줄) — Flat design (box-shadow 없음), 기본색 `#007bff`, 배경 `#ebf1ff`.
- **Admin.html** — 관리자 패널 (학생 등록, 일괄 등록, 시험 관리).

### 데이터 흐름

1. 로그인 → `loginUser()` 인증
2. 시험 목록 → `getStudentExamList()`
3. 시험 선택 → `getScoreAnalysis()` (모든 분석 데이터 한 번에 계산)
4. 프론트엔드에서 탭 전환 시 추가 API 호출 없이 캐시된 `currentData`로 렌더링

### Google Sheets 스키마

| 시트 | 주요 컬럼 | 용도 |
|------|----------|------|
| Students_Police | 수험번호, 이름, 직렬, 지원지역, 비밀번호, 연락처... | 경찰 수험생 |
| Students_Fire | (동일 구조) | 소방 수험생 |
| Exams | 시험ID, 시험명, 시행일, 과목1~3, 합격예상점수 | 시험 메타데이터 |
| Scores | 시험ID, 수험번호, 과목1~3점수, 총점, 평균, [문항응답...] | 성적 + 문항별 O/X |
| ItemAnalysis | 시험ID, 문항번호, 정답, 배점, 정답률, 난이도 | 문항 분석 |
| 로그인기록 | 이름, 연락처, 접속시간, 작업, IP정보 | 감사 로그 |

## Development & Deployment

GAS 프로젝트이므로 npm/빌드 도구 없음. 두 가지 개발 방식:

**방법 1 — GAS 에디터:** Google Sheets > 확장 프로그램 > Apps Script에서 직접 편집 후 배포.

**방법 2 — clasp (로컬 개발):**
```bash
npm install -g @google/clasp
clasp login
clasp clone <scriptId>
clasp push          # 로컬 파일 → GAS 업로드
clasp deploy        # 웹 앱 배포
```

**초기화:** Google Sheets 메뉴 "성적 시스템 관리" → "데이터 초기화 (시트 생성 + 목업 데이터)" 실행 → `initializeSampleData()`로 5개 시트에 샘플 데이터 생성.

**테스트 로그인:** 수험번호 `21889`, 이름 `이동근`, 유형 `경찰`, 비밀번호 `1234`.

## Code Conventions

- **CONFIG 객체:** 시트명, 컬럼 인덱스(STUDENT_COL, EXAM_COL, SCORE_COL, ITEM_COL) 등 모든 설정값을 최상단에 선언
- **Private 함수:** 내부 헬퍼는 trailing underscore 사용 (예: `getSheet_()`, `defaultPassword_()`)
- **응답 형식:** 모든 공개 GAS 함수는 `{ success: boolean, message: string, data: object }` 반환
- **프론트엔드 전역 상태:** `currentData`에 분석 결과 캐시, 탭 전환 시 재사용
- **Chart.js 관리:** 차트 인스턴스를 반드시 destroy 후 재생성 (메모리 누수 방지)
- **CSS:** Flat design 원칙 — box-shadow 사용 금지, 흰색 배경 기반

## Exam Type Logic

- 시험ID가 'P'로 시작 → 경찰, 'F'로 시작 → 소방
- 과목명 분석 폴백: "경찰학" → 경찰, "소방학" → 소방
- 소방 세부 트랙: 공채, 구급특채, 구조특채

## Key Formulas

- **석차:** 총점 내림차순, 동점자 처리 포함
- **백분위:** `(lowerCount + 0.5 * tieCount) / totalCount * 100`
- **상위%:** `(rank - 1) / totalCount * 100`
- **과목 등급:** 표본 < 10명이면 득점률 기준(≥80% → 우수), ≥ 10명이면 백분위 기준(≤30% → 우수)
- **합격권:** 상위% 기준 S(5%)/A(15%)/B(30%)/C(30%+)
