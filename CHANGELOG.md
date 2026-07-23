# Changelog

이 프로젝트의 모든 주요 변경 사항을 이 파일에 기록합니다.

형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/)를 따르며,
버전 표기는 [Semantic Versioning](https://semver.org/lang/ko/)을 따릅니다.

## [1.0.0] - 2026-07-23

### Added
- HWP·HWPX → PDF 일괄 변환 (드래그 & 드롭 GUI)
- CLI 모드: 파일·폴더 인자로 GUI 없이 일괄 변환 (종료 코드 0 / 2 / 1)
- 폴더 재귀 수집, 중복 파일 자동 제거, 파일별 진행 상태 표시
- 기존 PDF 덮어쓰기 옵션, `Delete` 키로 목록 항목 제거
- 한글 오토메이션 늦은 바인딩(IDispatch) 방식 — 한글 2010~2024 호환
- 앱 아이콘 (exe · 창 제목표시줄 · 작업표시줄)

[1.0.0]: ../../releases/tag/v1.0.0
