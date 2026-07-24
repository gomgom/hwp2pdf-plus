# Changelog

이 프로젝트의 모든 주요 변경 사항을 이 파일에 기록합니다.

형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/)를 따르며,
버전 표기는 [Semantic Versioning](https://semver.org/lang/ko/)을 따릅니다.

## [1.0.1] - 2026-07-24

### Fixed
- PDF 변환 시 한/글에 저장돼 있던 **마지막 인쇄/변환 범위**(예: 직전에 특정 페이지만 인쇄)를 물려받아 문서 일부만 저장되던 문제 수정. 이제 잔존 설정과 무관하게 **항상 문서 전체**가 PDF로 저장됩니다. (`SaveAs` 필터 대신 가상 인쇄 `PrintToPDFEx` + 기본값 강제 방식으로 변경, 모아찍기 해제)

## [1.0.0] - 2026-07-23

### Added
- HWP·HWPX → PDF 일괄 변환 (드래그 & 드롭 GUI)
- CLI 모드: 파일·폴더 인자로 GUI 없이 일괄 변환 (종료 코드 0 / 2 / 1)
- 폴더 재귀 수집, 중복 파일 자동 제거, 파일별 진행 상태 표시
- 기존 PDF 덮어쓰기 옵션, `Delete` 키로 목록 항목 제거
- 한글 오토메이션 늦은 바인딩(IDispatch) 방식 — 한글 2010~2024 호환
- 앱 아이콘 (exe · 창 제목표시줄 · 작업표시줄)

[1.0.1]: ../../releases/tag/v1.0.1
[1.0.0]: ../../releases/tag/v1.0.0
