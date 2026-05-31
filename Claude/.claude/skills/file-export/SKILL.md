---
name: file-export
description: 마크다운을 HTML로, JSON을 엑셀로 변환하는 스크립트 실행 방법
disable-model-invocation: true
allowed-tools: Bash(python *)
---

# 파일 변환 스킬

## HTML 변환 (.md → .html)

```bash
python "${CLAUDE_SKILL_DIR}/scripts/md_to_html.py" "<md_파일_경로>" "<출력_html_경로>"
```

예시:
```bash
python "${CLAUDE_SKILL_DIR}/scripts/md_to_html.py" "output/2026-04-25/AI_에이전트_트렌드_20260425.md" "output/2026-04-25/AI_에이전트_트렌드_20260425.html"
```

## 엑셀 변환 (analysis.json → .xlsx)

```bash
python "${CLAUDE_SKILL_DIR}/scripts/json_to_excel.py" "<analysis_json_경로>" "<출력_xlsx_경로>"
```

예시:
```bash
python "${CLAUDE_SKILL_DIR}/scripts/json_to_excel.py" "output/2026-04-25/analysis.json" "output/2026-04-25/AI_에이전트_트렌드_20260425.xlsx"
```

## 엑셀 시트 구조
- **요약 시트**: 키워드, 생성일, 핵심 요약
- **출처 목록 시트**: 번호, 제목, URL, 날짜, 중요도

## 라이브러리 설치 (미설치 시)
```bash
pip install markdown openpyxl
```

## 실패 처리 원칙
- HTML 변환 실패해도 엑셀 변환은 계속 진행
- 각 변환은 독립적으로 처리
- 오류 메시지를 오케스트레이터에게 보고
