# 🍎 신선한 산지직송 사과 주문서 수정기 (AppleSell Order Fixer)

네이버폼으로 받은 사과 주문 결과 엑셀(설문 답변)을 **"신선한 산지직송 사과 주문서"**(Sheet1 형식) 수정본으로 변환하는 웹앱입니다.

## 기능

1. **엑셀 업로드**: 네이버폼 → 응답 엑셀 파일(.xlsx)을 드래그앤드롭 또는 파일 선택으로 업로드
2. **자동 변환**: 설문 답변 시트에서 주문 데이터를 추출하여
   - 사과 상품 선택 정리 (예: `5키로 (배송비 포함 5만원)` → `5키로`)
   - 사과 수량/박스 정리 (예: `1. 1박스` → `1박스`)
   - 연락처 정규화 (예: `1020556927` → `01020556927`)
   - 주문서번호 / 주문자이름 / 배송지주소 / 입금자·총입금액 / 주문서내용확인 / 개인정보수집박의 매핑
3. **수정본 미리보기**: 변환 결과를 테이블로 확인 (수정 내역 표시)
4. **수정본 번호**: 수정본 v1, v2, ... 선택 후 해당 번호로 다운로
5. **수정본 다운로**: 수정된 "신선한 산지직송 사과 주문서"(Sheet1 형식) 엑셀 생성 → `<원본파일명>_v{n}.xlsx`

## 로컬 실행

```bash
# 정적 서버 예시 (로컬에서 확인할 때)
python3 -m http.server 8000
# → http://localhost:8000

# 또는 간단한 파일 서버
npx serve .
```

> 브라우저에서는 SheetJS CDN(`https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js`)를 사용합니다.

## GitHub Pages 배포

이 저장소는 GitHub Actions를 통해 **GitHub Pages**에 자동 배포됩니다.

- main 브랜치에 push → `.github/workflows/deploy.yml` 실행 → GitHub Pages 배포
- 배포 URL: `https://<사용자명>.github.io/Applesell/`

> GitHub 저장소 Settings → Pages에서 Source를 **GitHub Actions**로 설정해야 합니다.

## 사용 방법

1. GitHub Pages URL 접속
2. 네이버폼 응답 엑셀 업로드
3. 수정본 미리보기 확인 (수정본 번호 v1 선택 가능)
4. "수정본 다운로" 클릭 → 수정된 주문서 파일 다운로

## 파일 구조

```
/
├── index.html
├── style.css
├── app.js
└── .github/workflows/deploy.yml
```
