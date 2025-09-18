# 착공도서 자동생성 시스템 (Vercel 배포용 최종)

이 버전은 **실제 PPTX 파일**을 브라우저에서 생성합니다. (PptxGenJS 사용)

## ✅ 변경 요약
- 파일명 통일: `index.html`, `styles.css`, `app.js`
- JSON 대신 **PPTX 생성**으로 변경
- **미니맵 빨간 박스**를 장면별로 드로잉하여 저장/삽입
- Vercel 정적 배포 설정보완 (빌드 없음)

## 사용 방법
1. 엑셀 스펙리스트, 미니맵 이미지, 장면 이미지들을 업로드
2. 좌측에서 **장면 선택 → 미니맵 캔버스에서 위치 박스 드로잉**
3. 자재 체크박스로 선택
4. **PPT 생성하기** 클릭 → `착공도서.pptx` 다운로드

## 로컬 실행
```bash
npx serve .
```
브라우저로 `http://localhost:3000` 접속

## Vercel
- GitHub 저장소에 업로드 후 Vercel 연결
- 빌드 단계 없음 (정적 호스팅)

## 기술 스택
- SheetJS (엑셀 파서)
- PptxGenJS (PPT 생성)
- HTML/CSS/JS (ES5호환)
