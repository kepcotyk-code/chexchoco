책스초코 — Vercel 배포 가이드
=====================================

이 폴더는 클로드 아티팩트가 아니라 완전히 독립된 웹앱이에요.
Supabase(데이터/사진 저장) + Vercel(웹사이트 호스팅)로 동작해요.

[준비 완료 확인]
- Supabase 테이블 9개 생성 완료 ✅
- Storage 버킷 photos, notice-files 생성 (아직이면 지금 만들어주세요)
- src/supabaseClient.js 에 Project URL과 키가 이미 입력되어 있어요

[1단계] GitHub에 코드 올리기
1. github.com 접속 (아까 만드신 계정으로 로그인)
2. 우측 상단 + 버튼 → "New repository"
3. Repository name: chexchoco (아무 이름이나 가능)
4. Public 또는 Private 아무거나 선택 → "Create repository"
5. 생성된 저장소 화면에서 "uploading an existing file" 링크 클릭
6. 이 폴더(vercel-app) 안의 모든 파일과 폴더를 통째로 드래그해서 업로드
   (node_modules 폴더는 없으니 신경 안 쓰셔도 돼요)
7. 하단 "Commit changes" 클릭

[2단계] Vercel 배포
1. vercel.com 접속 → "Continue with GitHub"로 가입/로그인
2. "Add New..." → "Project" 클릭
3. 방금 만든 GitHub 저장소(chexchoco) 선택 → "Import"
4. Framework Preset은 자동으로 "Vite"로 잡힐 거예요 (안 잡히면 직접 Vite 선택)
5. 다른 설정은 그대로 두고 "Deploy" 클릭
6. 1~2분 기다리면 배포 완료, 실제 URL(예: chexchoco.vercel.app)이 생겨요!

[3단계] 확인
- 생성된 URL을 브라우저(크롬 등)에서 열어보세요
- 사용자관리 화면에서 첫 멤버 등록부터 시작하면 돼요
- 이 링크는 클로드와 완전히 무관하니, 카카오톡 등으로 자유롭게 공유하세요

[이후 업데이트 방법]
1. 코드가 바뀌면, 바뀐 파일을 GitHub 저장소에 다시 업로드(덮어쓰기)
2. Vercel이 자동으로 감지해서 새로 배포 (몇 초~1분)
3. 링크 새로고침하면 최신 버전 반영

[데이터 백업]
- 사용자관리 → "명단 다운로드"로 언제든 엑셀 백업 가능
- 출석관리(간사) → "이번 달 다운로드"로 출석기록 엑셀 백업 가능
- 원본 데이터는 Supabase에 안전하게 저장되어 있어서, PC를 껐다 켜도 전혀 영향 없어요
