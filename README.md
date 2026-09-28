# 밴드매니저

밴드 합주 일정 취합 · 합주곡 · 합주 이력. https://benny3s.github.io/band-manager/

- 오른쪽 위에서 **밴드 고르기** (목록은 누구나 봄). PIN 이 걸린 밴드는 PIN 을 한 번 넣으면 그 기기에 기억
- **📅 일정** — 시간표(여러 개)마다 멤버가 ○ △ ✕ 로 넣는다. 결과 보기에서 📌 정하면 **이력에 '예정' 회차가 자동으로** 생긴다
- **🎵 합주곡** — 합주곡 / 투표곡(♥) / 보류, 세션·링크·튜닝
- **🎸 이력** — 회차별 날짜·합주실·불참(이유)·한 곡 (처음 한 곡은 NEW)
- 링크: `?b=<밴드id>&t=<시간표id>`

## 구조 (2026-09-28 부터)
- `index.html` — 밴드 목록 · PIN · 시간표 전환 · 합주곡 · 이력
- 일정 화면과 디자인은 **약속 잡자(/when/)와 같은 파일**을 불러 쓴다: `../when/ui.css`, `core.js`, `sched.js`
  → 저장소 `benny3s/when`. 거기를 고치면 두 앱이 같이 바뀐다
- 저장소 — Firebase `benny-apps` 의 Firestore
  - `bands/{bid}` 이름·잠김 (목록용)
  - `bandData/{sha256(bid + ":" + PIN)}` 멤버 · `scheds`(시간표 = 약속 잡자의 약속과 같은 모양) · `songs` · `sessions`
    PIN 을 모르면 경로를 모른다. PIN 을 바꾸면 데이터를 새 경로로 옮긴다
  - 규칙: 저장소 `benny3s/benny-apps` 의 `firestore.rules`
- `apps-script.gs` — **옛 서버 (2026-09-28 까지)**. 시트에 남은 옛 기록을 볼 때만
