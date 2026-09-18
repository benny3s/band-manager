# 밴드매니저 — 합주 일정 취합

관리자가 **날짜 후보**를 올리면, 멤버들이 **가능한 시간을 칠하고**,
관리자는 **전원(또는 N명) 되는 시간**을 한눈에 보고 합주실을 예약한다.

GitHub Pages(정적 페이지) + Google Apps Script(저장) + Google 스프레드시트(데이터).

```
멤버 휴대폰 ──► GitHub Pages (index.html)
                      │  JSONP (GET + callback)
                      ▼
              Apps Script 웹앱 (/exec)
                      │
                      ▼
       스프레드시트  설정 / 응답 / 메모  3개 시트
                      │
                      ▼
            Claude 가 읽고 집계 → 그루브 예약
```

## 화면

| 탭 | 누가 | 하는 일 |
|---|---|---|
| 응답하기 | 멤버 | 이름 고르고 날짜×시간 격자를 손가락으로 쭉 칠함. 날짜별 메모 한 줄 |
| 결과 | 누구나 | 응답 현황, 추천 시간대(연속 2시간 이상), 인원수 히트맵, 메모 모음 |
| 관리 | 관리자 | 날짜 후보·멤버·시간 범위·제목. **`#admin` 을 URL 뒤에 붙여야 보임** |

- 멤버용 링크: `https://benny3s.github.io/band-manager/`
- 관리자용 링크: `https://benny3s.github.io/band-manager/#admin`

`#admin` 은 비밀번호가 아니라 **실수 방지용**이다. 저장소가 공개라 소스를 보면 누구나 알 수 있다.

## 조작

- **드래그**: 칸을 누른 채 끌면 연속 선택. 이미 칠해진 칸에서 시작하면 지우기 모드
- **날짜 헤더 클릭**: 그 날 전체 토글 / **시간 라벨 클릭**: 그 시간 전체 토글
- **자동 저장**: 변경 1.2초 뒤. 하단 상태바에 `저장됨 ✓`
- 이름은 localStorage 에 기억 → 재방문 시 자기 응답을 불러와 수정

## 데이터 형식

**`설정` 시트** — A=key, B=value

| key | 예 |
|---|---|
| `dates` | `2026-10-04,2026-10-11,2026-10-18` |
| `hourStart` / `hourEnd` | `9` / `22` (끝은 종료시각, 21시 칸이 마지막) |
| `title` | `애쉬만루트 합주` |

**`응답` 시트** — A=이름, B열부터 날짜(`YYYY-MM-DD`)

| 값 | 뜻 |
|---|---|
| `10,11,12` | 10·11·12시 칸 가능 (= 10:00~13:00) |
| `-` | 그 날 전부 불가 (**응답은 함**) |
| 빈칸 | 아직 응답 안 함 |

구버전 값(`O` / `X` / `09-12`)도 읽어준다. 저장할 때만 새 형식으로 바뀐다.

**`메모` 시트** — 응답 시트와 같은 배치, 값은 그 날에 대한 한 줄 사유.

## 배포

### 페이지
```
git add -A
git -c user.name="benny03-song" -c user.email="benny03-song@users.noreply.github.com" commit -m "..."
git push
```
1~2분 뒤 반영.

### Apps Script
시트 → 확장 프로그램 → Apps Script → `apps-script.gs` 붙여넣기 → Ctrl+S →
**배포 → 배포 관리 → 연필(수정) → 버전: 새 버전 → 배포**

⚠️ **'새 배포'를 쓰면 URL이 바뀐다.** 반드시 기존 배포를 수정할 것.

에디터에 코드 밀어넣기(Monaco 직접 주입):
```js
const code = await (await fetch("https://raw.githubusercontent.com/benny3s/band-manager/main/apps-script.gs",{cache:"no-store"})).text();
monaco.editor.getModels()[0].setValue(code);
```

## ⚠️ CORS — 반드시 JSONP

Apps Script `/exec` 는 `script.googleusercontent.com` 으로 302 리다이렉트되는데,
**브라우저 `fetch` 는 그 리다이렉트에서 응답이 pending 으로 멈춘다**(실측 확인).
그래서 페이지는 `<script>` 태그 JSONP(`?callback=`)로만 통신한다.
`doPost` 는 서버 대 서버(curl/PowerShell) 전용 — 브라우저에서 쓰지 말 것.

## 주의

- 저장소는 **공개**다(무료 GitHub Pages 제약). `/exec` 주소가 노출되므로
  링크가 단톡방 밖으로 새면 낙서가 가능하다. 그때는 Apps Script 를 새로 배포하고
  `index.html` 의 `ENDPOINT` 를 교체한다.
- 응답 데이터 자체는 스프레드시트에만 있고 저장소에는 남지 않는다.
- 사고 시 스프레드시트 **파일 → 버전 기록**으로 복구.
