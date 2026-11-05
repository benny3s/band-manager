/* 밴드매니저 index.html 상태 로직 시험 — 2026-11-05 검토 발견들 */
const { load, test, runAll, tick, eq, ok } = require("./harness");

const T1 = "2026-10-11 17:00~20:00";

/* ── 밴드 열기 + 멤버 맞추기 (가짜 시간표 하나, 사람 문서 A·B2·C) ── */
const STUB = `
var live = { id: "", done: true, meetState: "ok", peopleCache: false };
var apiCalls = [], keyGate = {};
function buildST(){
  const pre = "bandData/" + KEY + "/scheds/" + M + "/people/", pid = {};
  for (const [p, d] of db.docs) if (p.startsWith(pre) && !pid[d.data.name]) pid[d.data.name] = p.slice(pre.length);
  return { meet: {}, pid };
}
function api(p){ apiCalls.push(p); if (p.action === "member_add") return SCHED_COL().doc(p.m).collection("people").add({ name: p.name, hours: {}, notes: {} }); return Promise.resolve({}); }
function renderAll(){ if (B && KEY && tab === "sched") syncMembers(); }
function renderBsMembers(){} function renderBandSet(){} function unwatch(){} function paintSheet(){} function saveMe(){}
function openSched(id){ M = id; live.id = id; }
function bandOf(id){ return { id, name: id }; } function watchBand(){}
async function keyOf(bid){ if (keyGate[bid]) await keyGate[bid]; return "K" + bid; }
function ls(){ return ""; } function mineAdd(){} function setUrl(){}
async function ask(){ return true; }
`;
function bandEnv(){
  const E = load([{ line: "const syncTried" }, { opt: "holdSync" }, { optLine: "let openSeq" }, "closeBand", "openBand", "syncMembers",
                  "moveMember", "renameMember", "removeMember", "parseAbs", "fmtAbs", "parseParts", { code: STUB }]);
  const put = (bid, members, people) => {
    E.db.put("bandData/K" + bid, { band: bid, members });
    E.db.put("bandData/K" + bid + "/scheds/m1", { title: "t" });
    people.forEach((n, i) => E.db.put("bandData/K" + bid + "/scheds/m1/people/p" + i, { name: n, hours: {}, notes: {} }));
  };
  put("b1", ["A", "B2", "C"], ["A", "B2", "C"]);
  E.members = () => E.db.get("bandData/Kb1").members;
  E.people = () => E.db.list("bandData/Kb1/scheds/m1/people").map(p => p.name).sort();
  E.adds = () => E.ctx.apiCalls.filter(c => c.action === "member_add").map(c => c.name);
  E.put = put;
  return E;
}

test("일정 탭에서 멤버 삭제 — 지운 이름이 멤버로 되살아나지 않는다", async () => {
  const E = bandEnv(); await E.ctx.openBand("b1");
  await E.ctx.removeMember("C"); await tick(5);
  eq(E.members(), ["A", "B2"], "멤버"); eq(E.people(), ["A", "B2"], "시간표 사람"); eq(E.adds(), [], "member_add");
});
test("일정 탭에서 이름 바꾸기 — 옛 이름이 돌아오지 않고 새 이름 문서가 둘 생기지 않는다", async () => {
  const E = bandEnv(); await E.ctx.openBand("b1");
  await E.ctx.renameMember("A", "A2"); await tick(5);
  eq(E.members(), ["A2", "B2", "C"], "멤버"); eq(E.people(), ["A2", "B2", "C"], "시간표 사람"); eq(E.adds(), [], "member_add");
});
test("다른 기기가 멤버를 지움 — 이 기기가 되살리지 않는다", async () => {
  const E = bandEnv(); await E.ctx.openBand("b1");
  E.db.remote("bandData/Kb1", { members: ["A", "B2"] }); await tick(5);
  eq(E.members(), ["A", "B2"], "멤버");
});
test("다른 기기가 이름을 바꿈 — 옛 이름을 되살리거나 새 이름 문서를 만들지 않는다", async () => {
  const E = bandEnv(); await E.ctx.openBand("b1");
  E.db.remote("bandData/Kb1", { members: ["A2", "B2", "C"] }); await tick(5);
  eq(E.members(), ["A2", "B2", "C"], "멤버"); eq(E.adds(), [], "member_add");
});
test("(그대로) 밴드에 새 멤버 → 시간표에 넣는다 · 시간표에만 있는 이름 → 밴드 멤버로", async () => {
  const E = bandEnv(); await E.ctx.openBand("b1");
  E.db.remote("bandData/Kb1", { members: ["A", "B2", "C", "E"] }); await tick(5);
  eq(E.adds(), ["E"], "member_add");
  E.db.put("bandData/Kb1/scheds/m1/people/pF", { name: "F", hours: {}, notes: {} });
  E.ctx.renderAll(); await tick(5);
  eq(E.members(), ["A", "B2", "C", "E", "F"], "멤버");
});
test("순서 바꾸기 — 그 사이 다른 사람이 넣은 멤버를 지우지 않는다", async () => {
  const E = bandEnv(); await E.ctx.openBand("b1");
  E.db.put("bandData/Kb1", { band: "b1", members: ["A", "B2", "C", "E"] });   // 스냅샷이 아직 안 옴
  E.ctx.moveMember(0, 1); await tick(5);
  eq(E.members(), ["B2", "A", "C", "E"], "멤버");
});
test("이름 바꾸기 — 그 사이 다른 사람이 넣은 멤버를 지우지 않는다", async () => {
  const E = bandEnv(); await E.ctx.openBand("b1");
  E.db.put("bandData/Kb1", { band: "b1", members: ["A", "B2", "C", "E"] });
  await E.ctx.renameMember("A", "A2"); await tick(5);
  eq(E.members(), ["A2", "B2", "C", "E"], "멤버");
});
test("이름 바꾸기 — 그 사이 같은 이름이 생겼으면 겹치게 만들지 않는다", async () => {
  const E = bandEnv(); await E.ctx.openBand("b1");
  E.db.put("bandData/Kb1", { band: "b1", members: ["A", "B2", "C", "A2"] });
  await E.ctx.renameMember("A", "A2"); await tick(5);
  eq(E.members(), ["A", "B2", "C", "A2"], "멤버");
});
test("밴드 열기 중 다른 밴드를 열면 — 먼저 누른 밴드가 끼어들지 않는다", async () => {
  const E = bandEnv(); E.put("b2", ["X"], ["X"]);
  let go; E.ctx.keyGate.b1 = new Promise(r => go = r);
  const p1 = E.ctx.openBand("b1");
  await E.ctx.openBand("b2");
  go(); await p1; await tick(5);
  eq(E.ctx.KEY, "Kb2", "KEY"); eq(E.ctx.B, "b2", "B");
  eq(E.db.listeners.filter(l => l.path.startsWith("bandData/Kb1")).length, 0, "b1 리스너");
});

/* ── 📌 해제 → '예정' 회차 ── */
function fixEnv(answer){
  const E = load(["onFixesChanged", "sessKey", { code: `
var asked = [];
function fmtPick(t){ return t; }
function ask(o){ asked.push(o); return Promise.resolve(${answer}); }
` }]);
  E.run(`M = "m1"; KEY = "K";`);
  return E;
}
function sess(E, id, extra){
  const s = Object.assign({ date: "2026-10-11", from: "17:00", to: "20:00", place: "", room: "", status: "예정",
                            absent: [], songs: [], memo: "", added: "2026-10-01", auto: "m1|" + T1, fixNote: "" }, extra);
  E.db.put("bandData/K/sessions/" + id, s);
  E.run(`SESS.push(Object.assign({ id: ${JSON.stringify(id)} }, ${JSON.stringify(s)}))`);
}
test("📌 해제 — 곡·메모 없는 '예정' 회차는 묻지 않고 지운다 (그대로)", async () => {
  const E = fixEnv(true); sess(E, "h1", {});
  E.ctx.onFixesChanged([T1], [], {}); await tick(5);
  eq(E.ctx.asked.length, 0, "물음"); ok(!E.db.get("bandData/K/sessions/h1"), "안 지워짐");
});
test("📌 해제 — 곡·메모를 적은 회차는 묻는다 · '아니오' 면 남긴다", async () => {
  const E = fixEnv(false); sess(E, "h1", { songs: ["s1"], memo: "m" });
  E.ctx.onFixesChanged([T1], [], {}); await tick(5);
  eq(E.ctx.asked.length, 1, "물음");
  const s = E.db.get("bandData/K/sessions/h1"); ok(s, "지워짐"); eq(s.songs, ["s1"], "곡"); eq(s.auto, "", "auto");
});
test("📌 해제 — 곡·메모를 적은 회차도 '지우기' 면 지운다", async () => {
  const E = fixEnv(true); sess(E, "h1", { memo: "m" });
  E.ctx.onFixesChanged([T1], [], {}); await tick(5);
  eq(E.ctx.asked.length, 1, "물음"); ok(!E.db.get("bandData/K/sessions/h1"), "안 지워짐");
});

/* ── 곡·회차 수정 창 ── */
const FORM = `
var songEditing = null, songStatus = "", songView = "", voteOrder = null;
var histEditing = null, hAbsent = new Map(), hSongs = new Set(), hStatus = "예정";
function closeModal(){} function renderSongs(){} function ls(){ return ""; } function todayIso(){ return "2026-11-05"; }
function newId(p){ return p + "new"; } function fmtFull(d){ return d; } function buildParts(){ return $("#parts").value; }
`;
function formEnv(){
  const E = load(["spotOf", "autoOf", "fromField", "fmtAbs", "syncSchedFix", { line: '$("#f_save").onclick' }, { line: '$("#h_save").onclick' }, { code: FORM }]);
  E.run(`KEY = "K";`);
  return E;
}
const SONG = { status: "후보", title: "T", artist: "Ar", key: "", picker: "A", links: "", parts: "", tuning: "", memo: "", added: "2026-10-01", votes: {} };
function openSong(E, s){
  E.db.put("bandData/K/songs/s1", s);
  E.run(`songEditing = Object.assign({ id: "s1" }, ${JSON.stringify(s)}); songStatus = spotOf(songEditing.status);`);
  ["title", "artist", "key", "picker", "links", "tuning", "memo"].forEach(k => E.dom.$("#f_" + k).value = s[k] || "");
  E.dom.$("#parts").value = s.parts || "";
}
const lastWrite = (E, p) => E.db.writes.filter(w => w.path === p).pop();
test("곡 수정 — 그 사이 다른 사람이 옮긴 상태(status)를 되돌리지 않는다", async () => {
  const E = formEnv(); openSong(E, SONG);
  E.db.remote("bandData/K/songs/s1", { status: "연습중" });
  E.dom.$("#f_memo").value = "새 메모";
  E.dom.$("#f_save").onclick(); await tick(5);
  const s = E.db.get("bandData/K/songs/s1"); eq(s.status, "연습중", "status"); eq(s.memo, "새 메모", "memo");
});
test("곡 수정 — 옛 문서에 없던 칸은 같이 채운다 (규칙이 모든 칸을 본다)", async () => {
  const E = formEnv(); const old = Object.assign({}, SONG); delete old.tuning; openSong(E, old);
  E.dom.$("#f_memo").value = "x";
  E.dom.$("#f_save").onclick(); await tick(5);
  eq(Object.keys(lastWrite(E, "bandData/K/songs/s1").data).sort(), ["memo", "tuning"], "보낸 칸");
});
const SESS0 = { date: "2026-10-11", from: "17:00", to: "20:00", place: "P", room: "R", status: "예정", absent: ["A(일)"], songs: ["s1"], memo: "", added: "2026-10-01", auto: "", fixNote: "" };
function openHist(E, h){
  E.db.put("bandData/K/sessions/h1", h);
  E.run(`histEditing = Object.assign({ id: "h1" }, ${JSON.stringify(h)}); hStatus = histEditing.status;
         hAbsent = new Map([["A", "일"]]); hSongs = new Set(histEditing.songs);`);
  ["date", "from", "place", "room", "memo"].forEach(k => E.dom.$("#h_" + k).value = h[k] || "");
  E.dom.$("#h_to").value = h.to;
}
test("회차 수정 — 그 사이 다른 사람이 넣은 불참·곡을 지우지 않는다", async () => {
  const E = formEnv(); openHist(E, SESS0);
  E.db.remote("bandData/K/sessions/h1", { absent: ["A(일)", "C"], songs: ["s1", "s2"] });
  E.dom.$("#h_memo").value = "메모";
  E.dom.$("#h_save").onclick(); await tick(5);
  const s = E.db.get("bandData/K/sessions/h1");
  eq(s.absent, ["A(일)", "C"], "absent"); eq(s.songs, ["s1", "s2"], "songs"); eq(s.memo, "메모", "memo");
});
test("이력 → 📌 옮기기 — 읽고 쓰는 사이 다른 사람의 확정을 지우지 않는다", async () => {
  const E = formEnv(), P = "bandData/K/scheds/m1", T2 = "2026-10-12 10:00~12:00";
  E.db.put(P, { title: "t", fixed: T1, fixNotes: { [T1]: "합주" }, picks: [] });
  let once = true;
  E.db.hooks.onRead = p => { if (p === P && once){ once = false; E.db.remote(P, { fixed: T1 + ", " + T2 }); } };
  await E.ctx.syncSchedFix({ sid: "m1", key: T1 }, "move", "2026-10-11 18:00~21:00");
  const m = E.db.get(P);
  eq(m.fixed, "2026-10-11 18:00~21:00, " + T2, "fixed"); eq(m.fixNotes, { "2026-10-11 18:00~21:00": "합주" }, "fixNotes");
});

test("index.html 본 스크립트가 문법 오류 없이 읽힌다", () => {
  const src = require("fs").readFileSync(require("path").join(__dirname, "..", "index.html"), "utf8");
  const s = src.split("<script>").pop().split("</script>")[0];
  new (require("vm").Script)(s);
});

runAll();
