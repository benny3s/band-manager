/* 시험 도구 — index.html 에서 함수를 뽑아 node vm 에 올리고, Firestore 는 작은 가짜로 바꾼다 */
const fs = require("fs"), path = require("path"), vm = require("vm");
const SRC = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
const LINES = SRC.split(/\r?\n/);

/** 맨 앞 칸에서 시작하는 덩어리 — 시작 줄부터 맨 앞 칸의 '}' 또는 '};' 줄까지 (한 줄짜리는 그 줄만) */
function grab(start){
  const i = LINES.findIndex(l => l.startsWith(start));
  if (i < 0) return null;
  if (!/\{\s*$/.test(LINES[i])) return LINES[i];
  let j = i + 1;
  while (j < LINES.length && !/^\};?\s*$/.test(LINES[j])) j++;
  return LINES.slice(i, j + 1).join("\n");
}
function fn(name){ return grab("function " + name + "(") || grab("async function " + name + "("); }

/* ── 가짜 Firestore ── 쓰면 그 자리 리스너가 바로 불린다(로컬 반영처럼) */
function clone(v){ return v === undefined ? v : JSON.parse(JSON.stringify(v)); }
function makeDb(){
  const docs = new Map();                 // path → { data, ver }
  const listeners = [];
  const hooks = { onRead: null };
  const writes = [];
  const parent = p => p.split("/").slice(0, -1).join("/");
  const snapOf = p => { const d = docs.get(p); return { id: p.split("/").pop(), exists: !!d, data: () => d ? clone(d.data) : undefined, ref: docRef(p) }; };
  const colSnap = c => ({ docs: [...docs.keys()].filter(p => parent(p) === c).map(snapOf) });
  function fire(p){
    listeners.slice().forEach(l => {
      if (l.kind === "doc" && l.path === p) l.cb(snapOf(p));
      if (l.kind === "col" && parent(p) === l.path) l.cb(colSnap(l.path));
    });
  }
  function apply(p, d, merge){
    const cur = docs.get(p);
    if (merge && !cur) throw Object.assign(new Error("no doc " + p), { code: "not-found" });
    const out = merge ? clone(cur.data) : {};
    for (const k in d){
      const v = d[k];
      if (v && v.__op === "union"){ const a = (out[k] || []).slice(); v.v.forEach(x => { if (!a.includes(x)) a.push(x); }); out[k] = a; }
      else if (v && v.__op === "remove") out[k] = (out[k] || []).filter(x => !v.v.includes(x));
      else out[k] = clone(v);
    }
    docs.set(p, { data: out, ver: (cur ? cur.ver : 0) + 1 });
    writes.push({ path: p, data: clone(d), merge });
    fire(p);
  }
  function del(p){ docs.delete(p); writes.push({ path: p, del: true }); fire(p); }
  function read(p){ const s = snapOf(p); if (hooks.onRead) hooks.onRead(p); return s; }
  function docRef(p){
    return { path: p, id: p.split("/").pop(),
      collection: n => colRef(p + "/" + n),
      get: () => Promise.resolve(read(p)),
      set: d => { apply(p, d, false); return Promise.resolve(); },
      update: d => { try { apply(p, d, true); return Promise.resolve(); } catch(e){ return Promise.reject(e); } },
      delete: () => { del(p); return Promise.resolve(); },
      onSnapshot: cb => { const l = { kind: "doc", path: p, cb }; listeners.push(l); cb(snapOf(p)); return () => listeners.splice(listeners.indexOf(l), 1); } };
  }
  let auto = 0;
  function colRef(c){
    return { path: c,
      doc: id => docRef(c + "/" + id),
      add: d => { const r = docRef(c + "/auto" + (++auto)); r.set(d); return Promise.resolve(r); },
      get: () => Promise.resolve(colSnap(c)),
      where: (f, op, v) => ({ get: () => Promise.resolve({ docs: colSnap(c).docs.filter(s => s.data()[f] === v) }) }),
      onSnapshot: cb => { const l = { kind: "col", path: c, cb }; listeners.push(l); cb(colSnap(c)); return () => listeners.splice(listeners.indexOf(l), 1); } };
  }
  /* 트랜잭션 — 읽은 문서가 그 사이 바뀌었으면 다시 돈다 (서버와 같은 모양) */
  async function runTransaction(f){
    for (let n = 0; n < 5; n++){
      const seen = new Map(), out = [];
      const t = {
        get: ref => { const d = docs.get(ref.path); seen.set(ref.path, d ? d.ver : 0); return Promise.resolve(read(ref.path)); },
        update: (ref, d) => { out.push(() => apply(ref.path, d, true)); return t; },
        set: (ref, d) => { out.push(() => apply(ref.path, d, false)); return t; },
        delete: ref => { out.push(() => del(ref.path)); return t; }
      };
      const r = await f(t);
      if ([...seen].some(([p, v]) => ((docs.get(p) || {}).ver || 0) !== v)) continue;
      out.forEach(w => w());
      db.txCount++;
      return r;
    }
    throw new Error("transaction failed");
  }
  const db = { docs, writes, hooks, listeners, txCount: 0, collection: n => colRef(n), runTransaction,
    put: (p, d) => docs.set(p, { data: clone(d), ver: 1 }),
    remote: (p, d) => apply(p, d, true),          // 다른 기기가 쓴 것처럼
    get: p => { const d = docs.get(p); return d ? clone(d.data) : undefined; },
    list: c => [...docs.keys()].filter(p => parent(p) === c).map(p => clone(docs.get(p).data)) };
  return db;
}
const FV = { arrayUnion: (...v) => ({ __op: "union", v }), arrayRemove: (...v) => ({ __op: "remove", v }), serverTimestamp: () => "TS" };

/** $("#id") 가짜 — 값·hidden 을 담는 빈 상자 */
function fakeDom(){
  const box = {};
  const $ = sel => box[sel] || (box[sel] = { hidden: true, value: "", textContent: "", disabled: false,
                                             classList: { toggle(){}, add(){} }, setAttribute(){}, appendChild(){}, innerHTML: "" });
  return { $, box };
}

/** 뽑은 코드 + 가짜들로 vm 하나. parts: "함수이름" 또는 { line: "맨 앞 글자" } 또는 { code } */
function load(parts, extra){
  const db = makeDb(), dom = fakeDom(), status = [];
  const ctx = Object.assign({ console, setTimeout, clearTimeout, URLSearchParams, db, FV, $: dom.$, status,
    location: { search: "", pathname: "/" } }, extra || {});
  vm.createContext(ctx);
  const pre = `var B = "", KEY = "", BD = null, SCHEDS = [], SONGS = [], SESS = [], M = "", ST = { meet: null }, me = "", tab = "sched", bandUnsub = [], dirty = false;
function setStatus(t, k){ status.push([k, t]); }
function errText(e){ return (e && e.message) || "오류"; }
function bmMembers(){ return (BD && BD.members) || []; }
const dataRef = () => db.collection("bandData").doc(KEY);
function SCHED_COL(){ return db.collection("bandData").doc(KEY).collection("scheds"); }
const songRef = id => dataRef().collection("songs").doc(id);
const sessRef = id => dataRef().collection("sessions").doc(id);
`;
  const code = parts.map(p => {
    const c = typeof p === "string" ? fn(p) : p.line ? grab(p.line) : p.opt ? (fn(p.opt) || "") : p.optLine ? (grab(p.optLine) || "") : p.code;
    if (c == null) throw new Error("index.html 에서 못 찾음: " + JSON.stringify(p));
    return c;
  }).join("\n");
  vm.runInContext(pre + code + "\n;globalThis.__get = n => eval(n);", ctx, { filename: "index-extract.js" });
  return { ctx, db, dom, status, get: n => ctx.__get(n), run: c => vm.runInContext(c, ctx) };
}

/* 아주 작은 시험 틀 */
const T = [];
function test(name, f){ T.push({ name, f }); }
async function runAll(){
  let ok = 0, bad = 0;
  for (const t of T){
    try { await t.f(); ok++; console.log("  ok   " + t.name); }
    catch(e){ bad++; console.log("  FAIL " + t.name + " — " + (e && e.message)); }
  }
  console.log(path.basename(process.argv[1]) + ": " + ok + " 통과, " + bad + " 실패");
  process.exitCode = bad ? 1 : 0;
}
const tick = ms => new Promise(r => setTimeout(r, ms || 0));
function eq(a, b, msg){ const x = JSON.stringify(a), y = JSON.stringify(b); if (x !== y) throw new Error((msg || "다름") + ": " + x + " ≠ " + y); }
function ok(v, msg){ if (!v) throw new Error(msg || "거짓"); }
module.exports = { grab, fn, load, test, runAll, tick, eq, ok, clone };
