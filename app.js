"use strict";
const QS = window.QUESTIONS || [];
QS.forEach(q => { q.paper = "A"; q.year = q.exam.split("-")[0]; q.num = q.exam.split("-")[1]; });
const BY = Object.fromEntries(QS.map(q => [q.id, q]));
const UNITS_L = ["一、會計基本概念","二、會計循環及會計帳簿","三、會計基本法則","四、分錄與日記簿","五、過帳與分類帳","六、試算與試算表","七、調整","八、結帳","九、財務報表","十、加值型營業稅會計實務","十一、現金及內部控制","十二、應收款項","十三、存貨","十四、證券投資","十五、長期營業用資產","十六、負債","十七、權益"];
const UNITS = { indexOf(u) { return UNITS_L.indexOf(u); } };
const exLabel = x => { const [y, n] = x.split("-"); return `${y}學年 第${n}次`; };
const isOk = (q, k) => !!k && (q.accept_all || k === q.answer);
const EXAM_MIN = 40;
const $app = document.getElementById("app"), $nav = document.getElementById("nav");
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const rich = s => {
  const L = String(s ?? "").split("\n"), out = [], isRow = l => /^\s*\|.*\|\s*$/.test(l), isSep = l => /^\s*\|[\s:\-|]+\|\s*$/.test(l), cells = l => l.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map(c => c.trim());
  let buf = []; const flush = () => { const t = buf.join("\n").replace(/^\n+|\n+$/g, ""); if (t) out.push(`<span class="pw">${esc(t)}</span>`); buf = []; };
  for (let i = 0; i < L.length; i++) {
    if (isRow(L[i]) && i + 1 < L.length && isSep(L[i + 1])) {
      flush(); const h = cells(L[i]); let j = i + 2, rows = [];
      while (j < L.length && isRow(L[j])) { rows.push(cells(L[j])); j++; }
      out.push(`<div class="tbl"><table><thead><tr>${h.map(c => `<th>${esc(c)}</th>`).join("")}</tr></thead><tbody>${rows.map(r => `<tr>${r.map(c => `<td>${esc(c)}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`);
      i = j - 1;
    } else buf.push(L[i]);
  }
  flush(); return out.join("");
};
const shuffle = a => { a = a.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.random() * (i + 1) | 0; [a[i], a[j]] = [a[j], a[i]]; } return a; };
const hash = s => { let h = 5381; for (const c of s) h = ((h << 5) + h + c.charCodeAt(0)) | 0; return String(h); };
const ls = { get(k) { try { return localStorage.getItem(k); } catch { return null; } }, set(k, v) { try { localStorage.setItem(k, v); return true; } catch { return false; } } };

/* ---------- 使用者資料 ---------- */
let S = { user: null, data: null, view: "login", sess: null, cfg: { units: [], years: [], nums: [], count: 10 }, timer: null };
const key = n => "ac:u:" + n;
const blank = pin => ({ pin: hash(pin), wrong: {}, mastered: {}, log: [], created: Date.now() });
function save() { if (S.user) { ls.set(key(S.user), JSON.stringify(S.data)); schedulePush(); } }

/* ---------- 雲端同步(Google 試算表 + Apps Script) ---------- */
const CLOUD = () => String(window.CLOUD_URL || "").trim();
const SY = { st: CLOUD() ? "idle" : "off", timer: null, pin: null, retry: null };
function syncBadge() { const el = document.getElementById("sy"); if (el) el.textContent = { off: "", idle: "☁ 已同步", busy: "☁ 同步中…", err: "⚠ 尚未同步(會自動重試)" }[SY.st] || ""; }
async function cloud(body, keepalive) {
  const ac = new AbortController(), to = setTimeout(() => ac.abort(), 25000);
  try { const r = await fetch(CLOUD(), { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(body), signal: ac.signal, keepalive: !!keepalive && JSON.stringify(body).length < 60000 }); return await r.json(); }
  finally { clearTimeout(to); }
}
const pack = (d, limit = 1500) => ({ paper: "A", wrong: d.wrong, mastered: d.mastered, log: d.log.slice(-limit) });
function merge(a, b) {
  const out = { wrong: {}, mastered: {}, log: [] };
  const ids = new Set([a, b].flatMap(x => [...Object.keys(x.wrong || {}), ...Object.keys(x.mastered || {})]));
  for (const id of ids) {
    const c = [];
    for (const x of [a, b]) { if (x.wrong?.[id]) c.push(["w", x.wrong[id].t || 0, x.wrong[id]]); if (x.mastered?.[id]) c.push(["m", x.mastered[id].t || 0, x.mastered[id]]); }
    c.sort((p, q) => q[1] - p[1]); (c[0][0] === "w" ? out.wrong : out.mastered)[id] = c[0][2];
  }
  const seen = new Set();
  for (const l of [...(a.log || []), ...(b.log || [])]) { const k = l.id + "@" + l.t; if (!seen.has(k)) { seen.add(k); out.log.push(l); } }
  out.log.sort((p, q) => p.t - q.t); if (out.log.length > 6000) out.log.splice(0, out.log.length - 6000);
  return out;
}
function schedulePush() {
  if (!CLOUD() || !S.user || !SY.pin) return;
  SY.st = "busy"; syncBadge(); clearTimeout(SY.timer); SY.timer = setTimeout(push, 2500);
}
async function push(keep) {
  clearTimeout(SY.timer);
  SY.timer = null;
  clearTimeout(SY.retry);
  if (!CLOUD() || !S.user || !SY.pin || !S.data) return;
  const user = S.user, pin = SY.pin, local = S.data;
  const body = { a: "save", name: user, pin, data: pack(local, keep ? 300 : 1500) };
  try {
    const result = await cloud(body, keep);
    if (S.user !== user || SY.pin !== pin || S.data !== local) return;
    if (!result.ok) throw new Error("sync");
    if (result.data) {
      const merged = merge(local, result.data);
      Object.assign(local, merged);
      ls.set(key(user), JSON.stringify(local));
      if (S.view === "wrong") render();
      else drawNav();
    }
    SY.st = SY.timer ? "busy" : "idle";
  } catch {
    if (S.user !== user || SY.pin !== pin || S.data !== local) return;
    SY.st = "err";
    if (!SY.timer) SY.retry = setTimeout(push, 30000);
  }
  syncBadge();
}
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden" && SY.timer) push(true); });
window.addEventListener("pagehide", () => { if (SY.timer) push(true); });
function rec(q, ok, mode) {
  const d = S.data, now = Date.now();
  d.log.push({ id: q.id, ok: ok ? 1 : 0, t: now, m: mode });
  if (d.log.length > 6000) d.log.splice(0, d.log.length - 6000);
  if (!ok) { const w = d.wrong[q.id] || { n: 0 }; w.n++; w.t = now; d.wrong[q.id] = w; delete d.mastered[q.id]; }
  else if (mode === "wrong" && d.wrong[q.id]) { d.mastered[q.id] = { t: now, n: d.wrong[q.id].n }; delete d.wrong[q.id]; }
  save();
}
const paper = () => "A";
const PQ = () => QS.filter(q => q.paper === paper());
const examsOf = () => [...new Set(PQ().map(q => q.exam))].sort((a, b) => b.localeCompare(a, "en", { numeric: true }));
const unitsOf = () => { const u = [...new Set(PQ().map(q => q.unit))]; return u.sort((a, b) => (UNITS.indexOf(a) < 0 ? 99 : UNITS.indexOf(a)) - (UNITS.indexOf(b) < 0 ? 99 : UNITS.indexOf(b))); };
const wrongList = () => Object.keys(S.data.wrong).filter(id => BY[id] && BY[id].paper === paper());

/* ---------- 版面 ---------- */
let kfb = 0;
function loadKatexFallback() {
  if (kfb++) return;
  const base = "https://cdn.jsdelivr.net/npm/katex@0.16.47/dist/";
  const l = document.createElement("link"); l.rel = "stylesheet"; l.href = base + "katex.min.css"; document.head.appendChild(l);
  const load = (src, cb) => { const sc = document.createElement("script"); sc.src = src; sc.onload = cb; sc.onerror = () => { const m = document.getElementById("kerr"); if (!m) $app.insertAdjacentHTML("afterbegin", '<div id="kerr" class="hint">⚠️ 公式元件沒有載入成功,畫面會出現亂碼。請按 Ctrl+F5 重新整理;若還是這樣,請告訴老師。</div>'); }; document.head.appendChild(sc); };
  load(base + "katex.min.js", () => load(base + "contrib/auto-render.min.js", () => mathify()));
}
let cssChecked = 0;
function ensureKatexCss() {
  if (cssChecked) return; cssChecked = 1;
  setTimeout(() => {
    const t = document.createElement("span"); t.className = "katex"; t.innerHTML = '<span class="katex-html"><span class="strut" style="height:1em"></span></span>'; document.body.appendChild(t);
    const bad = getComputedStyle(t).fontSize === getComputedStyle(document.body).fontSize && getComputedStyle(t.firstChild).display !== "inline-block" && getComputedStyle(t.firstChild.firstChild).display !== "inline-block"; t.remove();
    if (bad) { const l = document.createElement("link"); l.rel = "stylesheet"; l.href = "https://cdn.jsdelivr.net/npm/katex@0.16.47/dist/katex.min.css"; document.head.appendChild(l); }
  }, 1500);
}
function mathify() { ensureKatexCss(); if (!window.renderMathInElement || !window.katex) { loadKatexFallback(); return; } try { renderMathInElement($app, { delimiters: [{ left: "\\[", right: "\\]", display: true }, { left: "\\(", right: "\\)", display: false }], throwOnError: false, output: "html" }); } catch (e) { console.error("math render failed", e); } }
function draw(html) { $app.innerHTML = html; mathify(); window.scrollTo(0, 0); }
function drawNav() {
  if (!S.user) { $nav.innerHTML = ""; return; }
  const t = [["practice", "自由練習"], ["exams", "模擬考"], ["wrong", `錯題本(${wrongList().length})`], ["me", "我的紀錄"]];
  $nav.innerHTML = t.map(([k, n]) => `<button class="${S.view === k || (k === "practice" && S.view === "run" && S.sess?.mode === "practice") ? "on" : ""}" data-go="${k}">${n}</button>`).join("") + `<span class="who">${esc(S.user)}</span><span class="who" id="sy"></span><button data-go="logout">登出</button>`;
  syncBadge();
}
function go(v) {
  if (v === "logout") { if (S.sess?.mode === "exam" && !confirm("考試進行中,登出會放棄本次作答,確定嗎?")) return; clearInterval(S.timer); if (SY.timer) push(); try { sessionStorage.removeItem("ac:cur"); sessionStorage.removeItem("ac:pin"); } catch {} SY.pin = null; S.user = S.data = S.sess = null; S.view = "login"; }
  else { if (S.sess?.mode === "exam" && S.sess.live && v !== "run" && !confirm("離開會放棄本次模擬考,確定嗎?")) return; clearInterval(S.timer); S.view = v; if (v !== "run") S.sess = null; }
  render();
}
document.addEventListener("click", e => { const pb = e.target.closest("[data-paper]"); if (pb && S.data) { if (S.sess?.mode === "exam" && S.sess.live && !confirm("切換會放棄本次模擬考,確定嗎?")) return; clearInterval(S.timer); S.data.paper = pb.dataset.paper; S.cfg.units = []; S.sess = null; save(); S.view = S.view === "wrong" || S.view === "exams" || S.view === "me" ? S.view : "practice"; render(); return; } const b = e.target.closest("[data-go]"); if (b) go(b.dataset.go); });

function render() { drawNav(); ({ print: vPrint, login: vLogin, practice: vPractice, exams: vExams, wrong: vWrong, me: vMe, run: vRun, result: vResult }[S.view] || vLogin)(); }

/* ---------- 登入 ---------- */
function vLogin() {
  draw(`<h1>歡迎來刷題</h1><p class="lead">自己取一個代稱(不要用真實姓名),再設定 4 位數密碼。第一次使用會自動建立帳號。</p>
  <div class="card" style="max-width:420px"><label>你的名字(自己取的代稱)</label><input id="nm" type="text" maxlength="20" autocomplete="off" autocapitalize="none" autocorrect="off" spellcheck="false" value="${esc(ls.get("ac:last") || "")}">
  <p class="note" style="margin:4px 0 0">中文、英文、數字都可以。用中文輸入法時,數字鍵會變成注音符號,<b>請改按下面的數字鍵</b>(也可刪除、清除)。要輸入英文字母,按 <a href="#" id="kbt">⌨ 英文字母鍵盤</a> 用畫面上的鍵盤輸入。英文不分大小寫。</p><div id="kq" class="kb"><button type="button" data-c="1">1</button><button type="button" data-c="2">2</button><button type="button" data-c="3">3</button><button type="button" data-c="4">4</button><button type="button" data-c="5">5</button><button type="button" data-c="6">6</button><button type="button" data-c="7">7</button><button type="button" data-c="8">8</button><button type="button" data-c="9">9</button><button type="button" data-c="0">0</button><button type="button" data-c="⌫" class="del">⌫ 刪除</button><button type="button" data-c="CLR" class="del">清除</button></div><div id="kb" class="kb" hidden></div>
  <label>4 位數密碼 <span class="note">(用下面的數字鍵輸入,或直接打字)</span></label><input id="pin" type="text" inputmode="numeric" maxlength="8" autocomplete="off" placeholder="例如 1234">
  <div class="pad">${[1,2,3,4,5,6,7,8,9].map(n=>`<button type="button" data-n="${n}">${n}</button>`).join("")}<button type="button" data-n="C">清除</button><button type="button" data-n="0">0</button><button type="button" data-n="B">←</button></div>
  <p id="err" class="note" style="color:var(--bad);min-height:22px"></p><div class="hint" style="margin:0 0 12px">⚠️ 請牢記你取的「名字」和「密碼」!以後要用<b>完全相同</b>的名字和密碼登入,才看得到自己的錯題與作答紀錄。忘記了無法找回。</div><button class="btn" id="ok">進入</button>
  <p class="note" style="margin-top:14px">${CLOUD() ? "紀錄會自動存到雲端,換手機或電腦用同一個名字和密碼登入即可。" : "紀錄存在這台裝置的瀏覽器中。要換手機或電腦,請到「我的紀錄」匯出/匯入。"}</p></div>`);
  let busy = false;
  const submit = async () => {
    if (busy) return;
    const nz = v => v.normalize("NFKC").trim(), n = nz($app.querySelector("#nm").value).toLowerCase(), p = nz($app.querySelector("#pin").value), err = $app.querySelector("#err"), btn = $app.querySelector("#ok");
    if (!n) return err.textContent = "請輸入你的名字";
    if (!/^\d{4}$/.test(p)) return err.textContent = "密碼要是 4 位數字";
    let d = null; const raw = ls.get(key(n));
    if (raw) { try { d = JSON.parse(raw); } catch { d = null; } }
    const localOk = d && d.pin === hash(p);
    let offline = false, note = "";
    if (CLOUD()) {
      busy = true; btn.disabled = true; btn.textContent = "登入中…"; err.textContent = "";
      try {
        const r = await cloud({ a: "login", name: n, pin: p, data: localOk ? pack(d, 6000) : null });
        if (!r.ok) {
          busy = false; btn.disabled = false; btn.textContent = "進入";
          return err.textContent = r.err === "pin" ? "密碼不對" : r.err === "locked" ? `密碼輸錯太多次,請 ${r.wait || 10} 分鐘後再試` : r.err === "busy" ? "伺服器忙碌中,請稍後再按一次" : "雲端回應異常,請稍後再試";
        }
        if (!d) d = blank(p);
        d.pin = hash(p);
        if (r.data) { const m = merge(d, r.data); d.wrong = m.wrong; d.mastered = m.mastered; d.log = m.log; if (!localOk && r.data.paper) d.paper = r.data.paper; }
        SY.st = "idle";
      } catch {
        busy = false; btn.disabled = false; btn.textContent = "進入";
        if (!localOk) return err.textContent = "連不上雲端,請檢查網路後再試";
        offline = true; SY.st = "err";
      }
      busy = false;
    } else {
      if (d && !localOk) return err.textContent = "密碼不對";
      if (!d) d = blank(p);
    }
    S.user = n; S.data = d; SY.pin = CLOUD() ? p : null; ls.set(key(n), JSON.stringify(d)); ls.set("ac:last", n);
    try { sessionStorage.setItem("ac:cur", n); if (CLOUD()) sessionStorage.setItem("ac:pin", p); } catch {}
    S.view = "practice"; render();
    if (CLOUD()) { SY.st = "busy"; push(); }
  };
  const nmEl = $app.querySelector("#nm"), kb = $app.querySelector("#kb");
  kb.innerHTML = "1234567890qwertyuiopasdfghjklzxcvbnm".split("").map(c => `<button type="button" data-c="${c}">${c}</button>`).join("") + '<button type="button" data-c="_">_</button><button type="button" data-c="-">-</button><button type="button" data-c="⌫" class="wide">⌫ 刪除</button>';
  const kq = $app.querySelector("#kq");
  kb.onclick = kq.onclick = e => { const b = e.target.closest("[data-c]"); if (!b) return; e.preventDefault(); const c = b.dataset.c, ch = Array.from(nmEl.value); nmEl.blur(); nmEl.value = c === "CLR" ? "" : c === "⌫" ? ch.slice(0, -1).join("") : ch.concat(c).slice(0, 20).join(""); };
  $app.querySelector("#kbt").onclick = e => { e.preventDefault(); kb.hidden = !kb.hidden; };
  $app.querySelector("#ok").onclick = submit;
  const pin = $app.querySelector("#pin");
  $app.querySelectorAll(".pad [data-n]").forEach(b => b.onclick = () => { const k = b.dataset.n; pin.value = k === "C" ? "" : k === "B" ? pin.value.slice(0, -1) : (pin.value + k).slice(0, 4); });
  $app.querySelectorAll("input").forEach(i => i.onkeydown = e => e.key === "Enter" && submit());
}

/* ---------- 自由練習 ---------- */
function pool() { const c = S.cfg; return PQ().filter(q => (!c.units.length || c.units.includes(q.unit)) && (!c.years.length || c.years.includes(q.year)) && (!c.nums.length || c.nums.includes(q.num))); }
function vPractice() {
  const c = S.cfg, uAvail = unitsOf(), yAvail = [...new Set(PQ().map(q => q.year))].sort().reverse(), nAvail = [...new Set(PQ().map(q => q.num))].sort();
  draw(`<h1>自由練習 <span class="tag">會計學</span></h1><p class="lead">挑單元與試卷,答完馬上看詳解。答錯的題目會自動進錯題本。</p>
  <div class="card"><h2>單元(不選 = 全部)</h2><div class="row" id="us">${uAvail.map(u => `<button class="chip ${c.units.includes(u) ? "on" : ""}" data-u="${esc(u)}">${esc(u)}</button>`).join("")}</div></div>
  <div class="card"><h2>學年度(不選 = 全部)</h2><div class="row">${yAvail.map(y => `<button class="chip ${c.years.includes(y) ? "on" : ""}" data-y="${y}">${y}學年</button>`).join("")}</div>
  <h2 style="margin-top:14px">第幾次模擬考(不選 = 全部)</h2><div class="row">${nAvail.map(n => `<button class="chip ${c.nums.includes(n) ? "on" : ""}" data-nn="${n}">第${n}次</button>`).join("")}</div></div>
  <div class="card"><h2>題數</h2><div class="row" id="cs">${[10, 20, 0].map(n => `<button class="chip ${c.count === n ? "on" : ""}" data-c="${n}">${n || "全部"}${n ? " 題" : ""}</button>`).join("")}</div>
  <p style="margin:16px 0 4px" class="note">符合條件</p><div class="count">${pool().length}<small style="font-size:16px;color:var(--sub)"> 題</small></div>
  <p><button class="btn" id="go" ${pool().length ? "" : "disabled"}>開始練習 →</button></p></div>`);
  const tog = (arr, v) => { const i = arr.indexOf(v); i < 0 ? arr.push(v) : arr.splice(i, 1); vPractice(); };
  $app.querySelectorAll("[data-u]").forEach(b => b.onclick = () => tog(c.units, b.dataset.u));
  $app.querySelectorAll("[data-y]").forEach(b => b.onclick = () => tog(c.years, b.dataset.y));
  $app.querySelectorAll("[data-nn]").forEach(b => b.onclick = () => tog(c.nums, b.dataset.nn));
  $app.querySelectorAll("[data-c]").forEach(b => b.onclick = () => { c.count = +b.dataset.c; vPractice(); });
  $app.querySelector("#go").onclick = () => { let l = shuffle(pool()); if (c.count) l = l.slice(0, c.count); start("practice", l); };
}

/* ---------- 模擬考 ---------- */
function vExams() {
  draw(`<h1>模擬考</h1><p class="lead">會計學(專二 第 1~25 題)。整份題目、限時 ${EXAM_MIN} 分鐘,交卷後才顯示答案與得分。</p>
  ${[...new Set(examsOf().map(x => x.split("-")[0]))].map(y => `<h2 style="margin:18px 0 8px">${y} 學年度</h2><div class="grid">${examsOf().filter(x => x.startsWith(y + "-")).sort().map(x => { const done = S.data.log.some(l => l.m === "exam" && l.id.startsWith(x + "-" + paper() + "-")); return `<div class="tile"><b>第${x.split("-")[1]}次</b><small>${PQ().filter(q => q.exam === x).length} 題${done ? " · 做過" : ""}</small><p style="margin:10px 0 0"><button class="btn sm" data-x="${x}">開始</button></p></div>`; }).join("")}</div>`).join("") || '<p class="empty">這個卷別還沒有題庫</p>'}`);
  $app.querySelectorAll("[data-x]").forEach(b => b.onclick = () => start("exam", PQ().filter(q => q.exam === b.dataset.x).sort((a, b) => a.q - b.q)));
}

/* ---------- 錯題本 ---------- */
S.sel = new Set();
function vWrong() {
  const l = wrongList().map(id => BY[id]).sort((a, b) => S.data.wrong[b.id].t - S.data.wrong[a.id].t), m = Object.keys(S.data.mastered).length;
  S.sel = new Set([...S.sel].filter(id => l.some(q => q.id === id)));
  draw(`<h1>錯題本</h1><p class="lead">在這裡答對的題目,會從錯題本移除(記入「已克服」)。答錯就繼續留著。</p>
  <div class="card"><div class="count">${l.length}<small style="font-size:16px;color:var(--sub)"> 題待複習</small></div><p class="note">已克服 ${m} 題</p>
  <button class="btn" id="go" ${l.length ? "" : "disabled"}>刷錯題 →</button></div>
  ${l.length ? `<div class="card"><h2>清單 <span class="note" id="selc"></span></h2>
  <div class="row" style="margin-bottom:10px"><button class="chip" id="sa">全選</button><button class="chip" id="s2">只選一錯再錯(錯 2 次以上)</button><button class="chip" id="sn">全部取消</button></div>
  ${l.map(q => `<label class="wr"><input type="checkbox" data-s="${esc(q.id)}" ${S.sel.has(q.id) ? "checked" : ""}><span><span class="tag">${exLabel(q.exam)} 第${q.q}題</span> ${esc(q.unit)}・${esc(q.topic)} <span class="note">錯 ${S.data.wrong[q.id].n} 次</span></span></label>`).join("")}
  <div class="pdfbar"><label class="inl"><input type="checkbox" id="wsp" checked> 每題留空白計算區</label>
  <button class="btn" id="pdf">產生 A4 練習卷</button></div></div>` : '<p class="empty">目前沒有錯題,太強了!</p>'}`);
  const b = $app.querySelector("#go"); if (b) b.onclick = () => start("wrong", shuffle(l));
  if (!l.length) return;
  const upd = () => { $app.querySelector("#selc").textContent = `(已勾選 ${S.sel.size} 題)`; $app.querySelectorAll("[data-s]").forEach(c => c.checked = S.sel.has(c.dataset.s)); };
  $app.querySelectorAll("[data-s]").forEach(c => c.onchange = () => { c.checked ? S.sel.add(c.dataset.s) : S.sel.delete(c.dataset.s); upd(); });
  $app.querySelector("#sa").onclick = () => { S.sel = new Set(l.map(q => q.id)); upd(); };
  $app.querySelector("#s2").onclick = () => { S.sel = new Set(l.filter(q => S.data.wrong[q.id].n >= 2).map(q => q.id)); upd(); if (!S.sel.size) alert("目前沒有錯 2 次以上的題目"); };
  $app.querySelector("#sn").onclick = () => { S.sel = new Set(); upd(); };
  $app.querySelector("#pdf").onclick = () => {
    const ids = l.filter(q => S.sel.has(q.id)).map(q => q.id);
    if (!ids.length) return alert("請先勾選要練習的題目");
    S.pr = { ids, ans: true, sp: $app.querySelector("#wsp").checked }; S.view = "print"; render();
  };
  upd();
}
function vPrint() {
  const { ids, ans, sp } = S.pr, qs = ids.map(id => BY[id]).filter(Boolean), d = new Date();
  const ds = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`, ks = ["A", "B", "C", "D"];
  const opt = (q, k, i) => { const o = q.options[i]; return `<div class="po"><b>(${k})</b> ${/^IMG:/.test(o) ? `<img src="${esc(o.slice(4))}" alt="">` : `${rich(o)}`}</div>`; };
  const anyImg = q => q.options.some(o => /^IMG:/.test(o));
  draw(`<div class="noprint ptool"><button class="btn ghost" data-go="wrong">← 回錯題本</button><button class="btn" id="dopdf">下載 A4 PDF</button> <button class="btn ghost" id="doprint">列印</button>
  <p id="pdf-status" class="note" role="status">練習題在前，答案與詳解在最後。</p></div>
  <div class="pv"><div class="phead"><h1>會計學 錯題練習卷</h1><div>姓名/代稱:${esc(S.user)}　日期:${ds}　共 ${qs.length} 題　　得分:______</div></div>
  ${qs.map((q, i) => `<section class="pq"><div class="pt"><b>${i + 1}.</b> <span class="src">(${exLabel(q.exam)} 第${q.q}題)</span></div>
  <div class="ps">${rich(q.stem)}</div>${q.figure ? `<img class="pfig" src="${esc(q.figure)}" alt="">` : ""}
  <div class="pos ${anyImg(q) ? "img" : ""}">${ks.map((k, j) => opt(q, k, j)).join("")}</div>${sp ? '<div class="pwork">計算區</div>' : ""}</section>`).join("")}
  ${ans ? `<div class="pbreak"></div><div class="phead"><h1>解答與詳解</h1></div>
  <p class="pans"><b>答案速查:</b>${qs.map((q, i) => `${i + 1}. ${q.accept_all ? "送分" : q.answer}`).join("　")}</p>
  ${qs.map((q, i) => { const e = q.explain || {}; return `<section class="pe"><div class="pt"><b>${i + 1}.</b> 答案 ${q.accept_all ? "(送分)" : q.answer}　<span class="src">${esc(q.unit)}・${esc(q.topic)}</span></div>
  ${e.concept ? `<div><b>考點:</b>${rich(e.concept)}</div>` : ""}${(e.steps || []).length ? `<ol>${e.steps.map(x => `<li>${rich(x)}</li>`).join("")}</ol>` : ""}${e.pitfalls ? `<div><b>常見陷阱</b> ${rich(e.pitfalls)}</div>` : ""}${Object.keys(e.distractors || {}).length ? `<h3>選項比對</h3><ul>${Object.entries(e.distractors).map(([letter, reason]) => `<li><b>${esc(letter)}.</b> ${rich(reason)}</li>`).join("")}</ul>` : ""}</section>`; }).join("")}` : ""}</div>`);
  $app.querySelector("#doprint").onclick = () => window.print();
  const download = $app.querySelector("#dopdf");
  const status = $app.querySelector("#pdf-status");
  download.onclick = async () => {
    download.disabled = true;
    status.textContent = "正在製作 PDF…";
    try {
      await window.saveWorksheet($app.querySelector(".pv"), page => {
        status.textContent = `正在製作第 ${page} 頁…`;
      });
      status.textContent = "PDF 已產生，請查看下載項目。";
    } catch {
      status.textContent = "PDF 下載失敗。可按「列印」，選擇儲存為 PDF。";
    } finally {
      download.disabled = false;
    }
  };
}
/* ---------- 作答流程 ---------- */
function start(mode, qs) {
  clearInterval(S.timer);
  S.sess = { mode, qs, i: 0, ans: {}, live: true, end: mode === "exam" ? Date.now() + EXAM_MIN * 60000 : 0, ok: 0 };
  S.view = "run"; render();
  if (mode === "exam") S.timer = setInterval(() => { const t = $app.querySelector("#clock"); const left = S.sess.end - Date.now(); if (left <= 0) { clearInterval(S.timer); submitExam(true); } else if (t) t.textContent = fmt(left); }, 1000);
}
const fmt = ms => { const s = Math.max(0, ms / 1000 | 0); return String((s / 60) | 0).padStart(2, "0") + ":" + String(s % 60).padStart(2, "0"); };

function optHtml(q, k, i, st) {
  const o = q.options[i], body = /^IMG:/.test(o) ? `<img src="${esc(o.slice(4))}" alt="選項${k}">` : `${rich(o)}`;
  return `<button class="opt ${st}" data-k="${k}" ${["right", "wrong", "lock"].includes(st) ? "disabled" : ""}><span class="k">${k}</span>${body}</button>`;
}
function qBlock(q, picked, reveal, noFeedbackSel) {
  const ks = ["A", "B", "C", "D"];
  const opts = ks.map((k, i) => { let st = ""; if (reveal) { st = (k === q.answer || (q.accept_all && k === picked)) ? "right" : (k === picked ? "wrong" : "lock"); } else if (k === picked) st = "sel"; return optHtml(q, k, i, st).replace('class="opt r"', 'class="opt right"'); }).join("");
  return `<div class="card"><div class="qhead"><span class="tag">${exLabel(q.exam)} 第${q.q}題</span><span>${esc(q.unit)}・${esc(q.topic)}</span>${q.verified === false ? '<span class="tag warn">解析待確認</span>' : ""}</div>
  <div class="stem">${rich(q.stem)}</div>${q.figure ? `<img class="fig" src="${esc(q.figure)}" alt="題目附圖">` : ""}<div id="opts">${opts}</div>${reveal ? expHtml(q, picked) : ""}</div>`;
}
function expHtml(q, picked) {
  const ok = isOk(q, picked), e = q.explain || {}, dis = e.distractors || {};
  return `<div class="verdict ${ok ? "ok" : "bad"}">${ok ? "答對了!" : `答錯了。正確答案是 ${q.answer}`}${q.accept_all ? "(本題官方一律給分)" : ""}</div>
  <div class="exp">${!ok && picked && dis[picked] ? `<div class="hint">你選 ${picked}:${rich(dis[picked])}</div>` : ""}
  ${e.concept ? `<h3>這題在考什麼</h3><div>${rich(e.concept)}</div>` : ""}
  ${(e.steps || []).length ? `<h3>解題步驟</h3><ol>${e.steps.map(s => `<li>${rich(s)}</li>`).join("")}</ol>` : ""}
  ${e.pitfalls ? `<h3>常見陷阱</h3><div>${rich(e.pitfalls)}</div>` : ""}
  ${Object.keys(dis).filter(k => k !== picked).length ? `<details><summary>其他選項為什麼不對</summary>${Object.keys(dis).filter(k => k !== picked).map(k => `<p style="margin:6px 0"><b>${k}</b>:${rich(dis[k])}</p>`).join("")}</details>` : ""}
  ${q.official ? `<details><summary>官方詳解(簡版)</summary><div style="margin-top:6px">${rich(q.official)}</div></details>` : ""}
  ${q.verified === false ? `<div class="hint">這題的解析還在請老師確認,若和你算的不同,請以老師為準。</div>` : ""}</div>`;
}

function vRun() {
  const s = S.sess; if (!s) return vPractice();
  const q = s.qs[s.i], n = s.qs.length, picked = s.ans[q.id], exam = s.mode === "exam";
  const reveal = s.mode === "review" || (!exam && picked);
  const pal = exam || s.mode === "review" ? `<div class="palette">${s.qs.map((x, i) => { const a = s.ans[x.id]; const c = s.mode === "review" ? (isOk(x, a) ? "r" : "w") : (a ? "done" : ""); return `<button class="${c} ${i === s.i ? "cur" : ""}" data-p="${i}">${i + 1}</button>`; }).join("")}</div>` : "";
  const title = { practice: "自由練習", wrong: "刷錯題", exam: "模擬考", review: "檢討" }[s.mode];
  const last = s.i === n - 1;
  draw(`<div class="sticky"><div><b>${title}</b> <span class="note">${s.i + 1} / ${n}</span></div>${exam ? `<div>剩餘 <b id="clock">${fmt(s.end - Date.now())}</b></div>` : s.mode === "review" ? `<button class="btn ghost sm" data-go="exams">回模擬考</button>` : ""}</div>
  <div class="bar"><i style="width:${(s.i + (picked ? 1 : 0)) / n * 100}%"></i></div>${pal}
  ${qBlock(q, picked, reveal)}
  <div class="row" style="justify-content:space-between"><button class="btn ghost" id="prev" ${s.i ? "" : "disabled"}>← 上一題</button>
  ${exam && last ? `<button class="btn" id="sub">交卷</button>` : last && s.mode !== "exam" ? `<button class="btn" id="fin" ${!picked && s.mode !== "review" ? "disabled" : ""}>${s.mode === "review" ? "結束檢討" : "看結果"}</button>` : `<button class="btn" id="next" ${!picked && !exam && s.mode !== "review" ? "disabled" : ""}>下一題 →</button>`}</div>`);
  $app.querySelectorAll(".opt").forEach(b => b.onclick = () => {
    if (reveal) return; const k = b.dataset.k;
    if (exam) { s.ans[q.id] = s.ans[q.id] === k ? undefined : k; vRun(); }
    else { s.ans[q.id] = k; const ok = isOk(q, k); if (ok) s.ok++; rec(q, ok, s.mode); drawNav(); vRun(); }
  });
  $app.querySelectorAll("[data-p]").forEach(b => b.onclick = () => { s.i = +b.dataset.p; vRun(); });
  const on = (id, f) => { const b = $app.querySelector(id); if (b) b.onclick = f; };
  on("#prev", () => { s.i--; vRun(); }); on("#next", () => { s.i++; vRun(); });
  on("#fin", () => { if (s.mode === "review") { S.view = "exams"; S.sess = null; render(); } else { S.view = "result"; render(); } });
  on("#sub", () => { const un = s.qs.filter(x => !s.ans[x.id]).length; if (!un || confirm(`還有 ${un} 題沒作答,確定交卷嗎?`)) submitExam(false); });
}
function submitExam(timeUp) {
  const s = S.sess; if (!s || s.mode !== "exam") return; clearInterval(S.timer);
  s.ok = 0; s.qs.forEach(q => { const ok = isOk(q, s.ans[q.id]); if (ok) s.ok++; rec(q, ok, "exam"); });
  s.live = false; s.timeUp = timeUp; S.view = "result"; render();
}
function vResult() {
  const s = S.sess; if (!s) return vPractice();
  const n = s.qs.length, ok = s.ok, wrongQs = s.qs.filter(q => !isOk(q, s.ans[q.id])), exam = s.mode === "exam";
  const by = {}; s.qs.forEach(q => { (by[q.unit] ||= [0, 0])[1]++; if (isOk(q, s.ans[q.id])) by[q.unit][0]++; });
  draw(`<h1>${exam ? "交卷完成" : "練習完成"}</h1>${s.timeUp ? '<p class="lead">時間到,系統已自動交卷。</p>' : ""}
  <div class="card"><div class="count">${ok}<small style="font-size:18px;color:var(--sub)"> / ${n} 題答對</small></div>${exam ? `<p class="note">以每題 4 分估算約 ${ok * 4} 分(僅供參考)</p>` : ""}<div class="bar" style="margin-top:12px"><i style="width:${ok / n * 100}%"></i></div></div>
  <div class="card"><h2>各單元</h2>${Object.entries(by).map(([u, [a, b]]) => `<p style="margin:6px 0">${esc(u)}:${a} / ${b}</p>`).join("")}</div>
  <div class="row">${exam ? `<button class="btn" id="rv">逐題檢討</button>` : ""}<button class="btn ghost" data-go="wrong">到錯題本(${wrongList().length})</button><button class="btn ghost" data-go="practice">再練一輪</button></div>`);
  const r = $app.querySelector("#rv"); if (r) r.onclick = () => { s.mode = "review"; s.i = 0; S.view = "run"; render(); };
}

/* ---------- 我的紀錄 ---------- */
function vMe() {
  const d = S.data, per = {}; PQ().forEach(q => { (per[q.unit] ||= { t: 0, w: 0 }).t++; });
  const last = {}; d.log.forEach(l => { last[l.id] = l.ok; });
  Object.keys(last).forEach(id => { const q = BY[id]; if (!q) return; const p = per[q.unit]; if (!p) return; p.a = (p.a || 0) + 1; if (last[id]) p.c = (p.c || 0) + 1; });
  draw(`<h1>我的紀錄</h1><p class="lead">${esc(S.user)} · 共作答 ${d.log.length} 次 · 已克服 ${Object.keys(d.mastered).length} 題</p>
  <div class="card"><h2>各單元表現(以每題最近一次作答計)</h2>${Object.entries(per).sort((a, b) => UNITS.indexOf(a[0]) - UNITS.indexOf(b[0])).map(([u, p]) => `<p style="margin:10px 0 2px">${esc(u)} <span class="note">已做 ${p.a || 0}/${p.t},最近答對 ${p.c || 0}</span></p><div class="bar"><i style="width:${p.a ? (p.c || 0) / p.a * 100 : 0}%"></i></div>`).join("")}</div>
  <div class="card"><h2>備份 / 換裝置</h2><p class="note">${CLOUD() ? "你的紀錄已自動同步到雲端,換裝置只要用同一個名字和密碼登入就會出現。這裡的匯出/匯入只是額外備份,或用來把舊紀錄併進來。" : "紀錄只存在這台裝置。換手機或電腦前,先匯出檔案,再到新裝置用同一個名字登入前匯入。"}</p>
  <div class="row"><button class="btn ghost" id="ex">匯出我的紀錄</button><button class="btn ghost" id="im">匯入紀錄</button><input id="file" type="file" accept=".json" hidden></div></div>`);
  $app.querySelector("#ex").onclick = () => { const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([JSON.stringify({ user: S.user, data: d })], { type: "application/json" })); a.download = `acc-${S.user}.json`; a.click(); };
  $app.querySelector("#im").onclick = () => $app.querySelector("#file").click();
  $app.querySelector("#file").onchange = async e => { try { const j = JSON.parse(await e.target.files[0].text()); if (!j.data?.log) throw 0; if (!confirm("匯入會把檔案裡的紀錄合併進目前這個帳號,確定嗎?")) return; const m = merge(S.data, j.data); S.data.wrong = m.wrong; S.data.mastered = m.mastered; S.data.log = m.log; save(); alert("匯入完成"); render(); } catch { alert("檔案格式不對"); } };
}

/* ---------- 啟動 ---------- */
(function init() {
  let cur = null; try { cur = sessionStorage.getItem("ac:cur"); } catch {}
  let pin = null;
  try { pin = sessionStorage.getItem("ac:pin"); } catch {}
  if (cur && (!CLOUD() || pin)) {
    try {
      const local = JSON.parse(ls.get(key(cur)));
      if (local) { S.user = cur; S.data = local; S.view = "practice"; SY.pin = pin; }
    } catch {}
  }
  render();
})();
