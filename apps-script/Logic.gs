/**
 * 일정 판단 로직 (서버와 휴대폰 화면이 같은 코드를 씀)
 *
 * makeLogic 안에는 바깥 변수를 쓰면 안 돼요.
 * 이 함수의 소스코드가 그대로 화면(Index.html)에도 들어가기 때문이에요.
 * 날짜는 모두 'YYYY-MM-DD' 문자열, 시간은 자정부터 센 '분'(예: 09:30 → 570)으로 다뤄요.
 */
function makeLogic(cfg) {
  "use strict";
  var OPEN = cfg.open, CLOSE = cfg.close, BUF = cfg.buffer;
  var FULL_START = cfg.fullStart, FULL_DUR = cfg.fullDur;
  var UNIT = cfg.unit, CLOSED = cfg.closedDays || [0];
  var DOW = "일월화수목금토";

  // ---------- 날짜/시간 도우미 ----------
  function pad(n) { return String(n).padStart(2, "0"); }
  function key(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function fromKey(k) { var p = k.split("-").map(Number); return new Date(p[0], p[1] - 1, p[2]); }
  function addDays(d, n) { var x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
  function day0(now) { return new Date(now.getFullYear(), now.getMonth(), now.getDate()); }
  function minsOf(now) { return now.getHours() * 60 + now.getMinutes(); }
  function mondayOf(d) { return addDays(d, -((d.getDay() + 6) % 7)); }
  function isClosed(k) { return CLOSED.indexOf(fromKey(k).getDay()) >= 0; }
  function fmtDate(k) { var d = fromKey(k); return (d.getMonth() + 1) + "월 " + d.getDate() + "일(" + DOW[d.getDay()] + ")"; }
  function fmtShort(k) { var d = fromKey(k); return (d.getMonth() + 1) + "/" + d.getDate() + "(" + DOW[d.getDay()] + ")"; }
  function fmtTime(min) { var h = Math.floor(min / 60), m = min % 60; return (h < 12 ? "오전 " : "오후 ") + (h % 12 || 12) + "시" + (m ? " " + m + "분" : ""); }
  function hm(min) { return pad(Math.floor(min / 60)) + ":" + pad(min % 60); }
  function fmtDur(d) { var h = Math.floor(d / 60), m = d % 60; return (h ? h + "시간" : "") + (m ? (h ? " " : "") + m + "분" : ""); }
  function whenText(date, start, full) { return fmtDate(date) + " " + (full ? "하루 종일" : fmtTime(start)); }

  // ---------- 작업 이름/소요시간 ----------
  function labelFor(kind, items, pyeong) {
    var py = pyeong ? " · " + pyeong + "평" : "";
    if (kind === "move_in") return "입주청소" + py;
    if (kind === "move_out") return "이사청소" + py;
    if (kind === "home") return "거주청소" + py;
    if (kind === "washer") return "세탁기 청소";
    if (kind === "aircon" && items.length) return "에어컨 청소 · " + items.join(", ");
    return "청소 (종류 확인 필요)";
  }
  function shortLabel(l) {
    var i = l.indexOf("·");
    var core = /^에어컨/.test(l) && i > 0 ? l.slice(i + 1).trim() : (i > 0 ? l.slice(0, i).trim() : l);
    return core.replace(/\s*(\d+)대/g, "$1").replace(/, /g, "·");
  }

  // ---------- 업무 문자인지 대충 거르기 ----------
  var NEGATIVE = /인증번호|\[Web발신\]|\(광고\)|무료수신거부|출금|입금|승인|결제|잔액|택배|배송|080-\d/;
  var KEYWORD = /청소|에어컨|에어콘|세척|입주|이사|견적|예약|방문|벽걸이|스탠드|시스템/;
  function prefilter(text) {
    if (NEGATIVE.test(text)) return false;
    return KEYWORD.test(text) || /\d|오늘|내일|모레|요일/.test(text);
  }

  // ---------- 규칙으로 문자 읽기 (무료) ----------
  function parse(raw, now) {
    var T = day0(now), t = String(raw).replace(/\s+/g, " "), m;
    var r = { via: "rule", isJob: false, date: null, dateText: "", start: null, timeText: "", win: null, full: false,
      dur: 60, label: "", customer: "", phone: "", address: "" };

    m = t.match(/01[016789][-\s.]?\d{3,4}[-\s.]?\d{4}/);
    if (m) { r.phone = m[0].replace(/[\s.]/g, "-"); t = t.replace(m[0], " "); }

    var rel = { "오늘": 0, "금일": 0, "내일": 1, "낼": 1, "모레": 2, "글피": 3 };
    function setMD(mo, d, txt) {
      var dt = new Date(T.getFullYear(), mo - 1, d);
      if (dt < addDays(T, -30)) dt = new Date(T.getFullYear() + 1, mo - 1, d);
      if (dt.getMonth() === mo - 1) { r.date = key(dt); r.dateText = txt; }
    }
    if ((m = t.match(/(오늘|금일|내일|모레|글피)/))) { r.date = key(addDays(T, rel[m[1]])); r.dateText = m[1]; }
    else if ((m = t.match(/(이번\s*주|금주|다다음\s*주|다음\s*주|담주)?\s*([월화수목금토일])\s*요일/))) {
      var idx = DOW.indexOf(m[2]), off = (idx + 6) % 7, w = (m[1] || "").replace(/\s/g, ""), mon = mondayOf(T), d;
      if (w === "다음주" || w === "담주") d = addDays(mon, 7 + off);
      else if (w === "다다음주") d = addDays(mon, 14 + off);
      else if (w === "이번주" || w === "금주") d = addDays(mon, off);
      else d = addDays(T, ((idx - T.getDay() + 7) % 7) || 7);
      r.date = key(d); r.dateText = m[0].trim();
    }
    else if ((m = t.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/))) setMD(+m[1], +m[2], m[0]);
    else if ((m = t.match(/(?:^|[^\d])(\d{1,2})\s*[\/.]\s*(\d{1,2})(?!\d)/)) && +m[1] >= 1 && +m[1] <= 12 && +m[2] >= 1 && +m[2] <= 31) setMD(+m[1], +m[2], m[1] + "/" + m[2]);
    else if ((m = t.match(/(?:^|[^\d월])(\d{1,2})\s*일(?![\d요])/))) {
      var dd = +m[1], dt = new Date(T.getFullYear(), T.getMonth(), dd);
      if (dt < T) dt = new Date(T.getFullYear(), T.getMonth() + 1, dd);
      if (dd >= 1 && dd <= 31) { r.date = key(dt); r.dateText = m[1] + "일"; }
    }

    if ((m = t.match(/(오전|오후|아침|낮|저녁|밤)?\s*(\d{1,2})\s*시(?!스템|간)\s*(반|(\d{1,2})\s*분)?/))) {
      var h = +m[2], mi = m[3] === "반" ? 30 : (m[4] ? +m[4] : 0), ap = m[1] || "";
      if (/오후|저녁|밤/.test(ap) && h < 12) h += 12;
      else if (ap === "낮" && h <= 6) h += 12;
      else if (!ap && h >= 1 && h <= 7) h += 12;
      r.start = h * 60 + mi; r.timeText = m[0].trim();
    } else if ((m = t.match(/(?:^|[^\d])(\d{1,2}):(\d{2})/))) {
      r.start = +m[1] * 60 + +m[2]; r.timeText = m[1] + ":" + m[2];
    } else if (/오전/.test(t)) { r.win = { from: 540, to: 690 }; r.timeText = "오전 중"; }
    else if (/오후|점심\s*이후/.test(t)) { r.win = { from: 780, to: 1050 }; r.timeText = "오후 중"; }
    else if (/저녁|퇴근/.test(t)) { r.win = { from: 1020, to: 1080 }; r.timeText = "저녁"; }
    else { r.win = { from: 540, to: 1080 }; r.timeText = ""; }

    var items = [], dur = 0, kind = "other";
    [["시스템", /시스템|천장형|천정형/], ["투인원", /투인원|2\s?in\s?1/i], ["스탠드", /스탠드/], ["벽걸이", /벽걸이/]].forEach(function (k) {
      var km = t.match(k[1]); if (!km) return;
      var after = t.slice(km.index + km[0].length), cm = after.match(/^[^\d.,!?]{0,10}?(\d+)\s*대/);
      var n = cm ? +cm[1] : 1; items.push(k[0] + " " + n + "대"); dur += UNIT[k[0]] * n;
    });
    if (!items.length && /에어컨|에어콘/.test(t)) { var cm2 = t.match(/(\d+)\s*대/), n2 = cm2 ? +cm2[1] : 1; items.push("벽걸이 " + n2 + "대"); dur += UNIT["벽걸이"] * n2; }
    var py = t.match(/(\d+)\s*평/), pyeong = py ? +py[1] : null;
    if (/(입주|이사)\s*청소/.test(t)) { kind = /입주/.test(t) ? "move_in" : "move_out"; r.full = true; dur = FULL_DUR; }
    else if (/거주\s*청소|집\s*청소|(?<!\d\s*)대청소/.test(t)) { kind = "home"; dur = Math.max(dur, 300); }
    else if (/세탁기/.test(t)) { kind = "washer"; dur = Math.max(dur, UNIT["세탁기"]); }
    else if (items.length) kind = "aircon";
    else dur = 120;
    r.label = labelFor(kind, items, pyeong);
    r.dur = Math.max(60, dur);
    if (r.full) { r.start = FULL_START; r.win = null; r.timeText = r.timeText || "하루 종일"; }

    var nm = t.match(/(?:이름은|성함은|저는|고객명\s*:?)\s*([가-힣]{2,4}?)(?=입니다|이에요|예요|이요|[\s.,!~]|$)/) ||
             t.match(/(?:^|[\s.,~!])([가-힣]{2,4}?)(?=(?:입니다|이에요|예요)(?:[\s.,!~]|$))/);
    if (nm && !/사무실|아파트|가게|매장|빌라|주택|상가|오피스텔|학원|식당/.test(nm[1])) r.customer = nm[1];

    var am = t.match(/(?:[가-힣]+(?:특별시|광역시|시|도)\s)?[가-힣]+(?:구|군)\s[가-힣0-9]+(?:동|읍|면|로|길)(?:\s+[A-Za-z0-9가-힣\-]+){0,5}/);
    if (am) {
      var toks = am[0].split(/\s+/), keep = [], okTok = /(\d|[A-Za-z]|동|호|층|타운|아파트|빌라|자이|힐즈|파크|캐슬|뷰|오피스텔|빌딩|타워|센터|마을|단지|힐스|푸르지오|래미안|시티)$/;
      for (var i = 0; i < toks.length; i++) {
        var tk = toks[i], stripped = tk.replace(/(입니다|이고|이에요|예요|이요|이라서?|에서|이며|고)$/, "");
        if (i < 3) { keep.push(stripped); if (stripped !== tk) break; continue; }
        if (!okTok.test(stripped)) break;
        keep.push(stripped); if (stripped !== tk) break;
      }
      r.address = keep.join(" ");
    }

    r.isJob = !NEGATIVE.test(t) && (KEYWORD.test(t) || (!!r.date && (r.start != null || /가능|되나요|될까요|되세요/.test(t))));
    return r;
  }

  // ---------- AI 결과를 같은 모양으로 바꾸기 ----------
  function fromAI(o, now) {
    var r = { via: "ai", isJob: !!o.is_job_request, date: null, dateText: o.date_text || "", start: null, timeText: "", win: null,
      full: false, dur: 60, label: "", customer: o.customer_name || "", phone: o.phone || "", address: o.address || "" };
    if (o.date && /^\d{4}-\d{2}-\d{2}$/.test(o.date)) r.date = o.date;
    var tm = o.start_time && String(o.start_time).match(/^(\d{1,2}):(\d{2})/);
    if (tm) { r.start = +tm[1] * 60 + +tm[2]; r.timeText = tm[1] + ":" + tm[2]; }
    else {
      var W = { morning: [540, 690, "오전 중"], afternoon: [780, 1050, "오후 중"], evening: [1020, 1080, "저녁"] }[o.time_window];
      r.win = W ? { from: W[0], to: W[1] } : { from: 540, to: 1080 }; r.timeText = W ? W[2] : "";
    }
    var names = { wall: "벽걸이", stand: "스탠드", two_in_one: "투인원", system: "시스템" }, items = [], dur = 0;
    (o.units || []).forEach(function (u) { var n = Math.max(1, u.count | 0), k = names[u.kind]; if (!k) return; items.push(k + " " + n + "대"); dur += UNIT[k] * n; });
    var kind = o.job_type;
    if (kind === "aircon" && !items.length) { items.push("벽걸이 1대"); dur = UNIT["벽걸이"]; }
    if (kind === "move_in" || kind === "move_out") { r.full = true; dur = FULL_DUR; r.start = FULL_START; r.win = null; r.timeText = "하루 종일"; }
    else if (kind === "home") dur = Math.max(dur, 300);
    else if (kind === "washer") dur = Math.max(dur, UNIT["세탁기"]);
    else if (kind !== "aircon") dur = 120;
    r.label = labelFor(kind, items, o.pyeong);
    r.dur = Math.max(60, dur);
    return r;
  }

  // ---------- 일정표와 비교 ----------
  function jobsOn(jobs, k) { return jobs.filter(function (j) { return j.date === k; }).sort(function (a, b) { return a.start - b.start; }); }
  function clash(jobs, k, start, dur, full) {
    var js = jobsOn(jobs, k);
    if (full) return js;
    return js.filter(function (j) { return j.fullDay || (start < j.start + j.dur + BUF && j.start < start + dur + BUF); });
  }
  function fits(jobs, k, start, dur, full, now) {
    if (k === key(day0(now)) && start < minsOf(now) + 60) return false;
    if (!full && (start < OPEN || start + dur > CLOSE)) return false;
    return clash(jobs, k, start, dur, full).length === 0;
  }
  function bestStart(jobs, k, dur, pref, lo, hi, skip, now) {
    var c = [];
    for (var s = lo; s <= hi; s += 30) if (s !== skip) c.push(s);
    c.sort(function (a, b) { return Math.abs(a - pref) - Math.abs(b - pref) || a - b; });
    for (var i = 0; i < c.length; i++) if (fits(jobs, k, c[i], dur, false, now)) return c[i];
    return null;
  }
  function findAlts(r, jobs, n, now) {
    var T = day0(now), tk = key(T), out = [];
    var base = r.date && r.date >= tk ? fromKey(r.date) : addDays(T, 1);
    var pref = r.start != null ? r.start : (r.win ? r.win.from : 600);
    for (var i = 0; i < 40 && out.length < n; i++) {
      var k = key(addDays(base, i));
      if (isClosed(k) || k < tk) continue;
      if (r.full) { if (k !== r.date && fits(jobs, k, FULL_START, FULL_DUR, true, now)) out.push({ date: k, start: FULL_START, sameDay: false }); continue; }
      var s = bestStart(jobs, k, r.dur, pref, OPEN, CLOSE - r.dur, k === r.date ? r.start : null, now);
      if (s != null) out.push({ date: k, start: s, sameDay: k === r.date });
    }
    return out;
  }
  function evaluate(r, jobs, now) {
    var tk = key(day0(now)), ev = { status: "", conflicts: [], start: r.start, alts: [], windowFull: false };
    if (!r.date) ev.status = "nodate";
    else if (r.date < tk || (r.date === tk && r.start != null && !r.full && r.start < minsOf(now))) ev.status = "past";
    else if (isClosed(r.date)) ev.status = "closed";
    else if (r.full) { ev.conflicts = clash(jobs, r.date, 0, 0, true); ev.status = ev.conflicts.length ? "conflict" : "ok"; ev.start = FULL_START; }
    else if (r.start != null) {
      if (r.start < OPEN || r.start + r.dur > CLOSE) ev.status = "outside";
      else { ev.conflicts = clash(jobs, r.date, r.start, r.dur, false); ev.status = ev.conflicts.length ? "conflict" : "ok"; }
    } else {
      var w = r.win || { from: 540, to: 1080 };
      var s = bestStart(jobs, r.date, r.dur, w.from, w.from, Math.min(w.to, CLOSE - r.dur), null, now);
      if (s != null) { ev.status = "suggest"; ev.start = s; }
      else { ev.status = "conflict"; ev.conflicts = jobsOn(jobs, r.date); ev.windowFull = true; }
    }
    if (ev.status !== "ok" && ev.status !== "suggest") ev.alts = findAlts(r, jobs, 3, now);
    return ev;
  }
  function pill(status) {
    if (status === "ok") return { tone: "ok", text: "가능" };
    if (status === "suggest") return { tone: "ok", text: "시간 추천" };
    if (status === "conflict") return { tone: "bad", text: "일정 겹침" };
    return { tone: "warn", text: "확인 필요" };
  }

  // ---------- 고객 답장 ----------
  function replyText(r, ev, sel, added) {
    var head = "안녕하세요 " + (r.customer ? r.customer + "님, " : "") + "위더스환경입니다.\n";
    if (sel) {
      return head + whenText(sel.date, sel.start, r.full) + " " + r.label +
        (added ? " 예약 확정되었습니다.\n방문 전날 다시 연락드리겠습니다. 감사합니다." : " 예약 가능합니다. 이 시간으로 잡아드릴까요?");
    }
    if (ev.status === "nodate") return head + "문의 감사합니다. 원하시는 날짜와 시간을 알려주시면 바로 확인해 드리겠습니다.";
    var why = ev.status === "closed" ? "휴무일이라" : ev.status === "outside" ? "작업 가능 시간(" + fmtTime(OPEN) + "~" + fmtTime(CLOSE) + ")이 아니라" : ev.status === "past" ? "이미 지난 시간이라" : "이미 다른 예약이 있어서";
    var asked = r.date ? fmtDate(r.date) + (r.full ? "" : " " + (r.start != null ? fmtTime(r.start) : r.timeText || "")) : "말씀하신 날짜";
    var t = head + "말씀하신 " + asked.trim() + "에는 " + why + " 어렵습니다.";
    if (ev.alts.length) {
      t += "\n가능한 시간은\n" + ev.alts.map(function (a, i) { return "①②③"[i] + " " + whenText(a.date, a.start, r.full); }).join("\n") +
        "\n입니다. 편하신 시간 알려주시면 바로 잡아드릴게요.";
    }
    return t;
  }

  return {
    OPEN: OPEN, CLOSE: CLOSE, BUF: BUF, FULL_START: FULL_START, FULL_DUR: FULL_DUR, DOW: DOW,
    pad: pad, key: key, fromKey: fromKey, addDays: addDays, day0: day0, minsOf: minsOf, mondayOf: mondayOf, isClosed: isClosed,
    fmtDate: fmtDate, fmtShort: fmtShort, fmtTime: fmtTime, hm: hm, fmtDur: fmtDur, whenText: whenText, shortLabel: shortLabel,
    prefilter: prefilter, parse: parse, fromAI: fromAI, jobsOn: jobsOn, evaluate: evaluate, pill: pill, replyText: replyText
  };
}
