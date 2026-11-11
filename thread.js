/* One PDF per email conversation.
   - trim(): keeps only the NEW part of an email (cuts the quoted earlier messages) when it is safe to do so.
   - pageHtml(): lays the messages out as one document, oldest first. Where a message was answered more than once
     (a branch), each branch is indented under it with a line down the side, like replies in a comment thread.
   - packMeta()/unpackMeta(): the messages live inside the PDF itself (a small compressed file embedded in the PDF),
     so the next save can add to them without reading any other file.
   - sanitize(): cleans an email's HTML before it goes into the PDF page (no scripts, forms, frames, handlers).
   Plain browser code; needs pdf-lib only for packMeta/unpackMeta. */
(function (root) {
  "use strict";
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  /* ---- the subject that identifies a conversation: no Re:/Fwd:, no job details added by the add-in ---- */
  // Job details at the start of a subject, in any of the layouts staff use: "J-01041 - Name (address) - ", "J-01041 Name - ", "J-01041 - Name"
  const JOB_LEAD = /^J-?\d{4,}(?:\s*[-\u2013]\s*|\s+)((?:(?!\s[-\u2013]\s).)*?)(?:\s+[-\u2013]\s+|\s*$)/i;
  function normSubject(subject) {
    let t = String(subject || "").replace(/^\s*((re|fw|fwd|aw|sv|vs)\s*:\s*)+/i, "");
    t = t.replace(/^\s*\[[^\]]*J-?\d{4,}[^\]]*\]\s*/i, ""); // "[J-01041 - Name] " (layout used since v1.70)
    t = t.replace(JOB_LEAD, "");
    t = t.replace(/^\s*((re|fw|fwd|aw|sv|vs)\s*:\s*)+/i, "");
    return t.replace(/\s+/g, " ").trim();
  }
  const subjectKey = (subject) => normSubject(subject).toLowerCase();

  /* ---- small stable hash (same in every browser) ---- */
  function h32(str) {
    let h = 2166136261;
    for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619); }
    return h >>> 0;
  }
  const paraKey = (p) => h32(p.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim());
  const HEADER_LINE = /^(from|sent|to|cc|bcc|subject|date|importance|attachments?):/i;
  const ATTRIBUTION = /^on\b.{6,300}\bwrote:?$/i; // Gmail / Apple "On Fri, 30 Oct 2026 at 11:05, Mark wrote:"
  const normLine = (l) => l.replace(/\s+/g, " ").trim();
  const usable = (l) => l.replace(/[^a-z0-9]/gi, "").length >= 2 && !HEADER_LINE.test(l) && !ATTRIBUTION.test(l) && !/^-{3,}\s*original message/i.test(l) && !/^_{5,}$/.test(l);
  function lineList(text) { // every real line of an email: { text, k, long }
    return String(text || "").split(/\r?\n|\u2028/).map(normLine).filter(usable).map((l) => ({ text: l, k: paraKey(l), long: l.length >= 25 }));
  }
  function paragraphs(text) { return lineList(text).filter((l) => l.long).map((l) => l.text); } // the lines long enough to recognise again
  const hashesOf = (text) => [...new Set(paragraphs(text).map(paraKey))];
  const pairKey = (a, b) => h32(a + "|" + b);
  // Short lines ("Yes", "Thanks,") are recognised by the line next to them, so a new short answer typed between old lines stands out.
  function pairsOf(text) {
    const ls = lineList(text), out = new Set();
    for (let i = 1; i < ls.length; i++) if (!ls[i].long || !ls[i - 1].long) out.add(pairKey(ls[i - 1].k, ls[i].k));
    return [...out];
  }

  /* ---- who a quoted message was from (tells which message this one answers) ---- */
  function parseQuoteHeader(text) {
    const t = String(text || "");
    let m = /(?:^|\n)\s*From:\s*([^\n]+)/i.exec(t);
    let from = m ? m[1].trim() : "", at = null;
    const s = /(?:^|\n)\s*(?:Sent|Date):\s*([^\n]+)/i.exec(t);
    if (!from) { // Gmail / Apple style: "On Fri, 30 Oct 2026 at 11:05, Mark Ellis <mark@x.com> wrote:"
      m = /(?:^|\n)\s*On\s+(.{6,200}?)\s+wrote:/i.exec(t);
      if (m) { const a = splitAttribution(m[1]); from = a.addr ? `${a.who} <${a.addr}>` : a.who || m[1]; at = parseWhen(a.when); }
    } else if (s) at = parseWhen(s[1]);
    return from ? { from: personKey(from), at } : null;
  }
  // "Fri, Oct 30, 2026 at 1:46 PM Tallwood | Jason <jason@x.com>" -> { when: "Fri, Oct 30, 2026 at 1:46 PM", who: "Tallwood | Jason", addr: "jason@x.com" }
  // (the date comes first, then the sender; Gmail, Apple Mail and Outlook all write it this way, with or without commas)
  function splitAttribution(str) {
    let s = String(str || "").replace(/\s+/g, " ").trim(), addr = "";
    const em = /<([^<>]+@[^<>]+)>\s*$/.exec(s); if (em) { addr = em[1].trim(); s = s.slice(0, em.index).trim(); }
    let when = "", who = s;
    const t = /^(.*?\b\d{1,2}[:.]\d{2}(?::\d{2})?(?:\s*[ap]\.?m\b\.?)?)(?=[\s,]|$)\s*,?\s*(.*)$/i.exec(s);
    if (t) { when = t[1]; who = t[2]; }
    else { const d = /^(.*?\b(?:19|20)\d{2})(?=[\s,]|$)\s*,?\s*(?:at\s+)?(.*)$/i.exec(s); if (d) { when = d[1]; who = d[2]; } }
    who = who.replace(/^[\s,"'-]+|[\s,"']+$/g, "");
    return { when, who, addr };
  }
  function parseWhen(str) {
    let c = String(str || "").replace(/\b(mon|tue|wed|thu|fri|sat|sun)[a-z]*,?\s*/ig, "").replace(/\s+at\s+/i, " ").replace(/,/g, " ").replace(/\s+/g, " ").trim();
    c = c.replace(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})\b/, "$2/$1/$3"); // day/month/year as written in Australia
    if (!/\b(19|20)\d{2}\b/.test(c)) return null; // no year: never guess (a bare "Oct 30" would become 2001)
    const d = new Date(c); return isNaN(d) ? null : d.toISOString();
  }
  function personKey(s) { // an address if there is one, otherwise the name
    const a = /<([^>]+@[^>]+)>/.exec(s) || /([^\s<>"']+@[^\s<>"']+)/.exec(s);
    return a ? a[1].toLowerCase() : String(s).toLowerCase().replace(/[^a-z ]+/g, " ").replace(/\s+/g, " ").trim();
  }

  /* ---- cut the quoted earlier messages off an email ---- */
  const BLOCK = /^(P|DIV|TR|LI|BLOCKQUOTE|H[1-6]|HR|TABLE|UL|OL|PRE)$/;
  function toText(node) { // text with a line break at every <br> and block, so each paragraph stands alone
    if (!node) return "";
    if (node.nodeType === 3) return node.nodeValue || "";
    if (node.nodeType !== 1) return "";
    if (node.tagName === "BR") return "\n";
    let t = ""; for (const c of node.childNodes) t += toText(c);
    return BLOCK.test(node.tagName) ? "\n" + t + "\n" : t;
  }
  const textOf = (nodes) => nodes.map(toText).join("\n");
  function findCut(doc) {
    const body = doc.body, cands = [];
    const add = (el) => { if (el && el !== body && body.contains(el)) cands.push(el); };
    add(doc.querySelector("#appendonsend"));
    const rply = doc.querySelector("#divRplyFwdMsg");
    if (rply) { const p = rply.previousElementSibling; add(p && p.tagName === "HR" ? p : rply); }
    add(doc.querySelector(".gmail_quote, .gmail_attr, .yahoo_quoted"));
    add(doc.querySelector("blockquote[type=cite]"));
    for (const d of doc.querySelectorAll("div, p")) { // classic Outlook: a block that starts "From: … Sent:"
      const t = (d.textContent || "").trim().slice(0, 240);
      if (/^from:\s.{0,200}(sent|date):/is.test(t) || /^-{3,}\s*original message\s*-{3,}/i.test(t)) { if (!d.querySelector("div, p")) { add(d); break; } }
    }
    const walker = doc.createTreeWalker(body, 4); // plain text "On <date> … wrote:"
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const t = (n.nodeValue || "").trim();
      if (/^On .{6,160} wrote:$/i.test(t)) { add(n.parentElement); break; }
    }
    if (!cands.length) return null;
    cands.sort((a, b) => (a === b ? 0 : a.compareDocumentPosition(b) & 4 ? -1 : 1)); // earliest in the email first
    return cands[0];
  }
  /* trim(html, known, parse, pairs): known = Set of line hashes already in the record (or null), pairs = Set of line-pair hashes.
     Returns { html, full, trimmed, quote, added }: `full` is the whole email's text (used to recognise it later),
     `quote` says who the first quoted message was from (used to place this message in the thread).
     What is kept: everything above the quoted earlier messages, plus any line inside the quoted part that is NOT already in the
     record (an answer typed between quoted lines, or an edit), shown in a "quote box". Lines already in the record are left out. */
  function trim(html, known, parse, pairs) {
    const P = parse || ((h) => new DOMParser().parseFromString(h || "", "text/html"));
    const doc = P(html);
    const fullText = toText(doc.body);
    const headStyles = Array.from(doc.querySelectorAll("head style")).map((e) => e.outerHTML).join(""); // an email's styles may sit outside its body
    const out = (o) => { const r = Object.assign({ full: fullText, trimmed: false, quote: null, quoteCount: 0, quoteRatio: 0, added: 0 }, o); if (headStyles && r.html != null) r.html = headStyles + r.html; return r; };
    const cut = findCut(doc);
    if (!cut) return out({ html: doc.body.innerHTML });
    const removed = [];
    for (let cur = cut; cur && cur !== doc.body; cur = cur.parentNode) { removed.push(cur); for (let s = cur.nextSibling; s; s = s.nextSibling) removed.push(s); }
    const quoted = textOf(removed);
    const quote = parseQuoteHeader(quoted);
    const qlines = lineList(quoted), quotedParas = qlines.filter((l) => l.long);
    if (!known || !known.size) return out({ html: doc.body.innerHTML, quote, quoteCount: quotedParas.length }); // nothing saved yet: keep it all, so no history is lost
    const missing = quotedParas.filter((l) => !known.has(l.k)).length;
    const qr = { quoteCount: quotedParas.length, quoteRatio: quotedParas.length ? 1 - missing / quotedParas.length : 0 };
    if (quotedParas.length && missing / quotedParas.length > 0.5) return out(Object.assign({ html: doc.body.innerHTML, quote }, qr)); // mostly history the record does not have yet: keep it all
    const havePairs = pairs && pairs.size > 0;
    const isNew = qlines.map((l, i) => {
      if (l.long) return !known.has(l.k);
      if (!havePairs) return false; // an older record: short lines are taken to be old
      const prev = qlines[i - 1], next = qlines[i + 1];
      return !((prev && pairs.has(pairKey(prev.k, l.k))) || (next && pairs.has(pairKey(l.k, next.k))));
    });
    const added = qlines.filter((l, i) => isNew[i]);
    // each new line is shown under the line it was typed after (greyed), so an answer reads with its question
    const boxLines = []; qlines.forEach((l, i) => { if (!isNew[i]) return; if (i > 0 && !isNew[i - 1]) boxLines.push({ text: qlines[i - 1].text.length > 140 ? qlines[i - 1].text.slice(0, 137) + "…" : qlines[i - 1].text, ctx: true }); boxLines.push(l); });
    for (let c = cut; c && c !== doc.body; c = c.parentNode) { while (c.nextSibling) c.parentNode.removeChild(c.nextSibling); }
    cut.parentNode.removeChild(cut);
    const left = (doc.body.textContent || "").replace(/\s+/g, " ").trim();
    if (!left && !doc.body.querySelector("img") && !added.length) return out(Object.assign({ html: P(html).body.innerHTML, quote }, qr)); // nothing new anywhere: keep it all
    let h = doc.body.innerHTML;
    if (added.length) {
      h += `<div style="margin:10px 0 4px;padding:6px 10px;border-left:3px solid #B9A24C;background:#FBF8EC;font-size:12px;line-height:1.45">` +
        `<div style="color:#7A6A1F;font-size:11px;margin-bottom:3px">Added within the quoted earlier message</div>` +
        boxLines.map((l) => (l.ctx ? `<div style="color:#8A8A8A;margin-top:4px">${esc(l.text)}</div>` : `<div style="padding-left:10px">${esc(l.text)}</div>`)).join("") + `</div>`;
    }
    return out(Object.assign({ html: h, trimmed: true, quote, added: added.length }, qr));
  }

  /* ---- clean an email's HTML before it goes into the PDF page ---- */
  // An email's own <style> rules are kept, but fenced in: every rule only applies inside that email's own block (.scope),
  // so one email cannot restyle another or the page. @page, @font-face, @import and animations are dropped.
  const cleanDecl = (d) => d.replace(/position\s*:\s*(fixed|absolute|sticky)/gi, "position:static").replace(/expression\s*\(/gi, "x(").replace(/behavior\s*:[^;]*;?/gi, "").replace(/@import/gi, "").replace(/-moz-binding[^;]*;?/gi, "");
  function scopeSel(sel, scope) { const t = sel.replace(/^(html|body)\b/i, "").trim(); return t ? "." + scope + " " + t : "." + scope; }
  function scopeCss(css, scope) {
    css = String(css || "").replace(/\/\*[\s\S]*?\*\//g, "").replace(/<!--|-->/g, "");
    let out = "", i = 0;
    while (i < css.length) {
      const open = css.indexOf("{", i); if (open < 0) break;
      const head = css.slice(i, open).trim().replace(/^[\s;]+/, "");
      let depth = 1, j = open + 1;
      while (j < css.length && depth) { const c = css[j]; if (c === "{") depth++; else if (c === "}") depth--; j++; }
      const body = css.slice(open + 1, j - 1);
      if (head.charAt(0) === "@") { if (/^@media/i.test(head)) out += head + "{" + scopeCss(body, scope) + "}"; }
      else if (head) out += head.split(",").map((x) => scopeSel(x.trim(), scope)).join(",") + "{" + cleanDecl(body) + "}";
      i = j;
    }
    return out;
  }
  function sanitize(html, opts, parse) {
    const o = opts || {}, scope = o.scope ? String(o.scope).replace(/[^A-Za-z0-9_-]/g, "") : "";
    const P = parse || ((h) => new DOMParser().parseFromString(h || "", "text/html"));
    const doc = P(html);
    const drop = "script,meta,link,base,iframe,frame,frameset,object,embed,applet,form,noscript,title,template" + (o.keepStyle || scope ? "" : ",style");
    for (const el of Array.from(doc.querySelectorAll(drop))) el.parentNode && el.parentNode.removeChild(el);
    for (const el of Array.from(doc.querySelectorAll("style"))) {
      if (scope) el.textContent = scopeCss(el.textContent, scope);
      else el.textContent = (el.textContent || "").replace(/@import[^;]*;?/gi, "").replace(/expression\s*\(/gi, "x(");
    }
    const bad = /^\s*(javascript|vbscript|data\s*:\s*text\/html)/i;
    for (const el of Array.from(doc.querySelectorAll("*"))) {
      for (const at of Array.from(el.attributes)) {
        const n = at.name.toLowerCase(), v = at.value || "";
        if (n.indexOf("on") === 0 || n === "srcset" || n === "formaction" || n === "action" || n === "ping") el.removeAttribute(at.name);
        else if ((n === "href" || n === "src" || n === "background" || n === "xlink:href") && bad.test(v.replace(/[\s\u0000-\u001f]+/g, ""))) el.removeAttribute(at.name);
        else if (n === "style") el.setAttribute("style", cleanDecl(v));
      }
    }
    const inner = (scope || o.keepStyle ? doc.head.innerHTML.replace(/<(?!\/?style\b)[^>]*>/gi, "") : "") + doc.body.innerHTML; // styles that sat in the head come along
    return scope ? `<div class="${scope}">${inner}</div>` : inner;
  }

  /* ---- who answered whom ---- */
  function resolveParents(entries) { // returns Map(id -> parent entry or null)
    const es = entries.slice().sort((a, b) => new Date(a.at) - new Date(b.at) || String(a.id).localeCompare(String(b.id)));
    const idx = new Map(es.map((e, i) => [e.id, i]));
    const byMsg = new Map(); for (const e of es) if (e.msgId) byMsg.set(e.msgId, e);
    const parent = new Map();
    for (const e of es) {
      let p = null;
      for (const ref of [e.inReplyTo, ...(e.refs || []).slice().reverse()]) { if (ref && byMsg.has(ref)) { p = byMsg.get(ref); break; } }
      if (!p && e.quote && e.quote.from) { // the latest earlier message from the person quoted
        const q = e.quote;
        const same = es.filter((x) => x !== e && idx.get(x.id) < idx.get(e.id) && (x.fromKey === q.from || x.nameKey === q.from));
        if (same.length) {
          if (q.at) { const t = new Date(q.at).getTime(); same.sort((a, b) => Math.abs(new Date(a.at) - t) - Math.abs(new Date(b.at) - t)); p = same[0]; }
          else p = same[same.length - 1];
        }
      }
      if (p && idx.get(p.id) >= idx.get(e.id)) p = null; // an answer always comes after the message it answers
      parent.set(e.id, p || null);
    }
    return { parent, order: es };
  }

  /* ---- layout ---- */
  const safeUrl = (u) => (/^https:\/\/[^\s"'<>]+$/i.test(String(u || "")) ? String(u) : "");
  const fmtWhen = (iso) => new Date(iso).toLocaleString("en-AU", { weekday: "short", day: "numeric", month: "short", year: "numeric", hour: "numeric", minute: "2-digit", hour12: true });
  // The grey bar that opens every message: who, received/sent, when, and the people it went to.
  function headerHtml(h) { // { from, dir, when (text), to, cc, atts[], link, emlName }
    const colour = "#4A6785"; // one colour for every message: sent / received depends on who is reading, so it is not shown
    const line = (label, v) => (v ? `<div><span style="color:#666">${label}</span> ${esc(v)}</div>` : "");
    const link = safeUrl(h.link);
    return `<div style="border-left:4px solid ${colour};background:#F3F5F7;padding:7px 12px;font:12px/1.45 Arial,Helvetica,sans-serif;break-inside:avoid">` +
      `<div style="display:flex;justify-content:space-between;gap:10px"><div style="font-size:14px;font-weight:bold">${esc(h.from)}</div>` +
      `<div style="white-space:nowrap">${esc(h.when || "")}</div></div>` +
      line("To", h.to) + line("Cc", h.cc) + (h.atts && h.atts.length ? line("Attachments", h.atts.join(", ")) : "") +
      (link ? `<div><a href="${esc(link)}" style="color:#1A5FB4">Open the original email file in Wunderbuild${h.emlName ? " (" + esc(h.emlName) + ")" : ""}</a></div>` : "") + `</div>`;
  }
  function entryHtml(e, linkFor) {
    const h = headerHtml({ from: e.from, dir: e.dir, when: fmtWhen(e.at), to: e.to, cc: e.cc, atts: e.atts, link: (linkFor && linkFor(e)) || e.link, emlName: e.emlName });
    return `<section style="margin:0 0 16px 0">${h}` +
      `<div style="font:13px/1.45 Arial,Helvetica,sans-serif;padding:8px 4px 0 16px" class="eb">${e.html}</div></section>`;
  }
  /* ---- earlier messages quoted inside an email: their "From: / Sent: / To: / Subject:" blocks become the same bar ---- */
  const HDR_LABEL = /^(from|sent|date|to|cc|bcc|subject|importance|attachments?|reply-to):/i;
  const names = (v) => String(v || "").split(/;\s*/).map((x) => x.replace(/\s*<[^>]*>/g, "").replace(/^["'\s]+|["'\s]+$/g, "")).filter(Boolean).join(", ").slice(0, 300);
  const whenText = (v) => { const raw = String(v || "").trim(); if (!raw || /^\d{1,2}[\/.]\d{1,2}[\/.]\d{2,4}/.test(raw)) return raw; const iso = parseWhen(raw); return iso ? fmtWhen(iso) : raw; }; // day/month dates are shown as written
  function restyleHeaders(html, opts, parse) {
    const P = parse || ((h) => new DOMParser().parseFromString(h || "", "text/html"));
    const me = ((opts && opts.me) || []).map((x) => String(x || "").toLowerCase().trim()).filter(Boolean);
    const doc = P(html), body = doc.body;
    const mk = (h) => { const d = doc.createElement("div"); d.setAttribute("style", "margin:14px 0 6px 0"); d.innerHTML = headerHtml(h); return d; };
    const dropRule = (n) => { // the separator line (or Outlook's top border) that used to mark the break
      let a = n; while (a.parentNode && a.parentNode !== body && !a.previousElementSibling) a = a.parentNode;
      if (a.previousElementSibling && a.previousElementSibling.tagName === "HR") a.previousElementSibling.parentNode.removeChild(a.previousElementSibling);
      for (let p = n.parentNode, k = 0; p && p !== body && k < 4; p = p.parentNode, k++) { const st = p.getAttribute && p.getAttribute("style"); if (st && /border-top/i.test(st)) p.setAttribute("style", st.replace(/border(-top)?\s*:[^;]*;?/gi, "")); }
    };
    const cands = [];
    for (const el of Array.from(body.querySelectorAll("div,p,font,td,blockquote,span"))) {
      const t = toText(el); if (t.length > 1500 || !/^\s*From:/i.test(t) || !/(^|\n)\s*(Sent|Date):/i.test(t)) continue;
      const lines = t.split("\n").map((l) => l.replace(/\s+/g, " ").trim()).filter(Boolean);
      if (lines.length && lines.every((l) => HDR_LABEL.test(l))) cands.push({ el, lines });
    }
    for (const c of cands.filter((x) => !cands.some((o) => o !== x && x.el.contains(o.el)))) { // the innermost element holding only the header lines
      if (!c.el.parentNode) continue;
      const v = {}; for (const l of c.lines) { const m = /^([A-Za-z-]+):\s*(.*)$/.exec(l); if (m && !v[m[1].toLowerCase()]) v[m[1].toLowerCase()] = m[2]; }
      const fromRaw = v.from || "", low = fromRaw.toLowerCase();
      const sent = me.some((x) => x && low.indexOf(x) >= 0);
      const node = mk({ from: names(fromRaw) || "Unknown", dir: sent ? "sent" : "received", when: whenText(v.sent || v.date), to: names(v.to), cc: names(v.cc), atts: v.attachment || v.attachments ? [String(v.attachment || v.attachments).replace(/\s*;\s*/g, ", ")] : [] });
      c.el.parentNode.replaceChild(node, c.el); dropRule(node);
    }
    for (const el of Array.from(body.querySelectorAll(".gmail_attr, div, p"))) { // Gmail / Apple: "On Fri, 30 Oct 2026 at 11:05, Mark Ellis <mark@x.com> wrote:"
      if (!el.parentNode || (el.tagName !== "DIV" && el.tagName !== "P" && !el.classList.contains("gmail_attr"))) continue;
      const t = (el.textContent || "").replace(/\s+/g, " ").trim();
      const m = /^On (.{6,200}?) wrote:$/i.exec(t); if (!m || el.querySelector("div,p")) continue;
      const a = splitAttribution(m[1]); const nm = a.who || a.addr || m[1]; const iso = parseWhen(a.when);
      const node = mk({ from: nm, dir: me.some((x) => x && m[1].toLowerCase().indexOf(x) >= 0) ? "sent" : "received", when: iso ? fmtWhen(iso) : "" });
      el.parentNode.replaceChild(node, el);
    }
    return doc.head.innerHTML.replace(/<(?!\/?style\b)[^>]*>/gi, "") + body.innerHTML;
  }
  /* A picture that appears in several emails (a signature logo) is written into the page once: each use becomes a small placeholder
     <img> that CSS fills with the shared picture. The record kept inside the PDF is not changed, only the page that is drawn. This keeps
     the page small (the Worker and the PDF service have limits on its size) and the drawing identical. */
  function dedupePictures(html) {
    const re = /<img\b[^>]*>/gi, uris = new Map();
    const srcOf = (tag) => { const m = /\ssrc\s*=\s*("(data:image\/[a-z0-9.+-]+;base64,[A-Za-z0-9+\/=]+)"|'(data:image\/[a-z0-9.+-]+;base64,[A-Za-z0-9+\/=]+)')/i.exec(tag); return m ? { whole: m[1], uri: m[2] || m[3] } : null; };
    let m;
    while ((m = re.exec(html))) { const s = srcOf(m[0]); if (s && s.uri.length > 3000) uris.set(s.uri, (uris.get(s.uri) || 0) + 1); }
    const shared = new Map(); let n = 0;
    for (const [u, c] of uris) if (c > 1) shared.set(u, "wbp" + n++);
    if (!shared.size) return { html, css: "" };
    const blank = "data:image/gif;base64,R0lGODlhAQABAAAAACw=";
    const out = html.replace(re, (tag) => {
      const s = srcOf(tag); const cls = s && shared.get(s.uri);
      if (!cls) return tag;
      let t = tag.replace(s.whole, `"${blank}"`);
      if (/\sclass\s*=\s*"/i.test(t)) t = t.replace(/(\sclass\s*=\s*")/i, `$1${cls} `);
      else if (/\sclass\s*=\s*'/i.test(t)) t = t.replace(/(\sclass\s*=\s*')/i, `$1${cls} `);
      else t = t.replace(/^<img\b/i, `<img class="${cls}"`);
      return t;
    });
    let css = ""; for (const [u, cls] of shared) css += `img.${cls}{content:url(${u})}`;
    return { html: out, css };
  }
  /* A straight run of replies stays at one level. Only where a message was answered more than once does each
     answer get its own indented branch, with a line down the side. */
  const LINES = ["#8FA3B8", "#9BB89B", "#C5A56B", "#B58FB8"];
  function pageHtml(meta, opts) {
    const linkFor = opts && opts.linkFor, topLink = safeUrl(opts && opts.topLink);
    const { parent, order } = resolveParents(meta.entries);
    const kids = new Map(order.map((e) => [e.id, []]));
    const roots = [];
    for (const e of order) { const p = parent.get(e.id); (p ? kids.get(p.id) : roots).push(e); }
    const draw = (e, depth) => { // a message, then whatever answered it
      let h = entryHtml(e, linkFor);
      const k = kids.get(e.id);
      if (k.length === 1) h += draw(k[0], depth);
      else for (const c of k) h += `<div style="margin:0 0 6px 10px;padding-left:14px;border-left:3px solid ${LINES[(depth + k.indexOf(c)) % LINES.length]}">${draw(c, depth + 1)}</div>`;
      return h;
    };
    const { html: body, css: picCss } = dedupePictures(roots.map((r) => draw(r, 0)).join(""));
    const first = order.length ? fmtWhen(order[0].at).replace(/, \d+:\d+ [ap]m$/i, "") : "", last = order.length ? fmtWhen(order[order.length - 1].at).replace(/, \d+:\d+ [ap]m$/i, "") : "";
    const branched = [...kids.values()].some((k) => k.length > 1);
    return `<!doctype html><html><head><meta charset="utf-8">` +
      `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data: https:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'">` +
      `<style>@page{size:A4;margin:16mm 14mm}` +
      `.eb img{max-width:100%!important;height:auto!important}` + // pictures keep their proportions
      `.eb *{break-inside:auto!important;page-break-inside:auto!important;min-height:0!important}` + // long emails flow across pages without gaps
      `.eb div:empty,.eb p:empty{display:none!important}` +
      `.eb p[class*=MsoNormal],.eb li[class*=MsoNormal]{margin:0}${picCss}</style></head>` +
      `<body style="margin:0"><div style="font:13px Arial,Helvetica,sans-serif;border-bottom:2px solid #333;padding-bottom:8px;margin-bottom:16px">` +
      `<div style="font-size:11px;letter-spacing:1px;color:#777;text-transform:uppercase">Email conversation</div>` +
      `<div style="font-size:19px;font-weight:bold;margin:2px 0 4px">${esc(meta.subject || "(No subject)")}</div>` +
      `<div style="color:#555">${order.length} message${order.length === 1 ? "" : "s"} · ${esc(first)}${first !== last ? " to " + esc(last) : ""} · oldest first${branched ? " · indented lines show separate branches of replies" : ""}</div>` +
      (topLink ? `<div style="margin-top:3px"><a href="${esc(topLink)}" style="color:#1A5FB4">Open this job's correspondence in Wunderbuild</a></div>` : "") + `</div>` +
      body + `</body></html>`;
  }

  /* ---- the record kept inside the PDF: a small compressed file embedded in it (PDF "attachment") ---- */
  async function deflate(str) { // bytes: "Z" + zlib data, or "R" + plain text when this browser cannot compress
    const bytes = new TextEncoder().encode(str);
    if (typeof CompressionStream === "undefined") { const o = new Uint8Array(bytes.length + 1); o[0] = 82; o.set(bytes, 1); return o; }
    const cs = new CompressionStream("deflate"); const w = cs.writable.getWriter(); w.write(bytes); w.close();
    const z = new Uint8Array(await new Response(cs.readable).arrayBuffer());
    const o = new Uint8Array(z.length + 1); o[0] = 90; o.set(z, 1); return o;
  }
  async function inflate(bytes) {
    const body = bytes.subarray(1);
    if (bytes[0] === 82) return new TextDecoder().decode(body);
    const ds = new DecompressionStream("deflate"); const w = ds.writable.getWriter(); w.write(body); w.close();
    return new TextDecoder().decode(await new Response(ds.readable).arrayBuffer());
  }
  const KEY = "WBThread", ATT_NAME = "wb-conversation.dat";
  async function packMeta(pdfBytes, meta, PDFLib) { // the PDF with the record stored inside it
    const doc = await PDFLib.PDFDocument.load(pdfBytes);
    doc.setTitle(`Email conversation: ${meta.subject || ""}`.slice(0, 200));
    doc.setCreator("Send to Wunderbuild"); doc.setProducer("Send to Wunderbuild");
    await doc.attach(await deflate(JSON.stringify(meta)), ATT_NAME, { mimeType: "application/octet-stream", description: "Record of this email conversation, used by Send to Wunderbuild to add later emails. Safe to ignore.", creationDate: new Date(), modificationDate: new Date() });
    return doc.save();
  }
  function embeddedRecord(doc, PDFLib) { // the bytes of our embedded file, or null
    const { PDFName, PDFDict, PDFArray } = PDFLib;
    const names = doc.catalog.lookupMaybe(PDFName.of("Names"), PDFDict);
    const ef = names && names.lookupMaybe(PDFName.of("EmbeddedFiles"), PDFDict);
    const arr = ef && ef.lookupMaybe(PDFName.of("Names"), PDFArray);
    if (!arr) return null;
    for (let i = 0; i + 1 < arr.size(); i += 2) {
      const nm = arr.lookup(i); if (!nm || typeof nm.decodeText !== "function" || nm.decodeText() !== ATT_NAME) continue;
      const spec = arr.lookupMaybe(i + 1, PDFDict), efd = spec && spec.lookupMaybe(PDFName.of("EF"), PDFDict);
      const stream = efd && efd.lookup(PDFName.of("F"));
      if (stream && PDFLib.decodePDFRawStream) return PDFLib.decodePDFRawStream(stream).decode();
    }
    return null;
  }
  async function unpackMeta(pdfBytes, PDFLib) { // the record, or null if this PDF was not made by this add-in
    try {
      const doc = await PDFLib.PDFDocument.load(pdfBytes, { updateMetadata: false });
      let text = null;
      const emb = embeddedRecord(doc, PDFLib);
      if (emb && emb.length > 1) text = await inflate(emb);
      else { // PDFs made by version 1.62 keep the record in the document info instead
        const v = doc.getInfoDict().get(PDFLib.PDFName.of(KEY));
        if (v) { const s = v.decodeText(), bin = atob(s.slice(1)), bytes = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i); const joined = new Uint8Array(bytes.length + 1); joined[0] = s[0] === "R" ? 82 : 90; joined.set(bytes, 1); text = await inflate(joined); }
      }
      if (!text) return null;
      const m = JSON.parse(text);
      return m && m.v === 1 && Array.isArray(m.entries) ? m : null;
    } catch { return null; }
  }

  /* ---- does this email belong to that conversation? (never decided by the subject alone) ---- */
const threadKeys = (e) => [...new Set([e.msgId, e.inReplyTo, ...(e.refs || [])].filter(Boolean))];
function belongs(meta, e) {
  if (meta.entries.some((x) => x.id === e.id)) return true; // the same email saved again
  const ids = new Set(meta.entries.map((x) => x.msgId).filter(Boolean));
  if (threadKeys(e).some((k) => ids.has(k) && k !== e.msgId)) return true; // it answers (or follows) a message in the record
  if (e.msgId && meta.entries.some((x) => (x.refs || []).includes(e.msgId) || x.inReplyTo === e.msgId)) return true; // a message in the record answers it
  return (e.quoteCount || 0) >= 2 && (e.quoteRatio || 0) >= 0.6; // most of what it quotes is already in the record
}

/* ---- add one email to a record (replacing it if it is already there) ---- */
  function knownHashes(meta, exceptId) { const s = new Set(); if (meta) for (const e of meta.entries) if (e.id !== exceptId) for (const h of e.hashes || []) s.add(h); return s; }
  function knownPairs(meta, exceptId) { const s = new Set(); if (meta) for (const e of meta.entries) if (e.id !== exceptId) for (const h of e.pairs || []) s.add(h); return s; }
  function addEntry(meta, subject, e) {
    const m = meta && meta.v === 1 ? meta : { v: 1, key: subjectKey(subject), subject: normSubject(subject), entries: [] };
    m.entries = m.entries.filter((x) => x.id !== e.id);
    m.entries.push(Object.assign({ fromKey: personKey(e.fromAddr || e.from || ""), nameKey: personKey(e.from || "") }, e));
    return m;
  }

  root.WBThread = { normSubject, subjectKey, trim, sanitize, restyleHeaders, hashesOf, pairsOf, knownHashes, knownPairs, pageHtml, packMeta, unpackMeta, addEntry, resolveParents, personKey, parseQuoteHeader, belongs, threadKeys };
})(typeof window !== "undefined" ? window : globalThis);
