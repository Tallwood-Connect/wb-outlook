/* One PDF per email conversation.
   - trim(): keeps only the NEW part of an email (cuts the quoted earlier messages) when it is safe to do so.
   - pageHtml(): lays the messages out as one document, oldest first. Where a message was answered more than once
     (a branch), each branch is indented under it with a line down the side, like replies in a comment thread.
   - packMeta()/unpackMeta(): the messages live inside the PDF itself (a compressed note in the PDF's info), so the
     next save can add to them without reading any other file.
   Plain browser code; needs pdf-lib only for packMeta/unpackMeta. */
(function (root) {
  "use strict";
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  /* ---- the subject that identifies a conversation: no Re:/Fwd:, no job details added by the add-in ---- */
  function normSubject(subject) {
    let t = String(subject || "").replace(/^\s*((re|fw|fwd|aw|sv|vs)\s*:\s*)+/i, "");
    t = t.replace(/^J-\d{5,}(?: - [^()]*?)?(?: \([^()]*\))?\s*-\s*/, "");
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
  function paragraphs(text) { // the lines of an email long enough to recognise again
    return String(text || "").split(/\r?\n|\u2028/).map((l) => l.replace(/\s+/g, " ").trim())
      .filter((l) => l.length >= 25 && !HEADER_LINE.test(l) && !/^-{3,}\s*original message/i.test(l));
  }
  const hashesOf = (text) => [...new Set(paragraphs(text).map(paraKey))];

  /* ---- who a quoted message was from (tells which message this one answers) ---- */
  function parseQuoteHeader(text) {
    const t = String(text || "");
    let m = /(?:^|\n)\s*From:\s*([^\n]+)/i.exec(t);
    let from = m ? m[1].trim() : "", at = null;
    const s = /(?:^|\n)\s*(?:Sent|Date):\s*([^\n]+)/i.exec(t);
    if (!from) { // Gmail / Apple style: "On Fri, 30 Oct 2026 at 11:05, Mark Ellis <mark@x.com> wrote:"
      m = /(?:^|\n)\s*On\s+(.{6,200}?)\s+wrote:/i.exec(t);
      if (m) { const em = /([^<,]*)<([^>]+)>/.exec(m[1]); from = em ? `${em[1].trim()} <${em[2]}>` : m[1]; at = parseWhen(m[1]); }
    } else if (s) at = parseWhen(s[1]);
    return from ? { from: personKey(from), at } : null;
  }
  function parseWhen(str) {
    const c = String(str || "").replace(/\b(mon|tue|wed|thu|fri|sat|sun)[a-z]*,?\s*/ig, "").replace(/\s+at\s+/i, " ").replace(/,/g, " ");
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
  /* trim(html, known): known = Set of paragraph hashes already in the record (or null).
     Returns { html, full, trimmed, quote }: `full` is the whole email's text (used to recognise it later),
     `quote` says who the first quoted message was from (used to place this message in the thread). */
  function trim(html, known, parse) {
    const P = parse || ((h) => new DOMParser().parseFromString(h || "", "text/html"));
    const doc = P(html);
    const fullText = toText(doc.body);
    const out = (o) => Object.assign({ full: fullText, trimmed: false, quote: null, quoteCount: 0, quoteRatio: 0 }, o);
    const cut = findCut(doc);
    if (!cut) return out({ html: doc.body.innerHTML });
    const removed = [];
    for (let cur = cut; cur && cur !== doc.body; cur = cur.parentNode) { removed.push(cur); for (let s = cur.nextSibling; s; s = s.nextSibling) removed.push(s); }
    const quoted = textOf(removed);
    const quote = parseQuoteHeader(quoted);
    const quotedParas = paragraphs(quoted);
    if (!known || !known.size) return out({ html: doc.body.innerHTML, quote, quoteCount: quotedParas.length }); // nothing saved yet: keep it all, so no history is lost
    const missing = quotedParas.filter((p) => !known.has(paraKey(p))).length;
    const qr = { quoteCount: quotedParas.length, quoteRatio: quotedParas.length ? 1 - missing / quotedParas.length : 0 };
    if (quotedParas.length && missing / quotedParas.length > 0.25) return out(Object.assign({ html: doc.body.innerHTML, quote }, qr)); // quotes things not in the record
    for (let c = cut; c && c !== doc.body; c = c.parentNode) { while (c.nextSibling) c.parentNode.removeChild(c.nextSibling); }
    cut.parentNode.removeChild(cut);
    const left = (doc.body.textContent || "").replace(/\s+/g, " ").trim();
    if (!left && !doc.body.querySelector("img")) return out(Object.assign({ html: P(html).body.innerHTML, quote }, qr)); // nothing above the quote: keep it all
    return out(Object.assign({ html: doc.body.innerHTML, trimmed: true, quote }, qr));
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
  const fmtWhen = (iso) => new Date(iso).toLocaleString("en-AU", { weekday: "short", day: "numeric", month: "short", year: "numeric", hour: "numeric", minute: "2-digit", hour12: true });
  function entryHtml(e, linkFor) {
    const dir = e.dir === "sent" ? "Sent" : "Received";
    const colour = e.dir === "sent" ? "#1D7A57" : "#2F5D8C";
    const line = (label, v) => (v ? `<div><span style="color:#666">${label}</span> ${esc(v)}</div>` : "");
    const link = linkFor ? linkFor(e) : "";
    return `<section style="margin:0 0 16px 0">` +
      `<div style="border-left:4px solid ${colour};background:#F3F5F7;padding:7px 12px;font:12px/1.45 Arial,Helvetica,sans-serif;break-inside:avoid">` +
      `<div style="display:flex;justify-content:space-between;gap:10px"><div style="font-size:14px;font-weight:bold">${esc(e.from)}</div>` +
      `<div style="white-space:nowrap"><span style="color:${colour};font-weight:bold">${dir}</span> · ${esc(fmtWhen(e.at))}</div></div>` +
      line("To", e.to) + line("Cc", e.cc) + (e.atts && e.atts.length ? line("Attachments", e.atts.join(", ")) : "") +
      (link ? `<div><a href="${esc(link)}" style="color:#1A5FB4">Open the original email (.eml)</a></div>` : "") + `</div>` +
      `<div style="font:13px/1.45 Arial,Helvetica,sans-serif;padding:8px 4px 0 16px">${e.html}</div></section>`;
  }
  /* A straight run of replies stays at one level. Only where a message was answered more than once does each
     answer get its own indented branch, with a line down the side. */
  const LINES = ["#8FA3B8", "#9BB89B", "#C5A56B", "#B58FB8"];
  function pageHtml(meta, opts) {
    const linkFor = opts && opts.linkFor;
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
    const body = roots.map((r) => draw(r, 0)).join("");
    const first = order.length ? fmtWhen(order[0].at).replace(/, \d+:\d+ [ap]m$/i, "") : "", last = order.length ? fmtWhen(order[order.length - 1].at).replace(/, \d+:\d+ [ap]m$/i, "") : "";
    const branched = [...kids.values()].some((k) => k.length > 1);
    return `<!doctype html><html><head><meta charset="utf-8">` +
      `<meta http-equiv="Content-Security-Policy" content="script-src 'none'; object-src 'none'; frame-src 'none'; connect-src 'none'; base-uri 'none'; form-action 'none'">` +
      `<style>@page{size:A4;margin:16mm 14mm}img{max-width:100%!important;height:auto!important}</style></head>` +
      `<body style="margin:0"><div style="font:13px Arial,Helvetica,sans-serif;border-bottom:2px solid #333;padding-bottom:8px;margin-bottom:16px">` +
      `<div style="font-size:11px;letter-spacing:1px;color:#777;text-transform:uppercase">Email conversation</div>` +
      `<div style="font-size:19px;font-weight:bold;margin:2px 0 4px">${esc(meta.subject || "(No subject)")}</div>` +
      `<div style="color:#555">${order.length} message${order.length === 1 ? "" : "s"} · ${esc(first)}${first !== last ? " to " + esc(last) : ""} · oldest first${branched ? " · indented lines show separate branches of replies" : ""}</div></div>` +
      body + `</body></html>`;
  }

  /* ---- the record kept inside the PDF ---- */
  async function deflate(str) {
    const bytes = new TextEncoder().encode(str);
    if (typeof CompressionStream === "undefined") { let b = ""; for (let i = 0; i < bytes.length; i += 8192) b += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192)); return "R" + btoa(b); }
    const cs = new CompressionStream("deflate"); const w = cs.writable.getWriter(); w.write(bytes); w.close();
    const out = new Uint8Array(await new Response(cs.readable).arrayBuffer());
    let bin = ""; for (let i = 0; i < out.length; i += 8192) bin += String.fromCharCode.apply(null, out.subarray(i, i + 8192));
    return "Z" + btoa(bin);
  }
  async function inflate(s) {
    const bin = atob(s.slice(1)), bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    if (s[0] === "R") return new TextDecoder().decode(bytes);
    const ds = new DecompressionStream("deflate"); const w = ds.writable.getWriter(); w.write(bytes); w.close();
    return new TextDecoder().decode(await new Response(ds.readable).arrayBuffer());
  }
  const KEY = "WBThread";
  async function packMeta(pdfBytes, meta, PDFLib) { // the PDF with the record stored inside it
    const doc = await PDFLib.PDFDocument.load(pdfBytes);
    doc.setTitle(`Email conversation: ${meta.subject || ""}`.slice(0, 200));
    doc.setCreator("Send to Wunderbuild"); doc.setProducer("Send to Wunderbuild");
    doc.getInfoDict().set(PDFLib.PDFName.of(KEY), PDFLib.PDFHexString.fromText(await deflate(JSON.stringify(meta))));
    return doc.save();
  }
  async function unpackMeta(pdfBytes, PDFLib) { // the record, or null if this PDF was not made by this add-in
    try {
      const doc = await PDFLib.PDFDocument.load(pdfBytes, { updateMetadata: false });
      const v = doc.getInfoDict().get(PDFLib.PDFName.of(KEY));
      if (!v) return null;
      const m = JSON.parse(await inflate(v.decodeText()));
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
  function knownHashes(meta) { const s = new Set(); if (meta) for (const e of meta.entries) for (const h of e.hashes || []) s.add(h); return s; }
  function addEntry(meta, subject, e) {
    const m = meta && meta.v === 1 ? meta : { v: 1, key: subjectKey(subject), subject: normSubject(subject), entries: [] };
    m.entries = m.entries.filter((x) => x.id !== e.id);
    m.entries.push(Object.assign({ fromKey: personKey(e.fromAddr || e.from || ""), nameKey: personKey(e.from || "") }, e));
    return m;
  }

  root.WBThread = { normSubject, subjectKey, trim, hashesOf, knownHashes, pageHtml, packMeta, unpackMeta, addEntry, resolveParents, personKey, parseQuoteHeader, belongs, threadKeys };
})(typeof window !== "undefined" ? window : globalThis);
