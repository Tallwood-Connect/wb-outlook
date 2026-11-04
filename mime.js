/* Reads the pictures out of an exported email (.eml): every image part that has a Content-ID, so each cid: reference in the
   email's HTML can be matched to its exact picture instead of being guessed from names, shapes or order.
   WBMime.pictures(bytes) -> { byCid: Map(content-id in lower case, without <> -> { type, bytes }), count }
   Plain browser code, no dependencies. Anything it cannot read is skipped (the caller falls back to its other ways). */
(function (root) {
  "use strict";
  const MAX_PICTURE = 30 * 1024 * 1024, MAX_DEPTH = 8;
  function toLatin1(bytes) { // one character per byte, so boundaries and headers can be searched as text
    let out = ""; const step = 32768;
    for (let i = 0; i < bytes.length; i += step) out += String.fromCharCode.apply(null, bytes.subarray(i, Math.min(bytes.length, i + step)));
    return out;
  }
  function fromLatin1(s) { const b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i) & 255; return b; }
  function splitHead(raw) { // { headers: Map(lowercase name -> value), body }
    let i = raw.search(/\r?\n\r?\n/), head, body;
    if (i < 0) { head = raw; body = ""; } else { head = raw.slice(0, i); body = raw.slice(i).replace(/^\r?\n\r?\n/, ""); }
    const headers = new Map();
    for (const line of head.replace(/\r?\n[ \t]+/g, " ").split(/\r?\n/)) {
      const m = /^([^:\s]+):\s*(.*)$/.exec(line); if (m && !headers.has(m[1].toLowerCase())) headers.set(m[1].toLowerCase(), m[2].trim());
    }
    return { headers, body };
  }
  function param(value, name) { // boundary="x" / name=x
    const m = new RegExp("(?:^|;)\\s*" + name + "\\s*=\\s*(\"([^\"]*)\"|[^;\\s]*)", "i").exec(value || "");
    return m ? (m[2] != null ? m[2] : m[1]) : "";
  }
  function decodeBody(body, enc) {
    enc = String(enc || "").toLowerCase();
    if (enc === "base64") {
      const clean = body.replace(/[^A-Za-z0-9+\/=]/g, "");
      try { const bin = atob(clean); return fromLatin1(bin); } catch { return null; }
    }
    if (enc === "quoted-printable") {
      const s = body.replace(/=\r?\n/g, "").replace(/=([0-9A-Fa-f]{2})/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
      return fromLatin1(s);
    }
    return fromLatin1(body);
  }
  function walk(raw, depth, out) {
    if (depth > MAX_DEPTH) return;
    const { headers, body } = splitHead(raw);
    const ct = (headers.get("content-type") || "text/plain").toLowerCase();
    const type = ct.split(";")[0].trim();
    if (type.startsWith("multipart/")) {
      const b = param(headers.get("content-type"), "boundary"); if (!b) return;
      const marker = "--" + b;
      const pieces = body.split(marker);
      for (let i = 1; i < pieces.length; i++) {
        let p = pieces[i]; if (p.startsWith("--")) break; // the closing marker
        p = p.replace(/^\r?\n/, ""); walk(p, depth + 1, out);
      }
      return;
    }
    if (type === "message/rfc822") { walk(body, depth + 1, out); return; }
    if (!type.startsWith("image/")) return;
    const id = (headers.get("content-id") || "").replace(/^\s*<|>\s*$/g, "").trim().toLowerCase();
    if (!id) return;
    const bytes = decodeBody(body, headers.get("content-transfer-encoding"));
    if (!bytes || !bytes.length || bytes.length > MAX_PICTURE) return;
    if (!out.has(id)) out.set(id, { type, bytes });
  }
  function pictures(bytes) {
    const byCid = new Map();
    try { walk(toLatin1(bytes), 0, byCid); } catch { /* unreadable: nothing found */ }
    return { byCid, count: byCid.size };
  }
  root.WBMime = { pictures };
})(typeof window !== "undefined" ? window : globalThis);
