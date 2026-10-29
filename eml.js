/* Builds an .eml (MIME) file from the parts of an email that is still being written. No dependencies. */
(function (root) {
  const CRLF = "\r\n";
  const utf8 = (s) => new TextEncoder().encode(String(s));
  function b64(bytes) {
    let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }
  const fold76 = (s) => (s.match(/.{1,76}/g) || [""]).join(CRLF);
  // Header text: plain if ASCII, otherwise RFC 2047 encoded words (split so each stays under the line limit).
  function encWord(text) {
    const t = String(text == null ? "" : text).replace(/[\r\n]+/g, " ");
    if (/^[\x20-\x7e]*$/.test(t)) return t;
    const out = []; let cur = "";
    for (const ch of t) {
      if (utf8(cur + ch).length > 36) { out.push(cur); cur = ""; }
      cur += ch;
    }
    if (cur) out.push(cur);
    return out.map((p) => `=?UTF-8?B?${b64(utf8(p))}?=`).join(CRLF + " ");
  }
  const quoteName = (n) => `"${String(n).replace(/(["\\])/g, "\\$1")}"`;
  function addr(a) {
    if (!a || !a.emailAddress) return "";
    const n = a.displayName && a.displayName !== a.emailAddress ? a.displayName : "";
    if (!n) return a.emailAddress;
    return /^[\x20-\x7e]*$/.test(n) ? `${quoteName(n)} <${a.emailAddress}>` : `${encWord(n)} <${a.emailAddress}>`;
  }
  const addrList = (l) => (l || []).map(addr).filter(Boolean).join(`,${CRLF} `);
  const rfcDate = (d) => d.toUTCString().replace("GMT", "+0000");
  const safeName = (n) => String(n || "attachment").replace(/["\\\r\n]/g, "_");
  const boundary = (tag, seed) => `----=_wb_${tag}_${seed}`;

  const textPart = (type, content) =>
    `Content-Type: ${type}; charset=utf-8${CRLF}Content-Transfer-Encoding: base64${CRLF}${CRLF}${fold76(b64(utf8(content)))}`;
  const filePart = (a, inline) => {
    const n = safeName(a.name);
    return [
      `Content-Type: ${a.type || "application/octet-stream"}; name="${n}"`,
      "Content-Transfer-Encoding: base64",
      `Content-Disposition: ${inline ? "inline" : "attachment"}; filename="${n}"`,
      ...(inline ? [`Content-ID: <${a.cid}>`] : []),
      "", fold76(b64(a.bytes)),
    ].join(CRLF);
  };

  // input: { id, from, to[], cc[], bcc[], subject, date, bodyHtml, bodyText, attachments:[{name,type,bytes,isInline}] }
  // Returns a Uint8Array of the .eml file.
  function buildEml(m) {
    const seed = String(m.id || Date.now());
    const inl = [], files = [];
    (m.attachments || []).forEach((a, i) => (a.isInline && (a.type || "").startsWith("image/") ? inl.push(Object.assign({}, a, { cid: `wb-inline-${i}@send-to-wunderbuild` })) : files.push(a)));

    let html = m.bodyHtml || "";
    if (html && inl.length) { // point the body's cid: references at the inline images we embed
      html = html.replace(/cid:([^"'\s)>]+)/gi, (match, id) => {
        const key = id.split("@")[0].toLowerCase();
        const hit = inl.find((a) => String(a.name || "").toLowerCase() === key);
        return hit ? `cid:${hit.cid}` : match;
      });
    }
    const text = m.bodyText || "";

    let body;
    if (html) {
      const alt = boundary("alt", seed);
      const alternative = [`Content-Type: multipart/alternative; boundary="${alt}"`, "",
        `--${alt}`, textPart("text/plain", text), `--${alt}`, textPart("text/html", html), `--${alt}--`].join(CRLF);
      if (inl.length) {
        const rel = boundary("rel", seed);
        body = [`Content-Type: multipart/related; boundary="${rel}"`, "",
          `--${rel}`, alternative, ...inl.flatMap((a) => [`--${rel}`, filePart(a, true)]), `--${rel}--`].join(CRLF);
      } else body = alternative;
    } else body = textPart("text/plain", text);

    const head = [
      `From: ${addr(m.from)}`,
      m.to && m.to.length ? `To: ${addrList(m.to)}` : null,
      m.cc && m.cc.length ? `Cc: ${addrList(m.cc)}` : null,
      m.bcc && m.bcc.length ? `Bcc: ${addrList(m.bcc)}` : null,
      `Subject: ${encWord(m.subject || "")}`,
      `Date: ${rfcDate(m.date || new Date())}`,
      `Message-ID: <${seed}@send-to-wunderbuild.invalid>`,
      "X-Reconstructed-By: Send to Wunderbuild add-in (rebuilt from the draft when it was saved; not the mail server's original)",
      "MIME-Version: 1.0",
    ].filter(Boolean).join(CRLF);

    let out;
    if (files.length) {
      const mix = boundary("mix", seed);
      out = [head, `Content-Type: multipart/mixed; boundary="${mix}"`, "", `--${mix}`, body,
        ...files.flatMap((a) => [`--${mix}`, filePart(a, false)]), `--${mix}--`, ""].join(CRLF);
    } else out = [head, body, ""].join(CRLF);
    return utf8(out);
  }

  root.buildEml = buildEml;
  if (typeof module !== "undefined") module.exports = { buildEml };
})(typeof window !== "undefined" ? window : globalThis);
