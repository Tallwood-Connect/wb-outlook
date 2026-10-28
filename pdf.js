/* Minimal text-only PDF writer (A4, Helvetica). No dependencies. */
(function (root) {
  const W = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584, 761, 556, 761, 222, 556, 333, 1000, 556, 556, 333, 1000, 667, 333, 1000, 761, 611, 761, 761, 222, 222, 333, 333, 350, 556, 1000, 333, 1000, 500, 333, 944, 761, 500, 667, 278, 333, 556, 556, 556, 556, 260, 556, 333, 737, 370, 556, 584, 333, 737, 333, 400, 584, 333, 333, 333, 556, 537, 278, 333, 333, 365, 556, 834, 834, 834, 611, 667, 667, 667, 667, 667, 667, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, 556, 556, 556, 556, 556, 556, 889, 500, 556, 556, 556, 556, 278, 278, 278, 278, 556, 556, 556, 556, 556, 556, 556, 584, 611, 556, 556, 556, 556, 500, 556, 500];
  const B = [278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584, 761, 556, 761, 278, 556, 500, 1000, 556, 556, 333, 1000, 667, 333, 1000, 761, 611, 761, 761, 278, 278, 500, 500, 350, 556, 1000, 333, 1000, 556, 333, 944, 761, 500, 667, 278, 333, 556, 556, 556, 556, 280, 556, 333, 737, 370, 556, 584, 333, 737, 333, 400, 584, 333, 333, 333, 611, 556, 278, 333, 333, 365, 556, 834, 834, 834, 611, 722, 722, 722, 722, 722, 722, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, 556, 556, 556, 556, 556, 556, 889, 556, 556, 556, 556, 556, 278, 278, 278, 278, 611, 611, 611, 611, 611, 611, 611, 584, 611, 611, 611, 611, 611, 556, 611, 556];
  const CP1252 = { 0x20AC:0x80,0x201A:0x82,0x0192:0x83,0x201E:0x84,0x2026:0x85,0x2020:0x86,0x2021:0x87,0x02C6:0x88,0x2030:0x89,0x0160:0x8A,0x2039:0x8B,0x0152:0x8C,0x017D:0x8E,0x2018:0x91,0x2019:0x92,0x201C:0x93,0x201D:0x94,0x2022:0x95,0x2013:0x96,0x2014:0x97,0x02DC:0x98,0x2122:0x99,0x0161:0x9A,0x203A:0x9B,0x0153:0x9C,0x017E:0x9E,0x0178:0x9F };
  const PAGE_W = 595, PAGE_H = 842, MARGIN = 50;

  // Unicode string -> array of cp1252 byte codes (unsupported characters become "?").
  function toCp(str) {
    const out = [];
    for (const ch of String(str).replace(/\u00A0/g, " ").replace(/\t/g, "    ")) {
      const c = ch.codePointAt(0);
      if (c >= 32 && c < 127) out.push(c);
      else if (c >= 160 && c < 256) out.push(c);
      else if (CP1252[c]) out.push(CP1252[c]);
      else if (c === 10 || c === 13) continue;
      else out.push(63);
    }
    return out;
  }
  const width = (codes, size, bold) => codes.reduce((s, c) => s + (bold ? B : W)[c - 32], 0) * size / 1000;
  const esc = (codes) => codes.map((c) => (c === 40 || c === 41 || c === 92) ? "\\" + String.fromCharCode(c) : (c < 127 ? String.fromCharCode(c) : "\\" + c.toString(8).padStart(3, "0"))).join("");

  // Wrap one paragraph to maxW points; returns arrays of cp1252 codes.
  function wrap(text, size, bold, maxW) {
    const lines = [];
    let cur = [];
    for (const word of toCp(text).reduce((a, c) => { if (c === 32) a.push([]); else { if (!a.length) a.push([]); a[a.length - 1].push(c); } return a; }, [])) {
      let w = word;
      while (width(w, size, bold) > maxW) { // break very long words (URLs)
        let n = w.length; while (n > 1 && width(w.slice(0, n), size, bold) > maxW) n--;
        if (cur.length) { lines.push(cur); cur = []; }
        lines.push(w.slice(0, n)); w = w.slice(n);
      }
      const trial = cur.length ? cur.concat([32], w) : w;
      if (width(trial, size, bold) <= maxW) cur = trial; else { lines.push(cur); cur = w; }
    }
    lines.push(cur);
    return lines;
  }

  // meta: [[label, value], ...]; returns Uint8Array
  function emailToPdf({ subject, meta, body }) {
    const pages = [[]]; let y = PAGE_H - MARGIN;
    const maxW = PAGE_W - 2 * MARGIN;
    const put = (x, size, bold, codes) => pages[pages.length - 1].push(`BT /${bold ? "F2" : "F1"} ${size} Tf 1 0 0 1 ${x} ${y.toFixed(1)} Tm (${esc(codes)}) Tj ET`);
    const need = (h) => { if (y - h < MARGIN) { pages.push([]); y = PAGE_H - MARGIN; } };

    for (const l of wrap(subject || "(No subject)", 15, true, maxW)) { need(20); y -= 18; put(MARGIN, 15, true, l); }
    y -= 6;
    for (const [label, value] of meta) {
      if (!value) continue;
      const lw = 62;
      wrap(value, 10, false, maxW - lw).forEach((l, i) => {
        need(14); y -= 13;
        if (i === 0) put(MARGIN, 10, true, toCp(label));
        put(MARGIN + lw, 10, false, l);
      });
    }
    y -= 8; need(10);
    pages[pages.length - 1].push(`0.6 G 0.5 w ${MARGIN} ${y.toFixed(1)} m ${PAGE_W - MARGIN} ${y.toFixed(1)} l S 0 G`);
    y -= 6;
    for (const para of String(body || "").replace(/\r\n?/g, "\n").split("\n")) {
      for (const l of wrap(para, 10, false, maxW)) { need(14); y -= 13; if (l.length) put(MARGIN, 10, false, l); }
    }

    // Assemble objects: 1 catalog, 2 pages, 3 F1, 4 F2, then page/content pairs.
    const objs = [];
    const kids = pages.map((_, i) => `${5 + i * 2} 0 R`).join(" ");
    objs[1] = "<< /Type /Catalog /Pages 2 0 R >>";
    objs[2] = `<< /Type /Pages /Kids [${kids}] /Count ${pages.length} >>`;
    objs[3] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>";
    objs[4] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>";
    pages.forEach((ops, i) => {
      const stream = ops.join("\n");
      objs[5 + i * 2] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${PAGE_W} ${PAGE_H}] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents ${6 + i * 2} 0 R >>`;
      objs[6 + i * 2] = `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`;
    });
    let pdf = "%PDF-1.4\n"; const offs = [];
    for (let i = 1; i < objs.length; i++) { offs[i] = pdf.length; pdf += `${i} 0 obj\n${objs[i]}\nendobj\n`; }
    const xref = pdf.length;
    pdf += `xref\n0 ${objs.length}\n0000000000 65535 f \n` + offs.slice(1).map((o) => String(o).padStart(10, "0") + " 00000 n \n").join("");
    pdf += `trailer\n<< /Size ${objs.length} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`;
    const bytes = new Uint8Array(pdf.length);
    for (let i = 0; i < pdf.length; i++) bytes[i] = pdf.charCodeAt(i) & 255; // all ASCII by construction
    return bytes;
  }

  root.emailToPdf = emailToPdf;
  if (typeof module !== "undefined") module.exports = { emailToPdf };
})(typeof window !== "undefined" ? window : globalThis);
