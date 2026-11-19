// The manifest version shown in the panel and in the button's hover text must match the manifest files.
const fs = require("fs"), path = require("path");
const root = path.join(__dirname, "..");
const read = (f) => fs.readFileSync(path.join(root, f), "utf8");
let bad = 0;
const check = (ok, msg) => { console.log(`${ok ? "ok  " : "FAIL"} ${msg}`); if (!ok) bad++; };
const js = read("taskpane.js");
const m = js.match(/MANIFEST_VERSION = [^?]*\? "([\d.]+)" : "([\d.]+)"/);
check(!!m, "taskpane.js declares MANIFEST_VERSION for dev and live");
for (const [file, want, label] of [["manifest.xml", m && m[2], "Manifest"], ["manifest-dev.xml", m && m[1], "Dev manifest"]]) {
  const x = read(file), v = (x.match(/<Version>([\d.]+)</) || [])[1];
  check(v === want, `${file} version ${v} matches taskpane.js (${want})`);
  const tips = [...x.matchAll(/id="(?:Button|Compose)\.Tip" DefaultValue="([^"]*)"/g)].map((t) => t[1]);
  check(tips.length === 2 && tips.every((t) => t.endsWith(`${label} ${v}`)), `${file} hover texts end with "${label} ${v}"`);
}
if (bad) { console.log(`${bad} failed`); process.exit(1); }
console.log("all passed");
