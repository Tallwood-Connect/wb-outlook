// Which copy of the add-in this person opens: live (site root) or beta (/beta/), decided by the private tester list
// held in the Worker (POST /channel). Plain old JavaScript, no dependencies. The dev copy (/dev/) never switches.
// Any failure (no key, Worker not reachable, no e-mail address) leaves the person where they are.
(function () {
  try {
    var path = location.pathname, onBeta = /\/beta\//.test(path);
    if (/\/dev\//.test(path) || /[?&]channel=stay/.test(location.search)) return;
    var CK = "wb_channel", TTL = 6 * 60 * 60 * 1000, now = Date.now();
    var liveUrl = location.origin + path.replace(/\/beta\//, "/") + location.search;
    var betaUrl = location.origin + path.replace(/\/([^\/]*)$/, "/beta/$1") + location.search;
    var go = function (url) { location.replace(url); };
    var cached = null;
    try { cached = JSON.parse(localStorage.getItem(CK) || "null"); } catch (e) { cached = null; }
    if (cached && now - cached.t < TTL) { // decided recently: no question to the Worker
      if (cached.c === "beta" && !onBeta) go(betaUrl);
      else if (cached.c === "live" && onBeta) go(liveUrl);
      return;
    }
    var key = null, base = null;
    try { key = localStorage.getItem("wb_key"); base = localStorage.getItem("wb_base"); } catch (e) { /* no storage */ }
    if (!key || !base || /api\.wunderbuild\.com/i.test(base)) return; // no Worker in use
    var root = base.replace(/\/+$/, "").replace(/\/api$/, "");
    var ask = function (email) {
      fetch(root + "/channel", { method: "POST", headers: { "x-api-key": key, "Content-Type": "application/json" }, body: JSON.stringify({ email: email }) })
        .then(function (r) { return r.ok ? r.json() : null; })
        .then(function (j) {
          if (!j || (j.channel !== "beta" && j.channel !== "live")) return;
          try { localStorage.setItem(CK, JSON.stringify({ c: j.channel, t: Date.now() })); } catch (e) { /* ignore */ }
          if (j.channel === "beta" && !onBeta) go(betaUrl);
          else if (j.channel === "live" && onBeta) go(liveUrl);
        })
        .catch(function () { /* stay where we are */ });
    };
    var start = function () {
      var email = "";
      try { email = Office.context.mailbox.userProfile.emailAddress || ""; } catch (e) { /* not available */ }
      if (email) ask(email);
    };
    if (typeof Office !== "undefined" && Office.onReady) Office.onReady(start);
  } catch (e) { /* never get in the way of the panel */ }
})();
