// ============================================================
//  SHARED HELPERS — talking to the Sheet, PIN, dates, messages.
//  Used by index.html (staff) and manager.html. Page-specific
//  code stays in each page, so a change to one page can't break the other.
// ============================================================
(function () {
  const C = window.EOD_CONFIG;
  const $ = function (id) { return document.getElementById(id); };

  function ymd(d) { return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
  function todayStr() { return ymd(new Date()); }
  function parseYmd(s) { const p = s.split("-").map(Number); return new Date(p[0], p[1] - 1, p[2]); }
  function fmtDate(s, long) { return parseYmd(s).toLocaleDateString(undefined, { weekday: long ? "long" : "short", month: "short", day: "numeric" }); }
  function fmtTime(iso) { const t = new Date(iso); return isNaN(t) ? "" : t.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" }); }
  function fmtDateTime(iso) { const t = new Date(iso); return isNaN(t) ? "" : t.toLocaleString(undefined, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }); }
  function toast(msg, ms) {
    const t = $("toast"); if (!t) return;
    t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(function () { t.hidden = true; }, ms || 3200);
  }

  // ---- PIN (one per page: staff form and manager page keep their own) ----
  function makePin(storageKey) {
    return {
      get: function () { try { return localStorage.getItem(storageKey) || window["__pin_" + storageKey] || ""; } catch (e) { return window["__pin_" + storageKey] || ""; } },
      set: function (p) { window["__pin_" + storageKey] = p; try { p ? localStorage.setItem(storageKey, p) : localStorage.removeItem(storageKey); } catch (e) {} }
    };
  }

  // ---- Talking to the Apps Script ----
  // Every call: POST text/plain JSON { v, action, pin, ... }  ->  { ok, apiVersion, ... }
  // onPinProblem(message) is called when the PIN is wrong / locked / not allowed.
  function makeApi(pin, onPinProblem) {
    const ready = /^https:\/\/script\.google\.com\//.test(C.SCRIPT_URL);
    async function call(action, payload) {
      if (!ready) { const e = new Error("not_connected"); e.code = "not_connected"; throw e; }
      const body = Object.assign({ v: C.API_VERSION, action: action, pin: pin.get() }, payload || {});
      const res = await fetch(C.SCRIPT_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(body) });
      let data;
      try { data = await res.json(); } catch (e) { const er = new Error("bad_response"); er.code = "bad_response"; throw er; }
      if (!data.ok) {
        const pinMsgs = {
          bad_pin: "Wrong or changed PIN. Enter it again.",
          locked: "Too many wrong PINs. Wait 10 minutes and try again.",
          manager_only: "That's the staff PIN. This page needs the manager PIN."
        };
        if (pinMsgs[data.error]) {
          pin.set(""); if (onPinProblem) onPinProblem(pinMsgs[data.error]);
          const e = new Error(data.error); e.code = data.error; e.pin = true; throw e;
        }
        if (data.error === "unknown_action" || data.error === "old_client") {
          toast("This page and the Sheet script are different versions. Ask the manager to update.", 6000);
        }
      }
      return data;
    }
    return { ready: ready, call: call };
  }

  // Wire up a PIN screen (#pinGate, #pinForm, #pinIn, #pinMsg, #pinGo). check() resolves true when unlocked.
  function wirePinGate(pin, check, onUnlocked) {
    const show = function (msg) { $("pinMsg").textContent = msg || ""; $("pinGate").hidden = false; setTimeout(function () { $("pinIn").focus(); }, 50); };
    $("pinForm").addEventListener("submit", async function (e) {
      e.preventDefault();
      const p = $("pinIn").value.trim();
      if (!/^\d{4,8}$/.test(p)) { $("pinMsg").textContent = "Enter the PIN (4–8 numbers)."; return; }
      const btn = $("pinGo"); btn.disabled = true; const label = btn.textContent; btn.textContent = "Checking…";
      pin.set(p);
      try { if (await check()) { $("pinGate").hidden = true; $("pinIn").value = ""; onUnlocked(); } }
      catch (err) { if (!err.pin) { pin.set(""); $("pinMsg").textContent = "Couldn't reach the Sheet. Check the connection and try again."; } }
      finally { btn.disabled = false; btn.textContent = label; }
    });
    return show;
  }

  window.EOD = { $: $, ymd: ymd, todayStr: todayStr, parseYmd: parseYmd, fmtDate: fmtDate, fmtTime: fmtTime, fmtDateTime: fmtDateTime, toast: toast, makePin: makePin, makeApi: makeApi, wirePinGate: wirePinGate };
})();
