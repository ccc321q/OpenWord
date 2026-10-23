/* ==========================================================================
   OpenWord — local persistence: preferences, autosave, recent titles.
   Everything lives in localStorage under openword.* — nothing leaves the
   machine. All access is guarded so the app still runs when storage is
   unavailable (private mode, file:// quirks, quota).
   ========================================================================== */
WW.storage = (function (U) {
  'use strict';

  var K_PREFS = 'openword.prefs';
  var K_AUTO = 'openword.autosave';
  var K_RECENT = 'openword.recent';
  var AUTO_DELAY = 2000;

  var DEFAULTS = {
    theme: 'light', zoom: 1, view: 'print',
    ruler: true, marks: false, guides: true,
    defaultFont: 'Calibri', defaultSize: 11,
    spellcheck: true, autosave: true
  };

  var prefs = Object.assign({}, DEFAULTS);
  var quotaWarned = false;
  var available = true;

  function read(key) {
    try {
      var raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : null;
    } catch (e) { available = false; return null; }
  }
  function write(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch (e) {
      if (!quotaWarned && e && (e.name === 'QuotaExceededError' || e.code === 22)) {
        quotaWarned = true;
        WW.events && WW.events.toast &&
          WW.events.toast('Local storage is full — autosave is paused for this session', 'err');
      }
      return false;
    }
  }

  /* --------------------------------------------------------- prefs ---- */
  function loadPrefs() {
    var stored = read(K_PREFS);
    if (stored) prefs = Object.assign(prefs, stored);
    return prefs;
  }
  function getPrefs() { return prefs; }
  function getPref(k) { return prefs[k]; }
  function setPref(k, v) {
    prefs[k] = v;
    write(K_PREFS, prefs);
  }
  function savePrefs(patch) {
    prefs = Object.assign(prefs, patch);
    write(K_PREFS, prefs);
  }
  function resetPrefs() {
    prefs = Object.assign({}, DEFAULTS);
    write(K_PREFS, prefs);
  }

  /* ------------------------------------------------------- autosave --- */
  var autoTimer = null;
  function scheduleAutosave() {
    if (!prefs.autosave || !available) return;
    clearTimeout(autoTimer);
    autoTimer = setTimeout(flushAutosave, AUTO_DELAY);
  }
  function flushAutosave() {
    clearTimeout(autoTimer);
    if (!prefs.autosave || !available || !WW.app) return;
    var doc = WW.app.collectDoc();
    if (!doc) return;
    var ok = write(K_AUTO, {
      title: doc.title, html: doc.html,
      headerHtml: doc.headerHtml, footerHtml: doc.footerHtml,
      page: doc.page, meta: doc.meta, ts: Date.now()
    });
    if (WW.app.setSaveState) WW.app.setSaveState(ok ? 'saved' : 'error');
  }
  function loadAutosave() {
    var a = read(K_AUTO);
    if (!a || typeof a.html !== 'string') return null;
    var empty = !a.html.replace(/<[^>]*>|\s|&nbsp;|\u200b/g, '');
    if (empty && (!a.title || a.title === 'Untitled document')) return null;
    return a;
  }
  function clearAutosave() {
    try { localStorage.removeItem(K_AUTO); } catch (e) { /* ignore */ }
  }

  /* --------------------------------------------------------- recent --- */
  function getRecent() { return read(K_RECENT) || []; }
  function addRecent(name) {
    if (!name) return;
    var list = getRecent().filter(function (r) { return r.name !== name; });
    list.unshift({ name: name, ts: Date.now() });
    write(K_RECENT, list.slice(0, 8));
  }

  return {
    loadPrefs: loadPrefs, getPrefs: getPrefs, getPref: getPref,
    setPref: setPref, savePrefs: savePrefs, resetPrefs: resetPrefs,
    scheduleAutosave: scheduleAutosave, flushAutosave: flushAutosave,
    loadAutosave: loadAutosave, clearAutosave: clearAutosave,
    getRecent: getRecent, addRecent: addRecent,
    isAvailable: function () { return available; }
  };
})(WW.util);
