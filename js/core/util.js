/* ==========================================================================
   OpenWord — utilities
   Global namespace: WW   (classic script so the app also runs from file://)
   ========================================================================== */
window.WW = window.WW || {};

WW.util = (function () {
  'use strict';

  var EMU_PER_IN = 914400;
  var TWIP_PER_IN = 1440;
  var TWIP_PER_PT = 20;
  var PX_PER_PT = 96 / 72;           // CSS px per point
  var CM_PER_IN = 2.54;

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

  function h(tag, attrs, children) {
    var el = document.createElement(tag);
    if (attrs) {
      Object.keys(attrs).forEach(function (k) {
        if (k === 'class') el.className = attrs[k];
        else if (k === 'text') el.textContent = attrs[k];
        else if (k === 'html') el.innerHTML = attrs[k];
        else if (k.slice(0, 2) === 'on' && typeof attrs[k] === 'function') el.addEventListener(k.slice(2), attrs[k]);
        else if (attrs[k] !== null && attrs[k] !== undefined && attrs[k] !== false) el.setAttribute(k, attrs[k]);
      });
    }
    (children || []).forEach(function (c) {
      if (c == null) return;
      el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return el;
  }

  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  /** XML escaping for the OOXML parts we generate. Strips XML-illegal control chars. */
  function escapeXml(s) {
    var out = String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '');
    return out
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }

  function clamp(v, lo, hi) { return Math.min(hi, Math.max(lo, v)); }

  function debounce(fn, ms) {
    var t = null;
    var wrapped = function () {
      var args = arguments, self = this;
      clearTimeout(t);
      t = setTimeout(function () { t = null; fn.apply(self, args); }, ms);
    };
    wrapped.flush = function () { if (t) { clearTimeout(t); t = null; fn(); } };
    wrapped.pending = function () { return t !== null; };
    return wrapped;
  }

  function throttle(fn, ms) {
    var last = 0, timer = null;
    return function () {
      var args = arguments, self = this, now = Date.now();
      var wait = ms - (now - last);
      if (wait <= 0) {
        if (timer) { clearTimeout(timer); timer = null; }
        last = now; fn.apply(self, args);
      } else if (!timer) {
        timer = setTimeout(function () { timer = null; last = Date.now(); fn.apply(self, args); }, wait);
      }
    };
  }

  var uidCounter = 0;
  function uid(prefix) { uidCounter += 1; return (prefix || 'ww') + '-' + Date.now().toString(36) + '-' + uidCounter; }

  /* -------------------------------------------------- units ------------------------------------------------- */
  function ptToTwip(pt) { return Math.round(pt * TWIP_PER_PT); }
  function twipToPt(tw) { return tw / TWIP_PER_PT; }
  function inToTwip(inch) { return Math.round(inch * TWIP_PER_IN); }
  function twipToIn(tw) { return tw / TWIP_PER_IN; }
  function ptToPx(pt) { return pt * PX_PER_PT; }
  function pxToPt(px) { return px / PX_PER_PT; }
  function ptToEmu(pt) { return Math.round(pt / 72 * EMU_PER_IN); }
  function emuToPt(emu) { return emu / EMU_PER_IN * 72; }
  function pxToEmu(px) { return Math.round(px / 96 * EMU_PER_IN); }
  function cmToIn(cm) { return cm / CM_PER_IN; }

  /** Parse a CSS length ("12pt", "1.2em", "16px", "0.5in") into points; returns null if unknown. */
  function cssLengthToPt(value, contextEl, prop) {
    if (value == null) return null;
    var s = String(value).trim();
    if (!s || s === 'normal' || s === 'auto') return null;
    var m = /^(-?[\d.]+)\s*(pt|px|em|rem|in|cm|mm|%)?$/i.exec(s);
    if (!m) return null;
    var n = parseFloat(m[1]), unit = (m[2] || 'px').toLowerCase();
    switch (unit) {
      case 'pt': return n;
      case 'px': return pxToPt(n);
      case 'in': return n * 72;
      case 'cm': return n / CM_PER_IN * 72;
      case 'mm': return n / 10 / CM_PER_IN * 72;
      case 'em':
      case 'rem':
        var base = 12;
        if (unit === 'em' && contextEl) {
          var fs = window.getComputedStyle(contextEl).fontSize;
          base = pxToPt(parseFloat(fs) || 12);
        } else if (contextEl) {
          var root = window.getComputedStyle(document.documentElement).fontSize;
          base = pxToPt(parseFloat(root) || 16);
        }
        return n * base;
      default: return pxToPt(n);
    }
  }

  /* -------------------------------------------------- colour ------------------------------------------------ */
  var HIGHLIGHT_COLORS = {
    yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF',
    red: '#FF0000', darkYellow: '#808000', lightGreen: '#99FF66', lightTurquoise: '#CCFFFF',
    turquoise: '#008080', gray50: '#808080', gray25: '#C0C0C0', violet: '#7F00FF',
    darkBlue: '#00008F', teal: '#008080', black: '#000000', white: '#FFFFFF',
    'light yellow': '#FFFF99', 'light green': '#99FF66', 'light turquoise': '#CCFFFF',
    'gray 50': '#808080', 'gray 25': '#C0C0C0', darkred: '#C00000', blue: '#0000FF',
    pink: '#FF99CC', gold: '#FFD966', orange: '#FF8C00', purple: '#7030A0'
  };

  function rgbStringToHex(str) {
    if (!str) return null;
    str = String(str).trim();
    if (/^#[0-9a-f]{6}$/i.test(str)) return str.toUpperCase();
    if (/^#[0-9a-f]{3}$/i.test(str)) {
      return ('#' + str[1] + str[1] + str[2] + str[2] + str[3] + str[3]).toUpperCase();
    }
    var m = /rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/i.exec(str);
    if (m) {
      var hex = '#' + [1, 2, 3].map(function (i) {
        return ('0' + clamp(parseInt(m[i], 10), 0, 255).toString(16)).slice(-2);
      }).join('');
      return hex.toUpperCase();
    }
    var known = COLOR_NAMES[str.toLowerCase()];
    return known || null;
  }

  var COLOR_NAMES = {
    black: '#000000', silver: '#C0C0C0', gray: '#808080', grey: '#808080', white: '#FFFFFF',
    maroon: '#800000', red: '#FF0000', purple: '#800080', fuchsia: '#FF00FF', green: '#008000',
    lime: '#00FF00', olive: '#808000', yellow: '#FFFF00', navy: '#000080', blue: '#0000FF',
    teal: '#008080', aqua: '#00FFFF', orange: '#FFA500', pink: '#FFC0CB', brown: '#A52A2A',
    transparent: null, auto: null
  };

  function hexToRgb(hex) {
    var h16 = rgbStringToHex(hex);
    if (!h16) return null;
    return [parseInt(h16.substr(1, 2), 16), parseInt(h16.substr(3, 2), 16), parseInt(h16.substr(5, 2), 16)];
  }

  /** Word highlight colour name from a CSS colour, or null when it is not in the Word palette. */
  function hexToHighlightName(hex) {
    var target = rgbStringToHex(hex);
    if (!target) return null;
    var keys = Object.keys(HIGHLIGHT_COLORS);
    for (var i = 0; i < keys.length; i++) {
      if (HIGHLIGHT_COLORS[keys[i]].toUpperCase() === target) return keys[i];
    }
    return null;
  }

  /* -------------------------------------------------- files ------------------------------------------------- */
  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = h('a', { href: url, download: filename });
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 1500);
  }

  function readFileAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(fr.error || new Error('Could not read ' + file.name)); };
      fr.readAsArrayBuffer(file);
    });
  }

  function readFileAsText(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(fr.error || new Error('Could not read ' + file.name)); };
      fr.readAsText(file, 'utf-8');
    });
  }

  function readFileAsDataURL(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(fr.error || new Error('Could not read ' + file.name)); };
      fr.readAsDataURL(file);
    });
  }

  function extensionOf(name) {
    var m = /\.([A-Za-z0-9]+)$/.exec(String(name || ''));
    return m ? m[1].toLowerCase() : '';
  }

  function baseNameOf(name) {
    return String(name || '').replace(/\.[A-Za-z0-9]+$/, '');
  }

  function formatBytes(n) {
    if (!n && n !== 0) return '';
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
    return (n / 1048576).toFixed(2) + ' MB';
  }

  /** Best-effort PNG / JPEG / GIF dimensions straight from the binary (no <img> round trip). */
  function sniffImageSize(bytes) {
    try {
      var u = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      function be(o, n) { var v = 0; for (var i = 0; i < n; i++) v = v * 256 + u[o + i]; return v; }
      function le(o, n) { var v = 0; for (var i = 0; i < n; i++) v += u[o + i] * Math.pow(256, i); return v; }
      if (u[0] === 0x89 && u[1] === 0x50) return { w: be(16, 4), h: be(20, 4), type: 'png' };
      if (u[0] === 0x47 && u[1] === 0x49) return { w: le(6, 2), h: le(8, 2), type: 'gif' };
      if (u[0] === 0xFF && u[1] === 0xD8) {
        var o = 2;
        while (o < u.length - 9) {
          if (u[o] !== 0xFF) { o++; continue; }
          var marker = u[o + 1];
          var len = le(o + 2, 2);
          if (marker >= 0xC0 && marker <= 0xCF && marker !== 0xC4 && marker !== 0xC8 && marker !== 0xCC) {
            return { h: le(o + 5, 2), w: le(o + 7, 2), type: 'jpeg' };
          }
          o += 2 + len;
        }
      }
      if (u[0] === 0x42 && u[1] === 0x4D) return { w: le(18, 4), h: le(22, 4), type: 'bmp' };
      if (String.fromCharCode(u[0], u[1], u[2], u[3]) === 'RIFF' && String.fromCharCode(u[8], u[9], u[10], u[11]) === 'WEBP') {
        return { w: 0, h: 0, type: 'webp' };
      }
    } catch (e) { /* ignore */ }
    return null;
  }

  function dataUrlToBytes(dataUrl) {
    var comma = String(dataUrl).indexOf(',');
    var b64 = comma >= 0 ? String(dataUrl).slice(comma + 1) : String(dataUrl);
    var bin = atob(b64.replace(/\s+/g, ''));
    var out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  function bytesToDataURL(bytes, mime) {
    var bin = '', chunk = 0x8000;
    var view = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    for (var i = 0; i < view.length; i += chunk) {
      bin += String.fromCharCode.apply(null, view.subarray(i, i + chunk));
    }
    return 'data:' + (mime || 'application/octet-stream') + ';base64,' + btoa(bin);
  }

  function mimeFromExt(ext) {
    ext = String(ext || '').toLowerCase();
    return ({
      png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
      webp: 'image/webp', svg: 'image/svg+xml', bmp: 'image/bmp', tiff: 'image/tiff'
    })[ext] || 'image/png';
  }

  function extFromMime(mime) {
    return ({
      'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg', 'image/gif': 'gif',
      'image/webp': 'webp', 'image/svg+xml': 'svg', 'image/bmp': 'bmp', 'image/tiff': 'tiff'
    })[String(mime || '').toLowerCase()] || 'png';
  }

  /** Zero-width characters that should never reach a document file. */
  function stripZwsp(s) {
    return String(s == null ? '' : s).replace(/[\u200B\u200C\u200D\uFEFF]/g, '');
  }

  function pluralize(n, one, many) { return n + ' ' + (n === 1 ? one : (many || one + 's')); }

  function deepClone(obj) {
    if (typeof structuredClone === 'function') { try { return structuredClone(obj); } catch (e) { /* fall through */ } }
    return JSON.parse(JSON.stringify(obj));
  }

  function nowStamp() {
    var d = new Date();
    function p(n) { return ('0' + n).slice(-2); }
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes());
  }

  return {
    $: $, $$: $$, h: h, escapeHtml: escapeHtml, escapeXml: escapeXml, clamp: clamp,
    debounce: debounce, throttle: throttle, uid: uid,
    ptToTwip: ptToTwip, twipToPt: twipToPt, inToTwip: inToTwip, twipToIn: twipToIn,
    ptToPx: ptToPx, pxToPt: pxToPt, ptToEmu: ptToEmu, emuToPt: emuToPt, pxToEmu: pxToEmu, cmToIn: cmToIn,
    cssLengthToPt: cssLengthToPt,
    rgbStringToHex: rgbStringToHex, hexToRgb: hexToRgb, hexToHighlightName: hexToHighlightName,
    HIGHLIGHT_COLORS: HIGHLIGHT_COLORS, COLOR_NAMES: COLOR_NAMES,
    downloadBlob: downloadBlob, readFileAsArrayBuffer: readFileAsArrayBuffer,
    readFileAsText: readFileAsText, readFileAsDataURL: readFileAsDataURL,
    extensionOf: extensionOf, baseNameOf: baseNameOf, formatBytes: formatBytes,
    sniffImageSize: sniffImageSize, dataUrlToBytes: dataUrlToBytes, bytesToDataURL: bytesToDataURL,
    mimeFromExt: mimeFromExt, extFromMime: extFromMime, stripZwsp: stripZwsp,
    pluralize: pluralize, deepClone: deepClone, nowStamp: nowStamp,
    EMU_PER_IN: EMU_PER_IN, TWIP_PER_IN: TWIP_PER_IN
  };
})();
