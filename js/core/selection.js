/* ==========================================================================
   OpenWord — selection & DOM position helpers
   A caret position is serialised as { area, path, offset } where `path` is the
   chain of child indices from the editable root. That is enough to restore a
   selection after an undo/redo snapshot swap.
   ========================================================================== */
WW.selection = (function (U) {
  'use strict';

  var BLOCK_TAGS = {
    P: 1, DIV: 1, H1: 1, H2: 1, H3: 1, H4: 1, H5: 1, H6: 1, LI: 1, UL: 1, OL: 1,
    BLOCKQUOTE: 1, PRE: 1, TABLE: 1, TR: 1, TD: 1, TH: 1, HR: 1, FIGURE: 1, SECTION: 1, BODY: 1
  };

  function rootOf(node) {
    if (!node) return null;
    var el = node.nodeType === 1 ? node : node.parentElement;
    return el ? el.closest('[contenteditable="true"]') : null;
  }

  function areaOf(node) {
    var r = rootOf(node);
    return r ? (r.dataset.hf || 'editor') : null;
  }

  function isBlock(node) {
    return !!(node && node.nodeType === 1 && BLOCK_TAGS[node.tagName]);
  }

  function closest(el, selector, boundary) {
    if (!el) return null;
    if (el.nodeType === 3) el = el.parentElement;
    if (!el) return null;
    var node = el;
    while (node && node !== boundary) {
      if (node.nodeType === 1 && node.matches(selector)) return node;
      node = node.parentElement;
    }
    return null;
  }

  /** Index chain from `root` down to `node`. */
  function pathFromRoot(root, node) {
    var path = [];
    var cur = node;
    while (cur && cur !== root) {
      var parent = cur.parentNode;
      if (!parent) return null;
      path.unshift(Array.prototype.indexOf.call(parent.childNodes, cur));
      cur = parent;
    }
    return cur === root ? path : null;
  }

  function nodeFromPath(root, path) {
    var cur = root;
    for (var i = 0; i < path.length && cur; i++) cur = cur.childNodes[path[i]];
    return cur || null;
  }

  /** Serialise the live selection. Returns null when there is no selection inside the app. */
  function capture(win) {
    win = win || window;
    var sel = win.getSelection();
    if (!sel || sel.rangeCount === 0) return null;
    var range = sel.getRangeAt(0);
    var aRoot = rootOf(range.startContainer);
    if (!aRoot) return null;
    var fRoot = rootOf(range.endContainer) || aRoot;
    var rec = {
      area: aRoot.dataset.hf || 'editor',
      anchorPath: pathFromRoot(aRoot, range.startContainer) || [],
      anchorOffset: range.startOffset,
      collapsed: sel.isCollapsed,
      sameArea: fRoot === aRoot
    };
    if (!rec.sameArea) { rec.focusPath = rec.anchorPath; rec.focusOffset = rec.anchorOffset; return rec; }
    rec.focusPath = pathFromRoot(fRoot, range.endContainer) || [];
    rec.focusOffset = range.endOffset;
    return rec;
  }

  function editableByArea(area) {
    if (!area || area === 'editor') return document.getElementById('editor');
    return document.querySelector('[data-hf="' + area + '"]');
  }

  /** Restore a captured selection. `focus:true` also focuses the editable. */
  function restore(rec, focus) {
    if (!rec) return false;
    var root = editableByArea(rec.area);
    if (!root) return false;
    if (focus !== false) {
      try { root.focus({ preventScroll: true }); } catch (e) { try { root.focus(); } catch (e2) { /* noop */ } }
    }
    var a = nodeFromPath(root, rec.anchorPath);
    var f = nodeFromPath(root, rec.focusPath || rec.anchorPath);
    if (!a || !f) return false;
    var range = document.createRange();
    try {
      range.setStart(a, Math.min(rec.anchorOffset, maxOffset(a)));
      range.setEnd(f, Math.min(rec.focusOffset, maxOffset(f)));
    } catch (e) {
      return false;
    }
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    return true;
  }

  function maxOffset(node) {
    if (!node) return 0;
    return node.nodeType === 3 ? node.nodeValue.length : node.childNodes.length;
  }

  function liveRange() {
    var sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return null;
    var r = sel.getRangeAt(0);
    if (!rootOf(r.startContainer)) return null;
    return r;
  }

  function isCollapsed() {
    var sel = window.getSelection();
    return !sel || sel.rangeCount === 0 ? true : sel.isCollapsed;
  }

  function selectNode(node, collapseToEnd) {
    var range = document.createRange();
    range.selectNodeContents(node);
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    if (collapseToEnd) sel.collapseToEnd();
  }

  function caretAtStartOf(node) {
    var range = document.createRange();
    range.setStart(node, 0);
    range.collapse(true);
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }

  function caretAfter(node) {
    var parent = node.parentNode;
    if (!parent) return;
    var range = document.createRange();
    range.setStartAfter(node);
    range.collapse(true);
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }

  /** All text nodes intersected by a range (optionally limited by a filter predicate). */
  function textNodesInRange(range, filter) {
    var out = [];
    if (!range) return out;
    var startNode = range.startContainer;
    var walker = document.createTreeWalker(
      range.commonAncestorContainer.nodeType === 1 ? range.commonAncestorContainer : range.commonAncestorContainer.parentNode,
      NodeFilter.SHOW_TEXT, null);
    var node;
    var started = false;
    while ((node = walker.nextNode())) {
      if (!range.intersectsNode(node)) continue;
      if (!rootOf(node)) continue;
      if (filter && !filter(node)) continue;
      out.push(node);
      if (node === range.endContainer) break;
    }
    return out;
  }

  /** Top-level (or cell-level) blocks that overlap the current selection. */
  function selectedBlocks() {
    var range = liveRange();
    if (!range) return [];
    var root = rootOf(range.startContainer);
    var blocks = [];
    var seen = new Set();
    function add(node) {
      var b = WW.commands ? WW.commands.blockFor(node, root) : null;
      if (b && !seen.has(b)) { seen.add(b); blocks.push(b); }
    }
    add(range.startContainer);
    add(range.endContainer);
    var nodes = textNodesInRange(range);
    nodes.forEach(function (n) { add(n); });
    // Element contents selected wholesale (Ctrl+A)
    if (blocks.length === 0) {
      var cands = range.commonAncestorContainer.querySelectorAll ?
        range.commonAncestorContainer.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,td,th,blockquote,pre,div') : [];
      Array.prototype.forEach.call(cands, function (el) { add(el); });
      if (blocks.length === 0) add(range.commonAncestorContainer);
    }
    return blocks;
  }

  /** Put the caret at the very beginning of the first block in `root`. */
  function focusStart(root) {
    var first = root.firstElementChild || root;
    root.focus();
    caretAtStartOf(first);
  }

  function scrollCaretIntoView() {
    var sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    var rect = sel.getRangeAt(0).getBoundingClientRect();
    var vp = document.getElementById('viewport');
    if (!vp || (rect.top === 0 && rect.bottom === 0 && rect.height === 0)) return;
    var pad = 60;
    if (rect.top < pad) vp.scrollTop -= (pad - rect.top);
    else if (rect.bottom > vp.clientHeight - pad) vp.scrollTop += (rect.bottom - (vp.clientHeight - pad));
  }

  return {
    capture: capture, restore: restore, liveRange: liveRange, isCollapsed: isCollapsed,
    pathFromRoot: pathFromRoot, nodeFromPath: nodeFromPath, rootOf: rootOf, areaOf: areaOf,
    editableByArea: editableByArea, selectNode: selectNode, caretAtStartOf: caretAtStartOf,
    caretAfter: caretAfter, textNodesInRange: textNodesInRange, selectedBlocks: selectedBlocks,
    focusStart: focusStart, scrollCaretIntoView: scrollCaretIntoView, isBlock: isBlock, closest: closest
  };
})(WW.util);
