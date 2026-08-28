(() => {
  'use strict';

  // Jeff shortcut map. Uses Control/Option/Shift on macOS, not Command.
  // Codes correspond to the previously ported Apps Script dispatcher.
  const SHORTCUTS = new Map([
    ['C+S+KeyN', '11'],
    ['C+S+KeyC', '12'],
    ['C+S+KeyV', '13'],
    ['C+S+KeyS', '17'],
    ['C+S+KeyD', '18'],
    ['C+A+KeyD', '19'],
    ['C+S+KeyZ', '21'],
    ['C+S+KeyR', '24'],
    ['C+A+KeyR', '25'],
    ['C+S+KeyQ', '26'],
    ['C+S+F4', '27'],
    ['C+A+F4', '28'],
    ['C+S+KeyE', '31'],
    ['C+A+S+KeyE', '32'],
    ['C+A+S+KeyC', '33'],
    ['C+A+S+KeyV', '34'],
    ['C+S+KeyW', '35'],
    ['C+S+KeyM', '36'],
    ['C+A+S+KeyM', '37'],
    ['C+S+KeyT', '38'],
    ['C+S+KeyI', '39'],
    ['C+A+S+KeyF', '41'],
    ['C+S+KeyB', '42'],
    ['C+A+S+KeyB', '43'],
    ['C+A+S+KeyW', '44'],
    ['C+A+S+KeyH', '45'],
    ['C+A+S+ArrowRight', '46'],
    ['C+A+S+ArrowLeft', '47'],
    ['C+A+S+ArrowDown', '48'],
    ['C+A+S+ArrowUp', '49'],
    ['C+A+KeyW', '51'],
    ['C+A+KeyL', '52'],
    ['C+A+KeyH', '53'],
    ['C+S+Comma', '54'],
    ['C+S+Period', '55'],
    ['C+A+S+KeyO', '59'],
    ['C+S+KeyK', 'ZIN'],
    ['C+S+KeyJ', 'ZOUT'],
    ['C+S+KeyH', 'Z100'],
    ['C+A+Digit1', 'F1TOGGLE'],
    ['C+A+S+F4', 'BLOCK']
  ]);

  let f1Blocked = true;
  let busy = false;
  let bridgePrefix = null;
  let bridgePrefixAt = 0;

  function comboFromEvent(e) {
    const parts = [];
    if (e.ctrlKey) parts.push('C');
    if (e.altKey) parts.push('A');
    if (e.shiftKey) parts.push('S');
    if (e.metaKey) parts.push('M');
    parts.push(e.code);
    return parts.join('+');
  }

  function inCellEditMode() {
    const el = document.activeElement;
    if (!el) return false;
    if (el.matches?.('input, textarea')) {
      // The name box is not cell-edit mode.
      if (el.classList?.contains('waffle-name-box')) return false;
      return true;
    }
    if (el.closest?.('.cell-input[role="textbox"]')) return true;
    if (el.isContentEditable && el.closest?.('.cell-input')) return true;
    return false;
  }

  function getSpreadsheetId() {
    const m = location.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
    return m ? m[1] : null;
  }

  function getSheetId() {
    const u = new URL(location.href);
    let gid = u.searchParams.get('gid');
    if (!gid) {
      const m = location.hash.match(/(?:^|[&#])gid=(\d+)/);
      gid = m ? m[1] : null;
    }
    return gid == null ? 0 : Number(gid);
  }

  function getA1() {
    const box = document.querySelector('.waffle-name-box');
    return box?.value?.trim() || null;
  }

  function context() {
    return {
      spreadsheetId: getSpreadsheetId(),
      sheetId: getSheetId(),
      a1: getA1()
    };
  }

  function toast(text, kind = 'info', ms = 2200) {
    let el = document.getElementById('__jeff_fast_toast__');
    if (!el) {
      el = document.createElement('div');
      el.id = '__jeff_fast_toast__';
      Object.assign(el.style, {
        position: 'fixed',
        left: '50%',
        bottom: '36px',
        transform: 'translateX(-50%)',
        zIndex: '2147483647',
        padding: '8px 12px',
        borderRadius: '6px',
        font: '12px -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',
        boxShadow: '0 2px 10px rgba(0,0,0,.25)',
        pointerEvents: 'none',
        maxWidth: '560px'
      });
      document.documentElement.appendChild(el);
    }
    el.style.background = kind === 'error' ? '#b3261e' : '#202124';
    el.style.color = '#fff';
    el.textContent = text;
    el.style.display = 'block';
    clearTimeout(el.__timer);
    el.__timer = setTimeout(() => { el.style.display = 'none'; }, ms);
  }

  function selectA1(a1) {
    if (!a1) return;
    const box = document.querySelector('.waffle-name-box');
    if (!box) return;
    box.focus();
    box.value = a1;
    box.dispatchEvent(new Event('input', {bubbles: true}));
    box.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
    }));
  }

  async function runCommand(command) {
    if (command === 'BLOCK') return;
    if (command === 'F1TOGGLE') {
      f1Blocked = !f1Blocked;
      toast('F1 Help: ' + (f1Blocked ? 'DISABLED' : 'ENABLED'));
      return;
    }

    const ctx = context();
    if (!ctx.spreadsheetId || !ctx.a1) {
      toast('Jeff: セルまたは範囲を選択してください。', 'error');
      return;
    }

    // Prevent accidental command piling. The API path is already much faster than Apps Script.
    if (busy) return;
    busy = true;
    try {
      const res = await chrome.runtime.sendMessage({type: 'JEFF_RUN', command, context: ctx});
      if (!res?.ok) {
        if (res?.code === 'AUTH_REQUIRED') {
          toast('Jeff: Chrome拡張アイコンをクリックし「Googleアカウントを承認」を一度実行してください。', 'error', 5000);
        } else {
          toast('Jeff: ' + (res?.error || '処理に失敗しました。'), 'error', 5000);
        }
        return;
      }
      if (res.toast) toast(res.toast);
      if (res.selectA1) setTimeout(() => selectA1(res.selectA1), 30);
    } catch (e) {
      toast('Jeff: ' + (e?.message || String(e)), 'error', 5000);
    } finally {
      busy = false;
    }
  }

  // Fast Hammerspoon bridge: Ctrl+Option+Command+digit, twice, encodes a 2-digit Jeff command.
  // Example: J11 = bridge 1 then 1. No Tool Finder and no deliberate delay.
  function maybeHandleBridge(e) {
    if (!(e.ctrlKey && e.altKey && e.metaKey) || e.shiftKey) return false;
    const m = /^Digit([1-9])$/.exec(e.code);
    if (!m) return false;
    e.preventDefault();
    e.stopImmediatePropagation();
    const digit = m[1];
    const now = performance.now();
    if (!bridgePrefix || now - bridgePrefixAt > 700) {
      bridgePrefix = digit;
      bridgePrefixAt = now;
      return true;
    }
    const code = bridgePrefix + digit;
    bridgePrefix = null;
    bridgePrefixAt = 0;
    runCommand(code);
    return true;
  }

  window.addEventListener('keydown', (e) => {
    if (maybeHandleBridge(e)) return;

    // Keep ordinary F1 disabled when enabled, matching the XLAM behavior.
    if (f1Blocked && e.code === 'F1' && !e.ctrlKey && !e.altKey && !e.shiftKey && !e.metaKey) {
      e.preventDefault();
      e.stopImmediatePropagation();
      return;
    }

    if (inCellEditMode()) return;
    const combo = comboFromEvent(e);
    const command = SHORTCUTS.get(combo);
    if (!command) return;

    e.preventDefault();
    e.stopImmediatePropagation();
    runCommand(command);
  }, true);
})();
