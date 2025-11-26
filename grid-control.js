(function () {

  /* ================================ JSS 共通UXユーティリティ ================================ */

  // 最小フォント(既存のものを流用)
  const SHRINK_MIN_PX = 9;
  const CLAMP_NO = 2;
  const LONG_MS = 270;
  const cleanCellText = (raw) => {
    if (raw == null) return '';
    let s = String(raw);
    s = s.replace(/\r/g,'')
        .replace(/\u00A0/g,' ')
        .replace(/[\u200B-\u200D\u2060\uFEFF]/g,'')
        .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g,'');
    return s;
  };

  function scrollColumnToLeft(container, colIndex, sheet){
    const holder = container.querySelector('.jexcel_content'); if(!holder) return;
    let offset = 0; const cols = sheet.options.columns; const fixed = sheet.options.freezeColumns || 0;
    const target = Math.max(colIndex, fixed);
    for(let c=fixed; c<target; c++){ offset += (cols[c]?.width || 100); }
    holder.scrollLeft = Math.max(0, offset-4);
  }

  function scrollRowToTop(container, rowIndex, jss, { position = 'top' } = {}) {
    const holder = container.querySelector('.jexcel_content');
    if (!holder) return;
    try {
      const td = jss.getCellFromCoords(Math.min(jss.options.freezeColumns || 0, 0), rowIndex);
      if (!td) return;
      const cellRect = td.getBoundingClientRect();
      const holderRect = holder.getBoundingClientRect();
      const delta = cellRect.top - holderRect.top;
      if (position === 'center') {
        holder.scrollTop += delta - holder.clientHeight / 2;
      } else if (position === 'bottom') {
        holder.scrollTop += delta - holder.clientHeight + td.offsetHeight;
      } else {
        holder.scrollTop += delta - 2;
      }
    } catch {}
  }

// ローカル日付で 'YYYY-MM-DD' を作る（JST安全）
function __todayYmdLocal() { return new Date().toLocaleDateString('sv-SE', { timeZone: 'Asia/Tokyo' }); }

/**
 * 初期表示：対象月(ym)が今月なら「今日」の行まで自動スクロール（縦）
 * - 日付は dateColIndex（既定0列目）に入っている前提
 * - AM/PMで1日が2行ある場合は最初に見つかった方へスクロール
 */
  function _autoScrollToToday(jss, container, {
    ym = null,              // 'YYYY-MM'（今月判定したい時だけ渡す）
    dateColIndex = 0,       // 日付が入っている列（A列=0）
    position = 'center'        // 'top' | 'center' | 'bottom'
  } = {}) {
    if (!jss || jss._autoScrollTodayDone) return;
    jss._autoScrollTodayDone = true;

    const todayYmd = __todayYmdLocal();
    const todayYm  = todayYmd.slice(0, 7);
    if (ym && ym !== todayYm) return; // 今月じゃないなら何もしない

    // グリッドの行数を取得
    const totalRows = Array.isArray(jss?.options?.data)
      ? jss.options.data.length
      : (Number(jss?.options?.rows) || 0);

    if (!totalRows) return;

    // 日付列を走査して "YYYY-MM-DD" が一致する最初の行を探す
    let targetRow = -1;
    for (let y = 0; y < totalRows; y++) {
      let v = '';
      try {
        v = jss.getValueFromCoords(dateColIndex, y) ?? '';
      } catch {}
      // セル表示が 'M/D' などでも、toSchedule で 'YYYY-MM-DD' を入れているならここで一致
      // もし 'M/D' 表記なら ym と組み合わせて比較
      const s = String(v).trim();
      if (s === todayYmd) { targetRow = y; break; }

      if (/^(\d{1,2})[\/.](\d{1,2})$/.test(s) && ym) {
        const [, m, d] = s.match(/^(\d{1,2})[\/.](\d{1,2})$/);
        const mm = String(m).padStart(2, '0');
        const dd = String(d).padStart(2, '0');
        if (todayYmd === `${ym}-${dd}` && ym.slice(5,7) === mm) {
          targetRow = y; break;
        }
      }
    }

    if (targetRow < 0) return;

    // レイアウト後に縦スクロール
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        try { scrollRowToTop(container, targetRow, jss, { position }); } catch {}
      });
    });
  }

  // 計測用 span(既存のものを流用／なければ作成)
  let __fitMeasureSpan = null;
  function getFitMeasureSpan() {
    if (!__fitMeasureSpan) {
      const s = document.createElement('span');
      s.style.position = 'fixed';
      s.style.left = '-99999px';
      s.style.top = '-99999px';
      s.style.whiteSpace = 'nowrap';
      s.style.visibility = 'hidden';
      document.body.appendChild(s);
      __fitMeasureSpan = s;
    }
    return __fitMeasureSpan;
  }

  // テキストをセル幅に収める(既存の applyShrinkToFit を汎用化)
  function applyShrinkToFit(td, minPx = SHRINK_MIN_PX, clamp= CLAMP_NO) {
    if (!td) return;
    if (td.querySelector('input, textarea, [contenteditable="true"]')) return;

    const text = (td.textContent || '').trim();
    if (!text) {
      td.style.fontSize = '';
      // clamp系もクリア
      td.style.whiteSpace = '';
      td.style.overflow = '';
      td.style.textOverflow = '';
      td.style.display = '';
      td.dataset.fitApplied = '0';
      td.dataset.fitKey = '';
      return;
    }

    const cs = getComputedStyle(td);
    const padL = parseFloat(cs.paddingLeft) || 0;
    const padR = parseFloat(cs.paddingRight) || 0;
    const avail = Math.max(0, td.clientWidth - padL - padR - 2);
    if (avail <= 0) return;

    const key = text + '|' + avail + '|' + cs.fontFamily + '|' + cs.fontWeight;
    let didShrink = false;
    if (td.dataset.fitApplied === '1' && td.dataset.fitKey === key) return;

    const measure = getFitMeasureSpan();
    measure.style.fontFamily = cs.fontFamily;
    measure.style.fontWeight = cs.fontWeight;
    measure.style.letterSpacing = cs.letterSpacing;
    measure.style.fontSize = cs.fontSize;
    measure.textContent = text;

    const basePx = parseFloat(cs.fontSize) || 14;
    let w = measure.offsetWidth;

    if (w <= avail) {
      td.style.fontSize = '';
      td.dataset.fitApplied = '0';
      td.dataset.fitKey = key;
      didShrink = true;
    } else {
      let newSize = Math.max(minPx, Math.floor(basePx * (avail / w)));
      measure.style.fontSize = newSize + 'px';
      for (let i = 0; i < 3; i++) {
        if (measure.offsetWidth > avail && newSize > minPx) {
          newSize -= 1;
          measure.style.fontSize = newSize + 'px';
        } else {
          break;
        }
      }
      td.style.fontSize = newSize + 'px';
      td.dataset.fitApplied = '1';
      td.dataset.fitKey = key;
      didShrink = true;
    }
    // ここで必ず1行/複数行(行数制限)の最終スタイルを上書き適用
    if (clamp && clamp > 0) {
      td.style.whiteSpace = 'normal';
      td.style.wordBreak = 'break-word';
      td.style.overflow = 'hidden';
      td.style.textOverflow = '';
      td.style.display = '';
    } else {
      td.style.whiteSpace = 'nowrap';
      td.style.overflow = 'hidden';
      td.style.textOverflow = 'ellipsis';
      td.style.display = ''; // ← 重要: unclamp時に display を戻す
    }
  }
  // 編集開始・終了フック(確定後に縮小適用)
  function installShrinkHooksForJSS(jss, container, { minPx = SHRINK_MIN_PX, dataColStart = 0 } = {}) {
    const prevStart = jss.options.oneditionstart;
    jss.options.oneditionstart = function (el, cell, x, y) {
      // ★ 読み取り専用やデータ列外はフォント解除しない
      if (cell && !cell.classList.contains('readonly') && (x == null || x >= dataColStart)) {
        cell.style.fontSize = '';
        cell.dataset.fitApplied = '0';
        const ed = cell.querySelector('input, textarea, [contenteditable="true"]');
        if (ed && ed.focus) { try { ed.focus(); } catch {} }
      }
      
      // ★ エディタDOMの blur 時にも縮小復帰(保存確定しなくても戻す)
      setTimeout(() => {
        const ed = cell?.querySelector?.('textarea, input, [contenteditable="true"]') ||
                  container.querySelector('.jexcel_editor textarea, .jexcel_editor input, .jexcel_editor [contenteditable="true"]');
        if (!ed) return;
        const handler = () => {
          const td = jss.getCellFromCoords(x, y);
          if (td && x >= dataColStart) requestAnimationFrame(() => applyShrinkToFit(td, SHRINK_MIN_PX));
        };
        ed.addEventListener('blur', handler, { once: true });
      }, 0);
      if (typeof prevStart === 'function') { try { prevStart(el, cell, x, y); } catch {} }
    };

    const prevEnd = jss.options.oneditionend;
    jss.options.oneditionend = function (el, info) {
        if (typeof prevEnd === 'function') { try { prevEnd(el, info); } catch {} }
        const x = Number(info?.x), y = Number(info?.y);
        if (!Number.isFinite(x) || !Number.isFinite(y)) return;
        if (x < dataColStart) return;
        const td = jss.getCellFromCoords(x, y);
        // ★ 読み取り専用セルならエディタを開かない
        if (td && td.classList.contains('readonly')) return;
        if (!td) return;
        requestAnimationFrame(() => applyShrinkToFit(td, minPx));
    };
  
    // ★ 選択変更(クリックで別セルへ移動)でも、直前の編集中セルを閉じて縮小を戻す
    const prevSel = jss.options.onselection;
    jss.options.onselection = function(el, x1, y1, x2, y2){
      const ed = jss._editing;
      if (ed && ed.x != null && ed.y != null) {
        const tdPrev = jss.getCellFromCoords(ed.x, ed.y);
        // 保存確定しないままでも閉じる
        try { jss.closeEditor(tdPrev, false); } catch {}
        if (tdPrev && ed.x >= dataColStart) requestAnimationFrame(() => applyShrinkToFit(tdPrev, minPx));
      }
      if (typeof prevSel === 'function') { try { prevSel(el, x1, y1, x2, y2); } catch {} }
    };
  }

  // ちょうど1列ぶんだけ横スクロールして、反映後に resolve
  function ensureVisibleColumnOneStep(container, jss, fromX, toX) {
    return new Promise((resolve) => {
      const holder = container.querySelector('.jexcel_content');
      if (!holder) return resolve(false);

      const fixed = jss.options.freezeColumns || 0;
      const cols  = jss.options.columns || [];
      const colW  = (i) => (cols[i]?.width || 100);

      // 凍結列はスクロールしない
      if (toX < fixed && fromX < fixed) return resolve(false);

      const start = holder.scrollLeft;
      let delta = 0;

      // 右へ1列：toX 側の幅ぶん、左へ1列：fromX 側の幅ぶん戻す
      if (toX > fromX) {
        delta = colW(toX);
      } else if (toX < fromX) {
        delta = -colW(fromX);
      } else {
        return resolve(false);
      }

      const target = Math.max(0, start + delta);

      if (target === start) return resolve(false);

      holder.scrollLeft = target;

      // スクロール反映を次々フレームまで待つ(フォールバックあり)
      const waitApplied = () => {
        if (holder.scrollLeft !== start) {
          requestAnimationFrame(() => requestAnimationFrame(() => resolve(true)));
        } else {
          setTimeout(() => resolve(true), 32);
        }
      };
      requestAnimationFrame(waitApplied);
    });
  }

  // Enter 即編集 + Delete/Backspace 一括クリア → 即編集（統合版）
  function installEditorEnterOverride(jss, container, {
    lastEditableCol = null,   // 省略時: columns.length - 1
    dataColStart = 0,         // スキップ列などの起点
  } = {}) {
    if (jss._enterOverrideInstalled) return;
    jss._enterOverrideInstalled = true;
    jss._uxDataColStart = dataColStart; // 他フックからも参照できるよう保持

    const getLastCol = () => {
      const cols = jss?.options?.columns?.length ?? 1;
      return (lastEditableCol == null) ? (cols - 1) : lastEditableCol;
    };
    const getLastRow = () => {
      const rows = Array.isArray(jss?.options?.data) ? jss.options.data.length : (Number(jss?.options?.rows) || 1000);
      return Math.max(0, rows - 1);
    };
    const moveAndEdit = (coords, mode = 'down') => {
      try {
        let x, y;
        if (coords && coords.x != null) { x = Number(coords.x); y = Number(coords.y); }
        else if (jss._editing && jss._editing.x != null) { x = Number(jss._editing.x); y = Number(jss._editing.y); }
        else {
          const sel = (typeof jss.getSelected === 'function') ? jss.getSelected() : null;
          x = sel ? Number(sel[0]) : 0; y = sel ? Number(sel[1]) : 0;
        }
        let nx = x, ny = y;
        if (mode === 'right') {
          const lastCol = getLastCol();
          if (x >= lastCol) return;
          nx = x + 1; ny = y;
        } else {
          const lastRow = getLastRow();
          if (y >= lastRow) return;
          nx = x; ny = y + 1;
        }

        const open = () => {
          if (jss._openingNext) return;
          jss._openingNext = { x: nx, y: ny };
          requestAnimationFrame(() => {
            requestAnimationFrame(() => {
              const tag = jss._openingNext; jss._openingNext = null;
              if (!tag) return;
              const td = jss.getCellFromCoords(tag.x, tag.y);
              if (td) { try { jss.openEditor(td); } catch {} }
            });
          });
        };

        if (mode === 'right') {
          ensureVisibleColumnOneStep(container, jss, x, nx).then(open);
        } else {
          try {
            const tdNext = jss.getCellFromCoords(nx, ny);
            tdNext?.scrollIntoView?.({ block: 'nearest' });
          } catch {}
          open();
        }
      } catch {}
    };

    // -------- Enter 差し替え（セルエディタ内）
    const origStart = jss.options.oneditionstart;
    jss.options.oneditionstart = function(el, cell, x, y){
      jss._editing = { x, y, cell };
      setTimeout(() => {
        const ed = cell?.querySelector?.('textarea, input, [contenteditable="true"]') ||
                  container.querySelector('.jexcel_editor textarea') ||
                  container.querySelector('.jexcel_editor input') ||
                  container.querySelector('.jexcel_editor [contenteditable="true"]');
        if (!ed) return;

        ed._enterHandlerBound && ed.removeEventListener('keydown', ed._enterHandlerBound);
        const handler = (ev) => {
          if (ev.key === 'Enter' && !ev.altKey) {
            ev.preventDefault(); ev.stopPropagation();
            try { applyShrinkToFit && applyShrinkToFit(cell); } catch {}
            try { jss.closeEditor && jss.closeEditor(cell || null, true); } catch {}
            try { jss.closeEditor && jss.closeEditor(); } catch {}
            const mode = jss._enterBehavior || 'down';
            moveAndEdit({ x, y }, mode);
          }
        };
        ed.addEventListener('keydown', handler);
        ed._enterHandlerBound = handler;
      }, 0);
      if (typeof origStart === 'function') try { origStart.call(this, el, cell, x, y); } catch {}
    };

    const origEnd = jss.options.oneditionend;
    jss.options.oneditionend = function(el, cell, x, y, save){
      try { applyShrinkToFit && applyShrinkToFit(cell); } catch {}
      setTimeout(() => { jss._editing = null; }, 0);
      if (typeof origEnd === 'function') try { origEnd.call(this, el, cell, x, y, save); } catch {}
    };

    // -------- Delete / Backspace：非編集中＆選択中 → 一括クリア → 即編集（強制適用）
    (function bindDeleteOpenToEdit() {
      if (jss._deleteOpenInstalled) return;
      jss._deleteOpenInstalled = true;

      const keyHandler = (ev) => {
        if (ev.isComposing) return;
        if (jss._editing) return;

        const isDel = ev.key === 'Delete';
        const isBS  = ev.key === 'Backspace' && !ev.metaKey && !ev.ctrlKey && !ev.altKey && !ev.shiftKey;
        if (!isDel && !isBS) return;

        const t = ev.target;
        if (t && (t.isContentEditable || /^(input|textarea|select)$/i.test(t.tagName))) return;
        if (!container.contains(document.activeElement)) return;

        // アンカーセル（最優先は getSelected）
        let r = jss._lastSel;
        if(!r) return;

        let td = null;
        try { td = jss.getCellFromCoords(r.x1, r.y1); } catch {}
        if (!td || td.classList.contains('readonly')) return;

        // 履歴戻り対策や他ハンドラ先行を防ぐため capture で抑止
        ev.preventDefault(); ev.stopPropagation(); ev.stopImmediatePropagation?.();

        for (let y = r.y1; y <= r.y2; y++) {
          for (let x = r.x1; x <= r.x2; x++) {
            try {
              const td = jss.getCellFromCoords(x, y);
              if (!td || td.classList.contains('readonly')) continue;
              jss.setValueFromCoords(x, y, '');
              // 見た目の整合（必要なら）
              try { applyShrinkToFit && applyShrinkToFit(td); } catch {}
            } catch {}
          }
        }
        try { jss.openEditor(td); } catch {}
      };

      // Backspace のナビゲーション防止も兼ねて capture で
      document.addEventListener('keydown', keyHandler, true);
    })();
  }

  // クリック短押し＝単一点なら openEditor、長押し・ドラッグは編集しない
  function installOpenEditorOnClick(jss, container, { minEditableCol = 0 } = {}) {
    const root = container.querySelector('.jexcel');
    if (!root || jss._clickOpenInstalled) return;
    jss._clickOpenInstalled = true;

    const MOVE_PX = 4;
    let downX=0, downY=0, dragging=false, longTimer=null, longActive=false, rightPress=false;

    const clearLong = ()=>{ if (longTimer) { clearTimeout(longTimer); longTimer=null; } };

    const tdToCoords = (td) => {
        if (!td) return null;
        if (typeof jss.getCoords === 'function') {
        try { const [x,y] = jss.getCoords(td); return { x:Number(x), y:Number(y) }; } catch {}
        }
        const ax = td.getAttribute('x') ?? td.getAttribute('data-x') ?? td.dataset?.x;
        const ay = td.getAttribute('y') ?? td.getAttribute('data-y') ?? td.dataset?.y;
        if (ax != null && ay != null) return { x:Number(ax), y:Number(ay) };
        return null;
    };

    const selectOne = (y, x) => {
        if (typeof jss.setHighlighted === 'function') jss.setHighlighted(y, x);
        else if (typeof jss.setSelection === 'function') jss.setSelection(y, x, y, x);
        else if (typeof jss.setSelectionFromCoords === 'function') jss.setSelectionFromCoords(x, y, x, y);
        jss._selActive = true; jss._lastSel = { x1:x, y1:y, x2:x, y2:y };
    };

    root.addEventListener('mousedown', (e)=>{
        rightPress = (e.button===2) || (e.ctrlKey && e.button===0);
        dragging=false; downX=e.clientX; downY=e.clientY; longActive=false;
        if (!rightPress) { try { jss.closeEditor && jss.closeEditor(); } catch {} }

        const td = e.target.closest?.('td');
        const c  = tdToCoords(td);
        clearLong();
        longTimer = setTimeout(()=>{
        if (!dragging && c){ longActive = true; selectOne(c.y, c.x); }
        }, LONG_MS);
    }, true);

    root.addEventListener('mousemove', (e)=>{
      if (!dragging && (Math.abs(e.clientX-downX)>MOVE_PX || Math.abs(e.clientY-downY)>MOVE_PX)){
        dragging = true; clearLong();
      }
    }, true);

    root.addEventListener('mouseup', (e)=>{
      clearLong();
      if (dragging) return;
      if (rightPress){ rightPress=false; const td = e.target.closest?.('td'); const c = tdToCoords(td); if (c) selectOne(c.y, c.x); return; }
      if (longActive){ longActive=false; return; }

      const s = jss._lastSel; if (!s) return;
      const { x1,y1,x2,y2 } = s;
      if (x1===x2 && y1===y2 && x1>=minEditableCol){
      const td = jss.getCellFromCoords(x1,y1);
      if (!td) return;

      // ★ 読み取り専用セルはエディタを開かない＆縮小を即復帰
      if (td.classList.contains('readonly')) {
        requestAnimationFrame(() => applyShrinkToFit(td, SHRINK_MIN_PX));
        return;
      }
      const ed = jss._editing, pending = jss._openingNext;
      const sameEditing = ed && Number(ed.x)===x1 && Number(ed.y)===y1;
      const samePending = pending && Number(pending.x)===x1 && Number(pending.y)===y1;
      if (!sameEditing && !samePending && td) try { jss.openEditor(td); } catch {}
      }
    }, true);
  }

  // クリップボード：空範囲コピー対策(OSが前回内容を残すのを防ぐ)
  function installCopyEmptyGuard(jss, container){
    const root = container.querySelector('.jexcel_content');
    if (!root || jss._copyEmptyHooked) return;
    jss._copyEmptyHooked = true;

    function getSelRect(){
      if (typeof jss.getSelected === 'function'){
        const sel = jss.getSelected();
        if (Array.isArray(sel)){
          if (sel.length>=4) return { x1:+sel[0], y1:+sel[1], x2:+sel[2], y2:+sel[3] };
          if (sel.length>=2) return { x1:+sel[0], y1:+sel[1], x2:+sel[0], y2:+sel[1] };
        }
      }
      const s = jss._lastSel; if (s) return { x1:+s.x1, y1:+s.y1, x2:+s.x2, y2:+s.y2 };
      return null;
    }

    root.addEventListener('copy', (e)=>{
      const r = getSelRect(); if (!r) return;
      const { x1,y1,x2,y2 } = r;
      let allEmpty = true;
      outer: for (let y=Math.min(y1,y2); y<=Math.max(y1,y2); y++){
        for (let x=Math.min(x1,x2); x<=Math.max(x1,x2); x++){
          let v=''; try { v = jss.getValueFromCoords(x,y) ?? ''; } catch {}
          const cleaned = String(v).replace(/\r/g,'').replace(/\u00A0/g,' ')
          .replace(/[\u200B-\u200D\u2060\uFEFF]/g,'')
          .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g,'');
          if (cleaned.trim()!==''){ allEmpty=false; break outer; }
        }
      }
      if (!allEmpty) return;
      e.preventDefault();
      const cd = e.clipboardData || window.clipboardData;
      if (cd && cd.setData) cd.setData('text/plain','\u200B');
    }, true);
  }

  // 貼り付け：選択起点からグリッドに展開(\n→行、\t→列)
  function installPasteInterceptor(jss, container) {
    const root = container.querySelector('.jexcel_content');
    if (!root || jss._pasteHooked) return;
    jss._pasteHooked = true;

    root.addEventListener('paste', (e)=>{
      const t = e.target; if (!container.contains(t)) return;
      const cd = e.clipboardData || window.clipboardData; if (!cd) return;
      const text = cd.getData('text'); if (!text) return;
      e.preventDefault();

      let startX, startY;
      if (jss._editing && jss._editing.x != null){ startX = Number(jss._editing.x); startY = Number(jss._editing.y); }
      else if (typeof jss.getSelected === 'function'){ const sel = jss.getSelected(); startX = Number(sel?.[0] ?? 0); startY = Number(sel?.[1] ?? 0); }
      else { const last = jss._lastSel; startX = Number(last?.x1 ?? 0); startY = Number(last?.y1 ?? 0); }

      const rows = text.replace(/\r/g,'').split('\n').map(r => r.split('\t').map(cleanCellText));
      let y = startY;
      for (const row of rows){
        if (row.length===1 && row[0]==='') continue;
        let x = startX;
        for (const v of row){
          const val = (v.replace(/\s+/g,' ').trim()==='') ? '' : v;
          try { jss.setValueFromCoords(x, y, val); } catch {}
          x++;
        }
        y++;
      }
    }, true);
  }

  // 追記：全セルに一度だけ縮小適用(初期表示用)
  function __applyInitialShrinkAll(jss, container, dataColStart) {
    const table = container.querySelector('.jexcel_content table');
    if (!table) return;
    // x>=dataColStart のセルだけ対象にする
    const tds = table.querySelectorAll('td');
    tds.forEach(td => {
      const ax = td.getAttribute('x') ?? td.getAttribute('data-x') ?? td.dataset?.x;
      const x  = ax != null ? Number(ax) : null;
      if (x != null && x >= dataColStart) {
        applyShrinkToFit(td, SHRINK_MIN_PX);
      }
    });
  }

  // 追記：onselection が未設定のシートに _lastSel を仕込む
  function __ensureSelectionTracker(jss) {
    if (jss._selectionTrackerInstalled) return;
    jss._selectionTrackerInstalled = true;

    const prev = jss.options.onselection;
    jss.options.onselection = function(el, x1, y1, x2, y2) {
      // 旧ハンドラを先に（または後に）呼ぶかは要件次第
      if (typeof prev === 'function') {
        try { prev.call(this, el, x1, y1, x2, y2); }
        catch (e) { /* console.warn('onselection error:', e); */ }
      }
      jss._selActive = true;
      jss._lastSel = { x1, y1, x2, y2 };
    };
  }

  // === クリーニング/ダーティ/休日ハイライト ===
  function _attachOnChangeCleaner(jss) {
    if (!jss || jss._onChangeCleanerInstalled) return;  // 二重適用防止
    const prev = jss.options?.onchange;

    // 既定設定：呼び出し元が set している dataColStart を尊重（未設定なら 1）
    const dataColStart = (jss._uxDataColStart != null) ? jss._uxDataColStart : 1;
    const holidayClass = 'ht-holiday';
    const outClass = 'ht-out';
    const cleanFn      = cleanCellText;

    jss._inCleanChange = false;
    jss._onChangeCleanerInstalled = true;

    jss.options.onchange = function (el, cell, x, y, newValue) {
      // 1) 値のクリーニング
      const cleaned = (typeof cleanFn === 'function') ? cleanFn(newValue) : String(newValue ?? '');
      // 2) クリーニング結果が異なるなら次フレームで差し替え（再入防止）
      if (!jss._inCleanChange && cleaned !== newValue) {
        jss._inCleanChange = true;
        requestAnimationFrame(() => {
          try { jss.setValueFromCoords(x, y, cleaned); } finally { jss._inCleanChange = false; }
        });
      }
      // 3) ダーティ管理（データ列のみ／readonly除外）
      if (x >= dataColStart) {
        try {
          const td = jss.getCellFromCoords(x, y);
          if (td && !td.classList.contains('readonly')) {
            // グローバルの markDirty があれば “件数管理＋見た目” を委譲
            if (typeof window.markDirty === 'function') {
              window.markDirty(y, x, true);            // ★ ここがポイント
            } else {
              // フォールバック
              td.classList.add('dirty-cell');
              if (typeof window.updateDirtyBadge === 'function') {
                window.updateDirtyBadge();
              }
            }
          }
        } catch {}
      }
      // 4) “休” ハイライト
      try {
        const td = jss.getCellFromCoords(x, y);
        if (td) {
          const s = String(newValue ?? '');
          if (s.includes('休')) td.classList.add(holidayClass);
          else td.classList.remove(holidayClass);
        }
        if (td) {
          const s = String(newValue ?? '');
          if (s.includes('講習') || s.includes('研修') || s.includes('健康診断')) td.classList.add(outClass);
          else td.classList.remove(outClass);
        }
      } catch {}
      // 5) 元の onchange も呼ぶ（最後に）
      if (typeof prev === 'function') {
        try { prev(el, cell, x, y, newValue); } catch {}
      }
    };
  }

  function installCommonJSSUX(jss, container, {
    // text fitting
    minPx = SHRINK_MIN_PX,
    dataColStart = 0,
    enableTextFit = true,
    enableInitialTextFit = true,
    // clipboard
    allowPaste = true,
    allowCopyEmptyGuard = true,
    // enter key behavior
    enterBehavior = 'down', // 'right' | 'down' | 'none'
    lastEditableCol = null, // ← これを追加
    // other helpers
    enableSelectionTracker = true,
    enableClickOpenEditor = true,
    onChangeCleaner = false,
    // ▼ 追加：初期表示で今日まで横スクロール
    autoScrollToToday = false,
    tym = null,                 // 'YYYY-MM' を渡すと「今月のときのみ」発火
  } = {}){
    if (enableSelectionTracker) __ensureSelectionTracker(jss);
    if (enableTextFit) installShrinkHooksForJSS(jss, container, { minPx, dataColStart });
    jss._enterBehavior = enterBehavior;
    if (enterBehavior !== 'none') {
      installEditorEnterOverride(jss, container, { lastEditableCol, dataColStart });
    }
    if (enableClickOpenEditor) installOpenEditorOnClick(jss, container, { minEditableCol: dataColStart });
    if (allowCopyEmptyGuard) installCopyEmptyGuard(jss, container);
    if (allowPaste) installPasteInterceptor(jss, container);

    if (enableInitialTextFit && enableTextFit) {
      requestAnimationFrame(() => __applyInitialShrinkAll(jss, container, dataColStart));
    }
    if (onChangeCleaner) {_attachOnChangeCleaner(jss);}    
    if (autoScrollToToday) {_autoScrollToToday(jss, container, tym);}
  };

  window.GridControl = { 
    SHRINK_MIN_PX,
    installCommonJSSUX, 
    scrollColumnToLeft,
    scrollRowToTop 
  };
})();
