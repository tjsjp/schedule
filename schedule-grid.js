(function () {
  'use strict';

  function toExcelAddr(colIndex, rowIndex) {
    // 0-based col/row -> Excel style (A1, A2, ...)
    let n = colIndex + 1;
    let s = '';
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s + String(rowIndex + 1);
  }
  // Build grid data/merges from a simple schedule list
  // schedule item format supports either:
  // - { date: 'YYYY-MM-DD', slot: 'AM'|'PM', values: [per-employee values] }
  // - { date: 'YYYY-MM-DD', slot: 'AM'|'PM', byName: { [employeeName]: value } }
  function buildGridDataFromSchedule(meta, schedule) {
    const employees = meta.employees || [];
    const labels = meta.labels || { am: 'AM', pm: 'PM', memo: '' };
    const data = [];
    const merges = {};

    // Row 0: memo row
    data.push(['', labels.memo || '', ...Array(employees.length).fill('')]);

    if (!Array.isArray(schedule) || schedule.length === 0) {
      return { data, merges };
    }

    // Group by date preserving order
    const order = [];
    const byDate = new Map();
    for (const item of schedule) {
      const d = String(item?.date || '').trim();
      if (!d) continue;
      if (!byDate.has(d)) { byDate.set(d, { AM: null, PM: null }); order.push(d); }
      const slotRaw = String(item?.slot || '').toUpperCase();
      const slot = slotRaw === 'PM' ? 'PM' : 'AM';
      let rowValues = null;
      if (Array.isArray(item?.values)) {
        rowValues = Array.from({ length: employees.length }, (_, i) => item.values[i] ?? '');
      } else if (item?.byName && typeof item.byName === 'object') {
        rowValues = employees.map(name => item.byName[name] ?? '');
      } else {
        rowValues = Array(employees.length).fill('');
      }
      byDate.get(d)[slot] = rowValues;
    }

    let r = 1; // start after memo row
    for (const d of order) {
      const rows = byDate.get(d) || { AM: null, PM: null };
      const amRow = rows.AM || Array(employees.length).fill('');
      const pmRow = rows.PM || Array(employees.length).fill('');
      data[r] = [d, labels.am || 'AM', ...amRow];
      data[r + 1] = [d, labels.pm || 'PM', ...pmRow];
      const addr = toExcelAddr(0, r); // column A, current date start
      merges[addr] = [1, 2];
      r += 2;
    }

    return { data, merges };
  }

  function getYmdForRow(data, y) {
    for (let i = y; i >= 0; i--) {
      const raw = String(data?.[i]?.[0] ?? '').trim(); // A列
      if (!raw) continue;
      if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw; // 既に YYYY-MM-DD
    }
    return '';
  }
  
  function init(elOrSelector, opts, ux) {
    var host = typeof elOrSelector === 'string'
      ? document.querySelector(elOrSelector)
      : elOrSelector;
    if (!host) throw new Error('ScheduleGrid: target element not found');
    host.classList.toggle('is-mobile', !!opts?.isProbablyMobile);
    if (typeof window.jspreadsheet !== 'function') {
      throw new Error('ScheduleGrid: jspreadsheet is not loaded');
    }
    host.innerHTML = '';
    // Options (ASCII labels). Replace to Japanese after encoding is sorted.
    const meta = Object.assign(
      {
        employees: [],
        dates: [],
        // labels
        labels: {
          date: 'Date',
          slot: 'Slot',
          am: 'AM',
          pm: 'PM',
          topLeft: '',
          group: 'Name',
          memo: ''
        }
      },
      (opts && opts.meta) || {}
    );
    // === 追加: 自分列・祝日セット ===
    const myIndex  = Number(opts?.myIndex);
    const isHolidayLocal = (ymd) => {
      try { return (typeof isHoliday === 'function') ? !!isHoliday(ymd) : false; } catch { return false; }
    };
    // Columns: 2 frozen meta columns + employees
    const columns = [
      { title: meta.labels.date, type: 'text', readOnly: true, width:(opts?.isProbablyMobile ? 30 : 35)},
      { title: meta.labels.slot, type: 'text', readOnly: true, width: 35},
      ...meta.employees.map((name) => ({
        title: name,
        type: 'text',
        width: (opts?.isProbablyMobile ? 80 : 150),
      }))
    ];
    // Data rows & merges: accept direct data/merges or map from schedule list; fallback to blank
    let data = [];
    let merges = {};
    const schedule = opts && Array.isArray(opts.schedule) ? opts.schedule : null;
    const memoByName = opts && opts.memoByName ? opts.memoByName: null;
    if (schedule) {
      const built = buildGridDataFromSchedule(meta, schedule);
      data = built.data;
      merges = built.merges;
    }
    if (memoByName) {
      const names = meta.employees || [];
      const row0 = data[0];
      for (let i = 0; i < names.length; i++) {
        const v = opts.memoByName[names[i]];
        if (v != null) row0[2 + i] = String(v);
      }
    }
    const grid = window.jspreadsheet(host, {
      data,
      columns,
      freezeColumns: 2,          // keep Date/Slot frozen on the left
      mergeCells: merges,        // merge Date cells per day (2 rows)
      tableOverflow: true,
      tableHeight: '70vh',
      tableWidth: '100%',
      allowInsertRow: false,
      allowDeleteRow: false,
      allowInsertColumn: false,
      allowDeleteColumn: false,
      allowRenameColumn: false,
      allowMoveColumn: false,
      allowExport: false,
      columnSorting: false,
      about: false,
      text: {
        copy: 'コピー',
        paste: '貼り付け'
      },
      updateTable: (el, cell, x, y, source, value) => {
        if (y === 0) {
          cell.classList.add('readonly');
          cell.classList.add('memo-row');
          return;
        }

        try {
          // 前回の痕跡を掃除
          const dynPrev = (cell.dataset.htYmdClasses || '').split(/\s+/).filter(Boolean);
          dynPrev.forEach(c => cell.classList.remove(c));
          cell.classList.remove('ht-today','ht-sat','ht-sun','ht-holiday');

          // ScheduleGrid では A列が日付（行＝日付、列＝社員）
          const raw = String((data?.[y]?.[0] ?? '')).trim(); // A列：日付
          const today = new Date().toLocaleDateString('sv-SE', { timeZone: 'Asia/Tokyo' });

          const ymd = getYmdForRow(data, y);

          // A列の「表示だけ」をMM-DDにしたい場合
          if (x === 0) {
            const m = String(ymd).match(/^\d{4}-(\d{2})-(\d{2})$/);
            if (m) cell.textContent = `${m[2]}`;
          }

          // 日付ハイライト（当日/土日/休日、「休」）
          if (x >= 2 && ymd) {
          // PC版の classesByDate 相当（weekday判定＋祝日判定）
            const [Y,M,D] = ymd.split('-').map(Number);
            const wd = (function weekday(y,m,d){ const t=[0,3,2,5,0,3,5,1,4,6,2,4]; if(m<3) y-=1; return (y+Math.floor(y/4)-Math.floor(y/100)+Math.floor(y/400)+t[m-1]+d)%7; })(Y,M,D);
            const add = [];
            const today = new Date().toLocaleDateString('sv-SE', { timeZone: 'Asia/Tokyo' });
            if (ymd === today) add.push('ht-today');
            if (wd === 6) add.push('ht-sat');
            if (wd === 0) add.push('ht-sun');
            if (isHolidayLocal(ymd)) add.push('ht-holiday');
            // 付与
            add.forEach(c => c && cell.classList.add(c));
            cell.dataset.htYmdClasses = add.join(' ');
            // 内容が「休」を含む場合も強制で休日色
            if (String(value || '').includes('休')) cell.classList.add('ht-holiday');
          }
          // 自分列の着色 & 他列は視覚的 readonly
          if (x >= 2) {
            const empCol = x - 2; // employees 配列の index
            if(opts?.isProbablyMobile){
              if (empCol === myIndex) {
                cell.classList.add('my-col');
              } else {
                cell.classList.remove('my-col');
              }
              cell.classList.add('readonly');
            }else{
              if (empCol === myIndex) {
                cell.classList.add('my-col');
                cell.classList.remove('readonly');
              } else {
                if (!(opts?.isAdmin)){
                  cell.classList.add('readonly');
                }
                cell.classList.remove('my-col');
              }
            }
          }
          
          if (y > 0 && (y % 2 === 1)) { 
            cell.classList.add('ht-row-am');
          }else{
            cell.classList.remove('ht-row-am');
          }
        } catch (e) {
        }
      },
      // Optional: disable edition for now to keep it simple
      editable: ux?.editable
    });
    if (ux && window.GridControl?.installCommonJSSUX) {
      try { window.GridControl.installCommonJSSUX(grid, host, ux); } catch (e) { console.warn('installCommonJSSUX failed', e); }
    }
    window.GridInstance = grid;
    return grid;
  }
  window.ScheduleGrid = { init };
})();
