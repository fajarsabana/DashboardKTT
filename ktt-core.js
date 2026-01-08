/* ====================== ktt-core.js ======================
   Core utilities untuk semua halaman KTT Dashboard
   - Parsing Excel "DATA PEMAKAIAN PER BULAN"
   - Deteksi bulan & tahun yang toleran (JAN/AGS/AGU dst, 2,025 dll)
   - Agregasi bulan, YTD, YoY, MoM
   - Helper Chart.js & DataTables
   - Session storage (dipakai silang index / industri / pelanggan)
========================================================= */

const KTT = (function () {
  // ---- Locale & konstanta dasar ----
  const nf = new Intl.NumberFormat('id-ID');
  const MONTHS = {
    1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR',
    5: 'MEI', 6: 'JUN', 7: 'JUL', 8: 'AGS',
    9: 'SEP', 10: 'OKT', 11: 'NOV', 12: 'DES'
  };
  const MONTH_LABEL = (m) => String(m).padStart(2, '0');

  // ---- Helper numerik & nama pelanggan ----
  function isFiniteNum(v) {
    return typeof v === 'number' && isFinite(v);
  }

  function toNumber(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    const s = String(v)
      .replace(/\s/g, '')
      .replace(/\./g, '')
      .replace(/,/g, '.')
      .replace(/[^0-9eE.+\-]/g, '');
    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }

  function isBadName(n) {
    const s = String(n ?? '').trim();
    if (!s) return true;
    // kalau isinya cuma angka / tanda baca → dianggap jelek
    return /^[0-9.,]+$/.test(s);
  }

  function prettyName(n, idpel) {
    const s = String(n ?? '').trim();
    if (!isBadName(s)) return s;
    const id = String(idpel ?? '').trim().replace(/\.0+$/, '');
    return id ? `IDPEL ${id}` : '(Tanpa Nama)';
  }

  // normalisasi string (buat key industri dll)
  function norm(s) {
    return String(s ?? '')
      .trim()
      .toUpperCase()
      .replace(/\s+/g, ' ');
  }

  // ---- Debug store ringan ----
  const debugStore = {
    lastParse: null,
    negatives: [],
    warnings: [],
    lastMonthSummary: null
  };
  function debug() {
    return JSON.parse(JSON.stringify(debugStore));
  }

  // ---- Peta token bulan (Indonesia + variasi) ----
  const ID_MONTHS = {
    'JANUARI': 1, 'JAN': 1,
    'FEBRUARI': 2, 'FEB': 2,
    'MARET': 3, 'MAR': 3,
    'APRIL': 4, 'APR': 4,
    'MEI': 5, 'MAY': 5,
    'JUNI': 6, 'JUN': 6,
    'JULI': 7, 'JUL': 7,
    'AGUSTUS': 8, 'AGU': 8, 'AGS': 8, 'AGT': 8, 'AGST': 8, 'AUG': 8,
    'SEPTEMBER': 9, 'SEPT': 9, 'SEP': 9,
    'OKTOBER': 10, 'OKT': 10, 'OCT': 10,
    'NOVEMBER': 11, 'NOV': 11,
    'DESEMBER': 12, 'DES': 12, 'DEC': 12
  };

  function extractMonthToken(s) {
    if (!s) return null;
    const t = String(s)
      .toUpperCase()
      .replace(/[()\/\-.,]/g, ' ')
      .replace(/\s+/g, ' ');
    // cari token terpanjang dulu supaya "SEPTEMBER" tidak ketendang "SEP"
    const keys = Object.keys(ID_MONTHS).sort((a, b) => b.length - a.length);
    for (const k of keys) {
      const re = new RegExp(`(^|\\s)${k}(\\s|$)`);
      if (re.test(t)) return ID_MONTHS[k];
    }

// month header yang *strict* (hanya token bulan saja, bukan kalimat seperti "SELISIH NOV vs OKT")
function extractMonthStrictHeader(s) {
  if (s === null || s === undefined) return null;
  const t = String(s)
    .toUpperCase()
    .trim()
    .replace(/[()\/\-.,]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // angka murni 1..12
  if (/^\d{1,2}$/.test(t)) {
    const n = Number(t);
    return n >= 1 && n <= 12 ? n : null;
  }

  // kalau ada spasi → bukan token bulan murni
  if (t.includes(' ')) return null;

  return ID_MONTHS[t] || null;
}

// fallback ekstra untuk format "BULAN 8" / "BLN 8" yang aman
function extractMonthFromTextSafe(s) {
  if (!s) return null;
  const t = String(s)
    .toUpperCase()
    .replace(/[()\/\-.,]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!/\b(BULAN|BLN|MONTH)\b/.test(t)) return null;
  const m = t.match(/\b(0?[1-9]|1[0-2])\b/);
  return m ? Number(m[1]) : null;
}
    return null;
  }

  // ---- Deteksi tahun ----
  function extractYearToken(s) {
    if (!s) return null;
    // buang semua selain digit → "2,025" → "2025"
    const onlyDigits = String(s).replace(/[^\d]/g, '');
    const m = onlyDigits.match(/(19|20)\d{2}/);
    return m ? +m[0] : null;
  }

  // forward-fill tahun berdasarkan kombinasi header baris 1 & 2
  function extractYearWithCarry(h0, h1, startIdx = 8) {
    const out = [];
    let last = null;
    const len = Math.max(h0.length, h1.length);
    for (let c = startIdx; c < len; c++) {
      const combined = `${h0[c] ?? ''} ${h1[c] ?? ''}`;
      const y = extractYearToken(combined);
      if (y) last = y;
      out[c] = last;
    }
    return out;
  }

  // ================== PARSE EXCEL ==================
  // fileOrBuffer: File dari <input> atau ArrayBuffer (session)
  async function parseExcelFile(fileOrBuffer) {
    let data;
    if (fileOrBuffer instanceof ArrayBuffer) {
      data = fileOrBuffer;
    } else {
      data = await fileOrBuffer.arrayBuffer();
    }

    const wb = XLSX.read(data, { type: 'array', cellDates: false, raw: false });
    const nameCandidate =
      wb.SheetNames.find((n) => /data pemakaian per bulan/i.test(n)) ||
      wb.SheetNames[0];
    const ws = wb.Sheets[nameCandidate];

    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const H0 = raw[0] || [];
    const H1 = raw[1] || [];

    debugStore.lastParse = {
      workbook: wb.SheetNames,
      sheet: nameCandidate,
      headerRow0: H0.slice(0, 40),
      headerRow1: H1.slice(0, 40)
    };

    // deteksi kolom bulan
    const yearByCol = extractYearWithCarry(H0, H1, 8);
    const monthCols = [];
    const maxLen = Math.max(H0.length, H1.length);

    for (let c = 8; c < maxLen; c++) {
      const comb = `${H0[c] ?? ''} ${H1[c] ?? ''}`;
      const mon =
  extractMonthStrictHeader(H1[c]) ||
  ((H1[c] == null || String(H1[c]).trim() === '')
    ? extractMonthStrictHeader(H0[c])
    : null) ||
  extractMonthFromTextSafe(comb);
      const yr = yearByCol[c] || extractYearToken(comb);
      if (mon && yr) {
        monthCols.push({ c, year: yr, month: mon, header: comb });
      }
    }

    if (!monthCols.length) {
      throw new Error(
        'Kolom bulan tidak terdeteksi — cek header sheet (baris 1–2).'
      );
    }

    // Indeks field baris data (1-based karena dari Excel)
    const IDX = {
      UID: 1,
      PROV: 2,
      IDPEL: 3,
      NAMA: 4,
      JENIS: 5,
      TARIF: 6,
      DAYA: 7
    };

    const DATA = [];
    const LABELS = new Map();

    // cari row awal data (lewati judul / header tambahan)
    let startRow = 2;
    for (let r = 2; r < Math.min(20, raw.length); r++) {
      const v = raw[r] && raw[r][IDX.IDPEL];
      if (v && String(v).trim() !== '') {
        startRow = r;
        break;
      }
    }

    for (let r = startRow; r < raw.length; r++) {
      const row = raw[r];
      if (!row) continue;

      const idpel = row[IDX.IDPEL];
      const nama = row[IDX.NAMA];

      // skip baris kosong
      if (
        (idpel == null || String(idpel).trim() === '') &&
        (nama == null || String(nama).trim() === '')
      ) {
        continue;
      }

      // skip baris TOTAL / JUMLAH dsb
      const upNama = String(nama ?? '').toUpperCase();
      if (/TOTAL|JUMLAH|GRAND TOTAL/.test(upNama)) continue;

      for (const mc of monthCols) {
        const rawVal = row[mc.c];
        if (rawVal === null || rawVal === undefined || rawVal === '') continue;

        const kwh = toNumber(rawVal);

        if (isFiniteNum(kwh) && kwh < 0) {
          // catat saja, tapi dipaksa 0 supaya nggak bikin grafik aneh
          debugStore.negatives.push({
            row: r + 1,
            col: mc.c + 1,
            idpel,
            val: kwh,
            year: mc.year,
            month: mc.month
          });
        }

        const safeKwh = isFiniteNum(kwh) && kwh > 0 ? kwh : 0;

        const obj = {
          UID: row[IDX.UID] || '',
          PROV: row[IDX.PROV] || '',
          IDPEL: String(idpel || '').trim(),
          NAMA: nama || '',
          JENIS: row[IDX.JENIS] || '',
          TARIF: row[IDX.TARIF] || '',
          DAYA: row[IDX.DAYA] || '',
          TAHUN: mc.year,
          BULAN: mc.month,
          KWH: safeKwh
        };
        DATA.push(obj);

        // update label pelanggan
        if (obj.IDPEL) {
          const cur = LABELS.get(obj.IDPEL);
          const cand = prettyName(obj.NAMA, obj.IDPEL);
          if (!cur) LABELS.set(obj.IDPEL, cand);
          else if (isBadName(cur) && !isBadName(cand)) LABELS.set(obj.IDPEL, cand);
        }
      }
    }

    // ringkasan per bulan (buat ngecek anomali kalau perlu)
    const monthSums = new Map();
    for (const d of DATA) {
      const key = `${d.TAHUN}-${MONTH_LABEL(d.BULAN)}`;
      monthSums.set(key, (monthSums.get(key) || 0) + (d.KWH || 0));
    }
    debugStore.lastMonthSummary = [...monthSums.entries()].sort();

    return {
      DATA,
      LABELS,
      meta: {
        sheet: nameCandidate,
        rows: raw.length,
        monthColsDetected: monthCols.length
      }
    };
  }

  // ================== Agregasi waktu ==================

  // Map 'YYYY-MM' → total kWh (dari subset DATA)
  function buildMonthMap(rows) {
    const mm = new Map();
    for (const r of rows) {
      if (!r || r.TAHUN == null || r.BULAN == null) continue;
      const key = `${r.TAHUN}-${MONTH_LABEL(r.BULAN)}`;
      mm.set(key, (mm.get(key) || 0) + (r.KWH || 0));
    }
    return mm;
  }

  // YTD sum untuk (year, monthEnd) dari map 'YYYY-MM'
  function ytdSum(mm, year, monthEnd) {
    if (!year || !monthEnd) return 0;
    let total = 0;
    for (let m = 1; m <= monthEnd; m++) {
      const key = `${year}-${MONTH_LABEL(m)}`;
      total += mm.get(key) || 0;
    }
    return total;
  }

  // Dari subset 1 tahun → cari last month aktif + array sum per bulan
  function lastActiveMonth(rows) {
    const sums = {};
    let last = 0;
    for (const r of rows) {
      if (!r || r.BULAN == null) continue;
      const m = Number(r.BULAN);
      if (!sums[m]) sums[m] = 0;
      sums[m] += r.KWH || 0;
      if (sums[m] > 0 && m > last) last = m;
    }
    return { last, sums };
  }

  // ================== Chart helpers ==================

  function drawLine(
    canvasId,
    year,
    sums,          // object: {1: kWh, 2: kWh, ...}
    lastM,
    selectedMonth,
    onClickMonth
  ) {
    const ctx = document.getElementById(canvasId).getContext('2d');
    const labels = [];
    const data = [];
    for (let m = 1; m <= 12; m++) {
      labels.push(MONTH_LABEL(m));
      data.push(sums[m] || 0);
    }

    const chart = new Chart(ctx, {
      type: 'line',
      data: {
        labels,
        datasets: [
          {
            label: `${year} (kWh)`,
            data,
            fill: false,
            tension: 0.2,
            pointRadius: 4,
            pointHoverRadius: 6
          }
        ]
      },
      options: {
        responsive: true,
        plugins: { legend: { display: true } },
        scales: {
          y: { beginAtZero: true }
        }
      }
    });

    if (typeof onClickMonth === 'function') {
      document.getElementById(canvasId).onclick = function (evt) {
        const el = chart.getElementsAtEventForMode(
          evt,
          'nearest',
          { intersect: true },
          true
        );
        if (!el.length) return;
        const idx = el[0].index;
        const month = idx + 1; // 0-based index
        onClickMonth(month);
      };
    }

    return chart;
  }

  function drawPie(canvasId, labels, data, onClickIdx) {
    const ctx = document.getElementById(canvasId).getContext('2d');
    const chart = new Chart(ctx, {
      type: 'pie',
      data: {
        labels,
        datasets: [{ data }]
      },
      options: { responsive: true }
    });

    if (typeof onClickIdx === 'function') {
      document.getElementById(canvasId).onclick = function (evt) {
        const el = chart.getElementsAtEventForMode(
          evt,
          'nearest',
          { intersect: true },
          true
        );
        if (!el.length) return;
        onClickIdx(el[0].index);
      };
    }

    return chart;
  }

  // ================== DataTables helper ==================

  function setDataTable(selector, opts = {}, rows = [], onInit) {
    if ($.fn.dataTable.isDataTable(selector)) {
      $(selector).DataTable().destroy();
    }
    $(selector + ' tbody').empty();

    const options = Object.assign(
      {
        data: rows,
        columns: Array.from(
          { length: (rows[0] || []).length },
          () => ({ title: '', defaultContent: '' })
        ),
        order:
          opts.defaultOrderIndex != null
            ? [[opts.defaultOrderIndex, opts.defaultOrderDir || 'asc']]
            : [],
        columnDefs: opts.defs || [],
        paging: opts.paging !== undefined ? opts.paging : true
      },
      opts.extra || {}
    );

    const table = $(selector).DataTable(options);
    if (typeof onInit === 'function') onInit(table);
    return table;
  }

  // ================== Session storage (localStorage) ==================

  const SESSION_KEY = 'KTT_SESSION_V1';

  function saveSession(DATA, LABELS, meta = {}) {
    const store = {
      DATA,
      LABELS: Object.fromEntries(LABELS),
      meta
    };
    try {
      localStorage.setItem(SESSION_KEY, JSON.stringify(store));
      return true;
    } catch (e) {
      console.error(e);
      return false;
    }
  }

  function loadSession() {
    try {
      const raw = localStorage.getItem(SESSION_KEY);
      if (!raw) return { ok: false };
      const obj = JSON.parse(raw);
      const LABELS = new Map(Object.entries(obj.LABELS || {}));
      return {
        ok: true,
        DATA: obj.DATA || [],
        LABELS,
        meta: obj.meta || {}
      };
    } catch (e) {
      console.error(e);
      return { ok: false };
    }
  }

  function clearSession() {
    localStorage.removeItem(SESSION_KEY);
  }

  // ================== Query & navigasi ==================

  function getQuery() {
    const q = {};
    const s = (location.search || '').replace(/^\?/, '');
    if (!s) return q;
    s.split('&').forEach((p) => {
      if (!p) return;
      const [k, v] = p.split('=');
      if (k) q[decodeURIComponent(k)] = decodeURIComponent(v || '');
    });
    return q;
  }

  function goIndustri(indKey, uid, year) {
    const u =
      `industri.html?ind=${encodeURIComponent(indKey || '')}` +
      (uid ? `&uid=${encodeURIComponent(uid)}` : '') +
      (year ? `&year=${encodeURIComponent(year)}` : '');
    location.href = u;
  }

  function goPelanggan(idpel, uid, year) {
    const u =
      `pelanggan.html?pel=${encodeURIComponent(idpel || '')}` +
      (uid ? `&uid=${encodeURIComponent(uid)}` : '') +
      (year ? `&year=${encodeURIComponent(year)}` : '');
    location.href = u;
  }

  // ================== expose API ==================
  return {
    nf,
    MONTHS,
    MONTH_LABEL,
    prettyName,
    isBadName,
    norm,
    toNumber,
    parseExcelFile,
    buildMonthMap,
    ytdSum,
    lastActiveMonth,
    drawLine,
    drawPie,
    setDataTable,
    saveSession,
    loadSession,
    clearSession,
    getQuery,
    goIndustri,
    goPelanggan,
    debug
  };
})();
