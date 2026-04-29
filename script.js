// =============================================
// DASHBOARD MADRASAH TARBIYAH IDADIYAH AL-MIFTAH
// script.js — Versi Offline
// =============================================

// =============================================
// GLOBAL STATE
// =============================================
let arsipAllData = [];
let arsipFiltered = [];
let currentPage = 1;
const PAGE_SIZE = 50;
let kelasChart, donutChart, guruChart, guruPageChart, bulanChart;

// =============================================
// PAGE NAVIGATION
// =============================================
function showPage(pageId, navEl) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById('page-' + pageId).classList.add('active');
  if (navEl) navEl.classList.add('active');
  const titles = {
    dashboard: 'Dashboard Akademik',
    arsip: 'Arsip Tes Tulis',
    murid: 'Data Murid',
    guru: 'Rekap Guru'
  };
  document.getElementById('topbarTitle').textContent = titles[pageId] || 'Dashboard';

  if (pageId === 'murid' && arsipAllData.length === 0) {
    document.getElementById('muridNoData').style.display = 'flex';
    document.getElementById('muridSearchBar').style.display = 'none';
    document.getElementById('muridEmptyState').style.display = 'none';
  } else if (pageId === 'murid') {
    document.getElementById('muridNoData').style.display = 'none';
    document.getElementById('muridSearchBar').style.display = 'flex';
    document.getElementById('muridEmptyState').style.display = 'block';
  }

  if (pageId === 'guru') {
    if (arsipAllData.length === 0) {
      document.getElementById('guruNoData').style.display = 'flex';
    } else {
      document.getElementById('guruNoData').style.display = 'none';
      renderGuruPage();
    }
  }
}

// =============================================
// FILE UPLOAD & PARSE
// =============================================
function handleDragOver(e) {
  e.preventDefault();
  document.getElementById('uploadZone').classList.add('drag-over');
}
function handleDragLeave(e) {
  document.getElementById('uploadZone').classList.remove('drag-over');
}
function handleDrop(e) {
  e.preventDefault();
  document.getElementById('uploadZone').classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
}

function handleFile(file) {
  if (!file) return;
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    alert('❌ File harus berformat .xlsx atau .xls');
    return;
  }
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array', cellDates: true });

      let sheetName = wb.SheetNames.find(n => n.trim().toUpperCase() === 'ARSIP TES TULIS');
      if (!sheetName) sheetName = wb.SheetNames.find(n => n.toUpperCase().includes('ARSIP'));
      if (!sheetName) sheetName = wb.SheetNames.find(n => n.toUpperCase().includes('TES TULIS'));
      if (!sheetName) sheetName = wb.SheetNames[0];

      const ws = wb.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      // Find header row
      let headerIdx = -1;
      for (let i = 0; i < Math.min(json.length, 10); i++) {
        const row = json[i].map(c => String(c).trim().toUpperCase());
        if (row.some(c => c.includes('ID') && c.includes('PPS'))) { headerIdx = i; break; }
      }
      if (headerIdx === -1) {
        for (let i = 0; i < Math.min(json.length, 10); i++) {
          const row = json[i].map(c => String(c).trim().toUpperCase());
          if (row.includes('NAMA') && row.includes('KELAS')) { headerIdx = i; break; }
        }
      }
      if (headerIdx === -1) headerIdx = 2;

      const headers = json[headerIdx].map(c => String(c).trim().toUpperCase());

      const colIdx = {
        idpps:      headers.findIndex(h => h.includes('ID') && h.includes('PPS')),
        nama:       headers.findIndex(h => h === 'NAMA'),
        dom:        headers.findIndex(h => h === 'DOM'),
        kelas:      headers.findIndex(h => h === 'KELAS'),
        pembimbing: headers.findIndex(h => h.includes('PEMBIMBING') || h.includes('PENGAJAR') || h.includes('GURU')),
        tempat:     headers.findIndex(h => h === 'TEMPAT'),
        ket:        headers.findIndex(h => (h.includes('KET') && h.includes('TES')) || h === 'KET. TES' || h === 'KET TES'),
        ltl:        headers.findIndex(h => h === 'L/TL' || h === 'LTL'),
        tgl:        headers.findIndex(h => h === 'TGL' || h.includes('TANGGAL')),
      };

      const rows = [];
      for (let i = headerIdx + 1; i < json.length; i++) {
        const row = json[i];
        if (!row || row.every(c => c === '' || c === null || c === undefined)) continue;
        const idpps = colIdx.idpps >= 0 ? String(row[colIdx.idpps] || '').trim() : '';
        const nama  = colIdx.nama >= 0  ? String(row[colIdx.nama]  || '').trim() : '';
        if (!idpps && !nama) continue;

        let tgl = '';
        if (colIdx.tgl >= 0 && row[colIdx.tgl]) {
          const rawTgl = row[colIdx.tgl];
          if (rawTgl instanceof Date) {
            tgl = rawTgl.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' });
          } else {
            tgl = String(rawTgl).trim();
          }
        }

        const ltlRaw = colIdx.ltl >= 0 ? String(row[colIdx.ltl] || '').trim().toUpperCase() : '';
        const ketRaw = colIdx.ket >= 0 ? String(row[colIdx.ket] || '').trim().toUpperCase() : '';

        // Normalize L/TL
        let ltl = '';
        if (ltlRaw === 'LULUS' || ltlRaw === 'L') ltl = 'LULUS';
        else if (ltlRaw === 'TIDAK LULUS' || ltlRaw === 'TL' || ltlRaw === 'TIDAK') ltl = 'TIDAK LULUS';
        else if (ketRaw.includes('LULUS') && !ketRaw.includes('TIDAK')) ltl = 'LULUS';
        else if (ketRaw.includes('TIDAK')) ltl = 'TIDAK LULUS';
        else ltl = ltlRaw;

        rows.push({
          idpps, nama,
          dom:        colIdx.dom >= 0        ? String(row[colIdx.dom] || '').trim()        : '',
          kelas:      colIdx.kelas >= 0      ? String(row[colIdx.kelas] || '').trim()      : '',
          pembimbing: colIdx.pembimbing >= 0 ? String(row[colIdx.pembimbing] || '').trim() : '',
          tempat:     colIdx.tempat >= 0     ? String(row[colIdx.tempat] || '').trim()     : '',
          ket:        ketRaw,
          ltl,
          tgl,
        });
      }

      arsipAllData = rows;
      arsipFiltered = [...rows];

      // Update UI
      document.getElementById('uploadZone').style.display = 'none';
      document.getElementById('fileInfoBar').style.display = 'flex';
      document.getElementById('fileInfoText').textContent =
        `${file.name} · Sheet: "${sheetName}" · ${rows.length} baris data ditemukan`;
      document.getElementById('sidebarFileName').textContent = file.name;
      document.getElementById('topbarDataInfo').textContent = `${rows.length} data tes tulis`;

      document.getElementById('arsipStatsBar').style.display = 'flex';
      document.getElementById('arsipToolbar').style.display = 'flex';
      document.getElementById('arsipTableWrap').style.display = 'block';
      document.getElementById('noBanner').style.display = 'none';

      // Populate kelas filter
      const kelasSet = [...new Set(rows.map(r => r.kelas).filter(Boolean))].sort();
      const selKelas = document.getElementById('filterKelas');
      selKelas.innerHTML = '<option value="">Semua Kelas</option>';
      kelasSet.forEach(k => {
        const opt = document.createElement('option');
        opt.value = k; opt.textContent = k;
        selKelas.appendChild(opt);
      });

      currentPage = 1;
      updateStats(rows);
      renderTable();
      updateDashboard(rows);

    } catch(err) {
      alert('❌ Gagal membaca file Excel.\n\nPastikan file tidak rusak dan berformat .xlsx/.xls\n\nError: ' + err.message);
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

// =============================================
// STATS
// =============================================
function isLulus(r) {
  return r.ltl === 'LULUS' || (r.ket && r.ket.includes('LULUS') && !r.ket.includes('TIDAK'));
}

function updateStats(rows) {
  const lulus  = rows.filter(r => isLulus(r)).length;
  const gagal  = rows.length - lulus;
  const kelas  = new Set(rows.map(r => r.kelas).filter(Boolean)).size;
  const pemb   = new Set(rows.map(r => r.pembimbing).filter(Boolean)).size;
  const muridUnik = new Set(rows.map(r => r.idpps || r.nama).filter(Boolean)).size;
  document.getElementById('statTotal').textContent      = rows.length;
  document.getElementById('statMuridUnik').textContent  = muridUnik;
  document.getElementById('statLulus').textContent      = lulus;
  document.getElementById('statGagal').textContent      = gagal;
  document.getElementById('statKelas').textContent      = kelas;
  document.getElementById('statPembimbing').textContent = pemb;
}

// =============================================
// DASHBOARD UPDATE FROM DATA
// =============================================
function updateDashboard(rows) {
  const lulus = rows.filter(r => isLulus(r)).length;
  const gagal = rows.length - lulus;
  const pct   = rows.length > 0 ? Math.round(lulus / rows.length * 100) : 0;
  const muridUnik = new Set(rows.map(r => r.idpps || r.nama).filter(Boolean)).size;

  document.getElementById('val-murid').innerHTML = muridUnik + '<span class="unit">murid</span>';
  document.getElementById('sub-murid').textContent = new Set(rows.map(r=>r.kelas).filter(Boolean)).size + ' kelas berbeda';
  document.getElementById('val-tulis').innerHTML = rows.length + '<span class="unit">tes</span>';
  document.getElementById('sub-tulis').textContent = 'Total entri di arsip';
  document.getElementById('val-lulus').innerHTML = pct + '<span class="unit">%</span>';
  document.getElementById('sub-lulus').textContent = lulus + ' lulus dari ' + rows.length + ' tes';

  // Guru stats
  const guruMap = buildGuruMap(rows);
  const guruArr = Object.values(guruMap).filter(g => g.total > 0);
  const avgGuru = guruArr.length > 0 ? Math.round(guruArr.reduce((a, g) => a + g.pct, 0) / guruArr.length) : 0;
  document.getElementById('val-guru').innerHTML = avgGuru + '<span class="unit">%</span>';
  document.getElementById('sub-guru').textContent = 'Rata-rata dari ' + guruArr.length + ' pembimbing';

  // Donut
  document.getElementById('donut-pct').textContent = pct + '%';
  document.getElementById('jumlah-lulus').textContent = lulus;
  document.getElementById('jumlah-tidak-lulus').textContent = gagal;
  buildDonut(lulus, gagal);

  // Kelas chart (grouped)
  buildKelasChart(rows);

  // Bulan hijriah chart
  buildBulanChart(rows);

  // Guru chart & table
  buildGuruChartDashboard(guruArr);
  buildGuruTable(guruArr);
}

function buildGuruMap(rows) {
  const guruMap = {};
  rows.forEach(r => {
    const nama = r.pembimbing || '(Tidak diketahui)';
    if (!guruMap[nama]) guruMap[nama] = { nama, total: 0, lulus: 0, tidakLulus: 0, kelas: new Set() };
    guruMap[nama].total++;
    if (isLulus(r)) guruMap[nama].lulus++;
    else guruMap[nama].tidakLulus++;
    if (r.kelas) guruMap[nama].kelas.add(r.kelas);
  });
  Object.values(guruMap).forEach(g => {
    g.pct = g.total > 0 ? Math.round(g.lulus / g.total * 100) : 0;
  });
  return guruMap;
}

function buildDonut(lulus, gagal) {
  const ctx = document.getElementById('donutChart').getContext('2d');
  if (donutChart) donutChart.destroy();
  donutChart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Lulus', 'Tidak Lulus'],
      datasets: [{ data: [lulus, gagal], backgroundColor: ['#639922', '#E24B4A'], borderWidth: 0, hoverOffset: 4 }]
    },
    options: {
      responsive: true, maintainAspectRatio: true,
      cutout: '72%',
      plugins: { legend: { display: false } }
    }
  });
}

// Map kelas string → kelompok global
function getKelasGroup(kelas) {
  const k = kelas.trim().toUpperCase();
  if (k.includes('SHIFIR') || k.startsWith('SHIFIR') || k === 'SHIFIR') return 'Shifir';
  const jilidMatch = k.match(/JILID\s*(\d+)/);
  if (jilidMatch) return 'Jilid ' + jilidMatch[1];
  if (k.includes('TATHBIQ') || k.includes('TATBIQ') || k.includes('TATHBI')) {
    const tNum = k.match(/(\d+)/);
    return tNum ? 'Tathbiq ' + tNum[1] : 'Tathbiq';
  }
  if (k.includes('TAHTBIQ') || k.includes('TAHTBI')) {
    const tNum = k.match(/(\d+)/);
    return tNum ? 'Tahtbiq ' + tNum[1] : 'Tahtbiq';
  }
  if (k.includes('GHORIB') || k.includes('GHARIB')) return 'Ghorib';
  if (k.includes('TAJWID')) return 'Tajwid';
  if (k.includes('QURAN') || k.includes('QUR')) return 'Al-Quran';
  return 'Lainnya';
}

const KELAS_ORDER = ['Shifir','Jilid 1','Jilid 2','Jilid 3','Jilid 4','Jilid 5','Jilid 6','Tathbiq','Tathbiq 1','Tathbiq 2','Tathbiq 3','Tahtbiq','Tahtbiq 1','Tahtbiq 2','Ghorib','Tajwid','Al-Quran','Lainnya'];

function buildKelasChart(rows) {
  const groupMap = {};
  rows.forEach(r => {
    const g = getKelasGroup(r.kelas || '');
    if (!groupMap[g]) groupMap[g] = { lulus: 0, tidakLulus: 0 };
    if (isLulus(r)) groupMap[g].lulus++;
    else groupMap[g].tidakLulus++;
  });

  const groups = Object.keys(groupMap).sort((a, b) => {
    const ia = KELAS_ORDER.indexOf(a), ib = KELAS_ORDER.indexOf(b);
    if (ia === -1 && ib === -1) return a.localeCompare(b);
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });

  const lulusData = groups.map(g => groupMap[g].lulus);
  const tlData    = groups.map(g => groupMap[g].tidakLulus);

  const ctx = document.getElementById('kelasChart').getContext('2d');
  if (kelasChart) kelasChart.destroy();
  kelasChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: groups,
      datasets: [
        { label: 'Lulus',       data: lulusData, backgroundColor: 'rgba(99,153,34,0.82)',  borderRadius: 4 },
        { label: 'Tidak Lulus', data: tlData,    backgroundColor: 'rgba(226,75,74,0.78)',  borderRadius: 4 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 11 }, color: '#8A9B78' } },
        y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 }, color: '#8A9B78' }, beginAtZero: true }
      }
    }
  });
}

const HIJRI_MONTHS = ['Muharram','Safar','Rabi\'ul Awwal','Rabi\'ul Akhir','Jumadal Ula','Jumadal Akhirah','Rajab','Sya\'ban','Ramadhan','Syawal','Dzulqa\'dah','Dzulhijjah'];

function parseTglHijri(tgl) {
  if (!tgl) return null;
  const parts = String(tgl).split('/');
  if (parts.length !== 3) return null;
  const d = parseInt(parts[0]), m = parseInt(parts[1]), y = parseInt(parts[2]);
  if (isNaN(d) || isNaN(m) || isNaN(y) || m < 1 || m > 12) return null;
  return { d, m, y, key: `${y}-${String(m).padStart(2,'0')}`, label: HIJRI_MONTHS[m-1] + ' ' + y };
}

function buildBulanChart(rows) {
  const bulanMap = {};
  rows.forEach(r => {
    const parsed = parseTglHijri(r.tgl);
    if (!parsed) return;
    if (!bulanMap[parsed.key]) bulanMap[parsed.key] = { label: parsed.label, lulus: 0, tidakLulus: 0 };
    if (isLulus(r)) bulanMap[parsed.key].lulus++;
    else bulanMap[parsed.key].tidakLulus++;
  });

  const sorted = Object.entries(bulanMap).sort((a, b) => a[0].localeCompare(b[0]));
  const labels    = sorted.map(e => e[1].label);
  const lulusData = sorted.map(e => e[1].lulus);
  const tlData    = sorted.map(e => e[1].tidakLulus);

  const ctx = document.getElementById('bulanChart').getContext('2d');
  if (bulanChart) bulanChart.destroy();
  bulanChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Lulus',       data: lulusData, backgroundColor: 'rgba(99,153,34,0.82)',  borderRadius: 4 },
        { label: 'Tidak Lulus', data: tlData,    backgroundColor: 'rgba(226,75,74,0.78)',  borderRadius: 4 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false },
        tooltip: { callbacks: { title: (items) => items[0].label } }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 }, color: '#8A9B78', maxRotation: 35 } },
        y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 }, color: '#8A9B78' }, beginAtZero: true }
      }
    }
  });
}

function buildGuruChartDashboard(guruArr) {
  const sorted = [...guruArr].sort((a, b) => b.pct - a.pct).slice(0, 10);
  const labels = sorted.map(g => g.nama.length > 20 ? g.nama.substring(0, 18) + '…' : g.nama);
  const values = sorted.map(g => g.pct);
  const colors = values.map(v => v >= 80 ? '#639922' : v >= 60 ? '#1D9E75' : '#BA7517');

  const ctx = document.getElementById('guruChart').getContext('2d');
  if (guruChart) guruChart.destroy();
  guruChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ label: '% Lulus', data: values, backgroundColor: colors, borderRadius: 4, barThickness: 16 }]
    },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { min: 0, max: 100, grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 10 }, color: '#8A9B78', callback: v => v + '%' } },
        y: { grid: { display: false }, ticks: { font: { size: 11 }, color: '#5A6B48' } }
      }
    }
  });
}

function buildGuruTable(guruArr) {
  const sorted = [...guruArr].sort((a, b) => b.pct - a.pct);
  const tbody = document.getElementById('guruTbody');
  if (sorted.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:24px;color:var(--text-muted);">Tidak ada data pembimbing</td></tr>';
    return;
  }
  tbody.innerHTML = sorted.map(g => {
    const cls = g.pct >= 80 ? 'tinggi' : g.pct >= 60 ? 'sedang' : 'rendah';
    const label = g.pct >= 80 ? 'Baik' : g.pct >= 60 ? 'Cukup' : 'Perlu Perhatian';
    return `<tr>
      <td><div class="guru-name">${g.nama}</div></td>
      <td>${g.total}</td>
      <td style="color:var(--green-600);font-weight:600;">${g.lulus}</td>
      <td style="color:#A32D2D;font-weight:600;">${g.tidakLulus}</td>
      <td>
        <div style="display:flex;align-items:center;gap:8px;">
          <div class="progress-bar-wrap"><div class="progress-bar" style="width:${g.pct}%;background:${g.pct>=80?'var(--green-400)':g.pct>=60?'var(--teal-400)':'#BA7517'}"></div></div>
          <span style="font-weight:700;font-size:12px;">${g.pct}%</span>
        </div>
      </td>
      <td><span class="badge-lulus ${cls}">${label}</span></td>
    </tr>`;
  }).join('');
}

function filterGuruTable(q) {
  const rows = document.querySelectorAll('#guruTbody tr');
  rows.forEach(r => {
    r.style.display = r.textContent.toLowerCase().includes(q.toLowerCase()) ? '' : 'none';
  });
}

// =============================================
// ARSIP TABLE
// =============================================
function applyFilters() {
  const q   = document.getElementById('arsipSearch').value.trim().toLowerCase();
  const kelas = document.getElementById('filterKelas').value;
  const ltl   = document.getElementById('filterLTL').value.toUpperCase();

  arsipFiltered = arsipAllData.filter(r => {
    if (q && !(r.nama.toLowerCase().includes(q) ||
               r.idpps.toLowerCase().includes(q) ||
               r.pembimbing.toLowerCase().includes(q) ||
               r.kelas.toLowerCase().includes(q) ||
               r.dom.toLowerCase().includes(q))) return false;
    if (kelas && r.kelas !== kelas) return false;
    if (ltl) {
      if (ltl === 'LULUS' && r.ltl !== 'LULUS') return false;
      if (ltl === 'TIDAK LULUS' && r.ltl !== 'TIDAK LULUS') return false;
    }
    return true;
  });

  currentPage = 1;
  renderTable();
}

function renderTable() {
  const total = arsipFiltered.length;
  const totalPages = Math.ceil(total / PAGE_SIZE) || 1;
  const start = (currentPage - 1) * PAGE_SIZE;
  const end   = Math.min(start + PAGE_SIZE, total);
  const pageData = arsipFiltered.slice(start, end);

  document.getElementById('arsipCountBadge').textContent = total + ' Data';
  document.getElementById('paginationInfo').textContent =
    total === 0 ? 'Tidak ada data' : `Menampilkan ${start + 1}–${end} dari ${total} data`;

  const tbody = document.getElementById('arsipTbody');
  if (pageData.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" style="text-align:center;padding:48px;color:var(--text-muted);">
      <div style="font-size:40px;margin-bottom:12px;">🔍</div>
      <div style="font-weight:600;">Tidak ada data ditemukan</div>
    </td></tr>`;
  } else {
    tbody.innerHTML = pageData.map((r, i) => {
      const ltlClass = r.ltl === 'TIDAK LULUS' ? 'tidak-lulus' : r.ltl === 'LULUS' ? 'lulus' : '';
      const ketClass = r.ket.includes('TIDAK') ? 'tidak-lulus' : r.ket.includes('LULUS') ? 'lulus' : '';
      return `<tr>
        <td style="text-align:center;color:var(--text-light);font-size:11px;">${start + i + 1}</td>
        <td class="id-pps">${r.idpps || '-'}</td>
        <td class="nama-col">${r.nama || '-'}</td>
        <td style="text-align:center;font-weight:600;font-size:12px;">${r.dom || '-'}</td>
        <td style="font-weight:600;color:var(--green-600);">${r.kelas || '-'}</td>
        <td style="font-size:12px;">${r.pembimbing || '-'}</td>
        <td style="font-size:12px;">${r.tempat || '-'}</td>
        <td><span class="ket-badge ${ketClass}">${r.ket || '-'}</span></td>
        <td><span class="ket-badge ${ltlClass}">${r.ltl || '-'}</span></td>
        <td style="font-size:11px;color:var(--text-muted);">${r.tgl || '-'}</td>
      </tr>`;
    }).join('');
  }

  // Pagination
  const btns = document.getElementById('paginationBtns');
  btns.innerHTML = '';

  const prevBtn = document.createElement('button');
  prevBtn.className = 'pg-btn';
  prevBtn.textContent = '← Prev';
  prevBtn.disabled = currentPage === 1;
  prevBtn.onclick = () => { currentPage--; renderTable(); };
  btns.appendChild(prevBtn);

  let startPage = Math.max(1, currentPage - 3);
  let endPage   = Math.min(totalPages, startPage + 6);
  startPage     = Math.max(1, endPage - 6);

  for (let p = startPage; p <= endPage; p++) {
    const pb = document.createElement('button');
    pb.className = 'pg-btn' + (p === currentPage ? ' active' : '');
    pb.textContent = p;
    pb.onclick = ((pg) => () => { currentPage = pg; renderTable(); })(p);
    btns.appendChild(pb);
  }

  const nextBtn = document.createElement('button');
  nextBtn.className = 'pg-btn';
  nextBtn.textContent = 'Next →';
  nextBtn.disabled = currentPage >= totalPages;
  nextBtn.onclick = () => { currentPage++; renderTable(); };
  btns.appendChild(nextBtn);
}

function clearData() {
  arsipAllData = [];
  arsipFiltered = [];
  document.getElementById('fileInput').value = '';
  document.getElementById('uploadZone').style.display = '';
  document.getElementById('fileInfoBar').style.display = 'none';
  document.getElementById('arsipStatsBar').style.display = 'none';
  document.getElementById('arsipToolbar').style.display = 'none';
  document.getElementById('arsipTableWrap').style.display = 'none';
  document.getElementById('noBanner').style.display = '';
  document.getElementById('sidebarFileName').textContent = 'Belum ada file';
  document.getElementById('topbarDataInfo').textContent = 'Belum ada data';
  document.getElementById('arsipSearch').value = '';
  document.getElementById('filterKelas').value = '';
  document.getElementById('filterLTL').value = '';
}

function exportFiltered() {
  if (!arsipFiltered.length) return;
  const wb = XLSX.utils.book_new();
  const headers = ['No','ID PPS','Nama','DOM','Kelas','Pembimbing','Tempat','Ket. Tes','L/TL','Tgl'];
  const wsData  = [headers, ...arsipFiltered.map((r, i) => [
    i + 1, r.idpps, r.nama, r.dom, r.kelas, r.pembimbing, r.tempat, r.ket, r.ltl, r.tgl
  ])];
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [4,10,28,6,10,24,14,16,12,10].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws, 'Arsip Tes Tulis');
  XLSX.writeFile(wb, 'arsip_tes_tulis_filtered.xlsx');
}

// =============================================
// MURID SEARCH & DETAIL
// =============================================
function searchMurid(q) {
  const resultsList = document.getElementById('muridResultsList');
  const emptyState  = document.getElementById('muridEmptyState');
  const detailSec   = document.getElementById('muridDetailSection');
  const countEl     = document.getElementById('muridSearchCount');

  if (!q || q.length < 2) {
    resultsList.style.display = 'none';
    emptyState.style.display = 'block';
    detailSec.style.display = 'none';
    countEl.textContent = '';
    return;
  }

  const muridMap = {};
  arsipAllData.forEach(r => {
    const key = r.idpps || r.nama;
    if (!muridMap[key]) muridMap[key] = { idpps: r.idpps, nama: r.nama, riwayat: [] };
    muridMap[key].riwayat.push(r);
  });

  const ql = q.toLowerCase();
  const results = Object.values(muridMap).filter(m =>
    m.nama.toLowerCase().includes(ql) || m.idpps.toLowerCase().includes(ql)
  ).slice(0, 30);

  countEl.textContent = results.length + ' murid ditemukan';

  if (results.length === 0) {
    resultsList.style.display = 'block';
    resultsList.innerHTML = '<div style="padding:24px;text-align:center;color:var(--text-muted);">Tidak ada murid yang cocok</div>';
    emptyState.style.display = 'none';
    return;
  }

  detailSec.style.display = 'none';
  emptyState.style.display = 'none';
  resultsList.style.display = 'block';
  resultsList.innerHTML = results.map(m => {
    const l  = m.riwayat.filter(r => isLulus(r)).length;
    const tl = m.riwayat.length - l;
    return `<div class="murid-result-item" onclick="showMuridDetail('${(m.idpps || m.nama).replace(/'/g,"\\'")}')">
      <div>
        <div class="murid-result-name">${m.nama || '-'}</div>
        <div class="murid-result-meta">ID: ${m.idpps || '-'} · ${m.riwayat.length}× tes · ${l}L / ${tl}TL</div>
      </div>
      <span class="murid-result-badge">${m.riwayat.length}× tes</span>
    </div>`;
  }).join('');
}

function showMuridDetail(key) {
  const muridMap = {};
  arsipAllData.forEach(r => {
    const k = r.idpps || r.nama;
    if (!muridMap[k]) muridMap[k] = { idpps: r.idpps, nama: r.nama, riwayat: [] };
    muridMap[k].riwayat.push(r);
  });

  const m = muridMap[key];
  if (!m) return;

  document.getElementById('muridResultsList').style.display = 'none';
  document.getElementById('muridEmptyState').style.display = 'none';
  document.getElementById('muridDetailSection').style.display = 'block';

  const l  = m.riwayat.filter(r => isLulus(r)).length;
  const tl = m.riwayat.length - l;
  const kelasList = [...new Set(m.riwayat.map(r => r.kelas).filter(Boolean))];

  const initials = m.nama ? m.nama.split(' ').map(w => w[0]).slice(0, 2).join('').toUpperCase() : 'م';
  document.getElementById('muridAvatarBig').textContent = initials;
  document.getElementById('muridDetailName').textContent = m.nama || '—';
  document.getElementById('muridDetailMeta').textContent = `ID PPS: ${m.idpps || '—'} · DOM: ${m.riwayat[0]?.dom || '—'}`;
  document.getElementById('muridTotalTes').textContent = m.riwayat.length;
  document.getElementById('muridTotalL').textContent  = l;
  document.getElementById('muridTotalTL').textContent = tl;
  document.getElementById('muridKelasCount').textContent = kelasList.length;

  const tagsEl = document.getElementById('muridKelasTags');
  tagsEl.innerHTML = kelasList.map(k => `<span class="kelas-tag">${k}</span>`).join('');

  document.getElementById('muridRiwayatCount').textContent = m.riwayat.length + ' Entri';
  const tbody = document.getElementById('muridRiwayatTbody');
  tbody.innerHTML = m.riwayat.map((r, i) => {
    const ltlClass = r.ltl === 'TIDAK LULUS' ? 'tidak-lulus' : r.ltl === 'LULUS' ? 'lulus' : '';
    const ketClass = r.ket.includes('TIDAK') ? 'tidak-lulus' : r.ket.includes('LULUS') ? 'lulus' : '';
    return `<tr>
      <td style="color:var(--text-light);font-size:11px;">${i+1}</td>
      <td><span class="kelas-tag">${r.kelas || '-'}</span></td>
      <td style="font-weight:500;">${r.pembimbing || '-'}</td>
      <td style="font-size:12px;">${r.tempat || '-'}</td>
      <td><span class="ket-badge ${ketClass}">${r.ket || '-'}</span></td>
      <td><span class="ket-badge ${ltlClass}">${r.ltl || '-'}</span></td>
      <td style="font-size:11px;color:var(--text-muted);">${r.tgl || '-'}</td>
    </tr>`;
  }).join('');
}

function closeMuridDetail() {
  document.getElementById('muridDetailSection').style.display = 'none';
  document.getElementById('muridEmptyState').style.display = 'block';
  document.getElementById('muridSearchInput').value = '';
  document.getElementById('muridSearchCount').textContent = '';
}

// =============================================
// REKAP GURU PAGE
// =============================================
let guruPageAllData = [];

function renderGuruPage() {
  const guruMap = buildGuruMap(arsipAllData);
  const guruArr = Object.values(guruMap).sort((a, b) => b.pct - a.pct);
  guruPageAllData = guruArr;

  document.getElementById('guruPageCount').textContent = guruArr.length + ' Guru';

  const ctx = document.getElementById('guruPageChart').getContext('2d');
  if (guruPageChart) guruPageChart.destroy();
  const labels = guruArr.map(g => g.nama.length > 22 ? g.nama.substring(0, 20) + '…' : g.nama);
  const values = guruArr.map(g => g.pct);
  const colors = values.map(v => v >= 80 ? '#639922' : v >= 60 ? '#1D9E75' : '#BA7517');

  guruPageChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '% Lulus',
        data: values,
        backgroundColor: colors,
        borderRadius: 5,
        barThickness: 18
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (c) => ` ${c.raw}% lulus (${guruArr[c.dataIndex].lulus} dari ${guruArr[c.dataIndex].total} tes)`
          }
        }
      },
      scales: {
        x: { min: 0, max: 100, grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 10 }, color: '#8A9B78', callback: v => v + '%' } },
        y: { grid: { display: false }, ticks: { font: { size: 11 }, color: '#5A6B48' } }
      }
    }
  });

  renderGuruPageTable(guruArr);
}

function renderGuruPageTable(guruArr) {
  const tbody = document.getElementById('guruPageTbody');
  if (!guruArr.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:var(--text-muted);">Tidak ada data</td></tr>';
    return;
  }
  tbody.innerHTML = guruArr.map((g, i) => {
    const cls = g.pct >= 80 ? 'tinggi' : g.pct >= 60 ? 'sedang' : 'rendah';
    const label = g.pct >= 80 ? '✅ Baik' : g.pct >= 60 ? '⚠️ Cukup' : '❌ Perlu Perhatian';
    const kelasList = [...g.kelas].sort().join(', ');
    return `<tr>
      <td style="color:var(--text-light);font-size:11px;text-align:center;">${i+1}</td>
      <td style="font-weight:600;">${g.nama}</td>
      <td style="font-size:11px;color:var(--text-muted);max-width:200px;white-space:normal;">${kelasList || '-'}</td>
      <td style="text-align:center;font-weight:600;">${g.total}</td>
      <td style="text-align:center;font-weight:700;color:var(--green-600);">${g.lulus}</td>
      <td style="text-align:center;font-weight:700;color:#A32D2D;">${g.tidakLulus}</td>
      <td>
        <div style="display:flex;align-items:center;gap:8px;">
          <div class="progress-bar-wrap" style="width:80px;"><div class="progress-bar" style="width:${g.pct}%;background:${g.pct>=80?'var(--green-400)':g.pct>=60?'var(--teal-400)':'#BA7517'}"></div></div>
          <span style="font-weight:700;font-size:13px;">${g.pct}%</span>
        </div>
      </td>
      <td><span class="badge-lulus ${cls}">${label}</span></td>
    </tr>`;
  }).join('');
}

function filterGuruPageTable(q) {
  const filtered = q
    ? guruPageAllData.filter(g => g.nama.toLowerCase().includes(q.toLowerCase()))
    : guruPageAllData;
  document.getElementById('guruPageCount').textContent = filtered.length + ' Guru';
  renderGuruPageTable(filtered);
}
