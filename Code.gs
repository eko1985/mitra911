/**
 * MITRA 911 ONCALL — Server Google Spreadsheet + Penerjemah otomatis
 * -------------------------------------------------------------------
 * Pasang skrip ini dari spreadsheet: Ekstensi > Apps Script.
 * 1) Ganti KUNCI di bawah dengan kata sandi tim (wajib).
 * 2) Jalankan fungsi "setup" sekali, izinkan akses.
 * 3) Terapkan > Deployment baru > Aplikasi web (Jalankan sebagai: Saya, Akses: Siapa saja).
 * 4) Tempel URL /exec dan KUNCI di aplikasi: Pengaturan > Server & sinkronisasi.
 */
const KUNCI = 'GANTI-KATA-SANDI-TIM';
const TZ = 'Asia/Jakarta';

const OUT = { hotel: 'Rawat di hotel / perbaikan', rujuk: 'Dirujuk ke RS', evak: 'Evakuasi medis', tolak: 'Menolak tindakan (APS)', lain: 'Lainnya' };
const DT = 'dd/MM/yyyy HH:mm';
const SHEETS = {
  Pasien: [['ID', '@'], ['Server TS', '0'], ['Diperbarui (ms)', '0'], ['Status', '@'], ['Dihapus', '@'], ['Tgl Datang', DT], ['Nama', '@'], ['No. Paspor', '@'], ['Kewarganegaraan', '@'], ['JK', '@'], ['Usia', '0'], ['Hotel', '@'], ['Kamar', '@'], ['Telepon', '@'], ['Asuransi', '@'], ['Alergi', '@'], ['Keluhan Utama', '@'], ['Diagnosis', '@'], ['NEWS2 Terakhir', '0'], ['Hasil', '@'], ['Tujuan Rujukan', '@'], ['Waktu Selesai', DT], ['Dokter', '@'], ['Perawat', '@'], ['Diubah Oleh', '@'], ['JSON1', '@'], ['JSON2', '@'], ['JSON3', '@'], ['JSON4', '@']],
  TTV: [['ID Pasien', '@'], ['Nama', '@'], ['No. Paspor', '@'], ['Waktu', DT], ['Sistolik', '0'], ['Diastolik', '0'], ['Nadi', '0'], ['RR', '0'], ['Suhu', '0.0'], ['SpO2', '0'], ['O2 Tambahan', '@'], ['AVPU', '@'], ['GDS', '0'], ['Nyeri', '0'], ['NEWS2', '0'], ['Catatan', '@'], ['ID TTV', '@']],
  Advice: [['ID Pasien', '@'], ['Nama', '@'], ['Waktu', DT], ['Dokter', '@'], ['Diagnosis', '@'], ['Terapi', '@'], ['Instruksi', '@'], ['Tindak Lanjut', '@'], ['ID Advice', '@']],
  Tindakan: [['ID Pasien', '@'], ['Nama', '@'], ['Waktu', DT], ['Perawat', '@'], ['Tindakan', '@'], ['Detail', '@'], ['ID Tindakan', '@']],
  Log: [['Waktu', DT], ['Perangkat / Petugas', '@'], ['Aksi', '@'], ['Dikirim', '0'], ['Diterima', '0']]
};
const P = name => SHEETS.Pasien.findIndex(h => h[0] === name); // 0-based

/* ================= endpoint ================= */
function doGet() { return out_({ ok: true, service: 'Mitra 911 OnCall server' }); }

function doPost(e) {
  let b;
  try { b = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (x) { return out_({ ok: false, error: 'Format permintaan salah' }); }
  if (!KUNCI || KUNCI === 'GANTI-KATA-SANDI-TIM') return out_({ ok: false, error: 'KUNCI di skrip belum diganti' });
  if (b.key !== KUNCI) return out_({ ok: false, error: 'Kunci tidak cocok' });
  const action = b.action || (b.texts ? 'translate' : 'ping');
  try {
    if (action === 'translate') return out_({ ok: true, translations: translate_(b) });
    if (action === 'ping') {
      const sh = sheet_('Pasien');
      return out_({ ok: true, patients: Math.max(0, sh.getLastRow() - 1), sheetUrl: SpreadsheetApp.getActive().getUrl(), serverTime: Date.now() });
    }
    if (action === 'sync') return out_(sync_(b));
    return out_({ ok: false, error: 'Aksi tidak dikenal' });
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}

/* ================= sinkronisasi ================= */
function sync_(b) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = sheet_('Pasien'), W = SHEETS.Pasien.length;
    const now = Date.now(), dev = String(b.device || '').slice(0, 60);
    const last = sh.getLastRow();
    const idx = last > 1 ? sh.getRange(2, 1, last - 1, 3).getValues() : [];
    const map = {};
    idx.forEach((r, i) => { if (r[0]) map[r[0]] = { row: i + 2, ts: Number(r[1]) || 0, upd: Number(r[2]) || 0 }; });

    const applied = [];
    (b.changes || []).forEach(c => {
      if (!c || !c.id) return;
      const upd = Number(c.updatedAt) || 0, ex = map[c.id];
      if (ex && upd < ex.upd) return;                       // versi server lebih baru
      const row = c.deleted ? tombRow_(c, now, dev) : rowOf_(c, now, dev);
      if (ex) { sh.getRange(ex.row, 1, 1, W).setValues([row]); ex.ts = now; ex.upd = upd; }
      else { sh.appendRow(row); map[c.id] = { row: sh.getLastRow(), ts: now, upd: upd }; }
      applied.push(c);
    });
    if (applied.length) rebuildChildren_(applied);

    const since = Number(b.since) || 0, out = [];
    const want = Object.keys(map).map(k => map[k]).filter(m => m.ts > since);
    if (want.length) {
      const data = sh.getRange(2, 1, sh.getLastRow() - 1, W).getValues();
      const j = P('JSON1'), d = P('Dihapus');
      want.forEach(m => {
        const r = data[m.row - 2];
        if (r[d] === 'YA') out.push({ id: r[0], deleted: true, updatedAt: Number(r[2]) || 0 });
        else { try { out.push(JSON.parse(String(r[j]) + r[j + 1] + r[j + 2] + r[j + 3])); } catch (x) { } }
      });
    }
    if (applied.length || out.length) log_(dev, 'sinkron', applied.length, out.length);
    return { ok: true, serverTime: now, patients: out, applied: applied.length };
  } finally {
    lock.releaseLock();
  }
}

function rowOf_(p, now, dev) {
  const i = p.identity || {}, c = p.complaint || {}, o = p.outcome, t = p.team || {};
  const vs = (p.vitals || []).slice().sort((a, b) => a.t - b.t), lv = vs[vs.length - 1];
  const js = JSON.stringify(p);
  if (js.length > 4 * 45000) throw new Error('Data pasien ' + (i.name || p.id) + ' terlalu besar');
  const ch = [0, 1, 2, 3].map(k => js.substr(k * 45000, 45000));
  return [p.id, now, Number(p.updatedAt) || now, p.status === 'aktif' ? 'Ditangani' : 'Selesai', '', d_(p.createdAt),
    i.name || '', i.passport || '', i.nationality || '', i.sex || '', age_(i), i.hotel || '', i.room || '', i.phone || '',
    i.insurance || '', i.allergy || '', c.chief || '',
    (p.advice || []).map(a => a.dx).filter(String).join('; '), lv ? news_(lv) : '',
    o ? (OUT[o.type] || o.type || '') : '', o ? (o.dest || '') : '', o ? d_(o.t) : '',
    t.doctor || '', t.nurse || '', dev].concat(ch);
}
function tombRow_(c, now, dev) {
  const r = new Array(SHEETS.Pasien.length).fill('');
  r[0] = c.id; r[1] = now; r[2] = Number(c.updatedAt) || now; r[3] = 'Dihapus'; r[4] = 'YA'; r[P('Diubah Oleh')] = dev;
  return r;
}

function rebuildChildren_(list) {
  const ids = {}; list.forEach(p => ids[p.id] = 1);
  const live = list.filter(p => !p.deleted);
  const nm = p => (p.identity && p.identity.name) || '';
  const sets = {
    TTV: live.flatMap(p => (p.vitals || []).map(v => [p.id, nm(p), (p.identity || {}).passport || '', d_(v.t), n_(v.sbp), n_(v.dbp), n_(v.hr), n_(v.rr), n_(v.temp), n_(v.spo2), v.o2 ? 'Ya' : '', v.avpu || '', n_(v.gds), n_(v.pain), news_(v), v.note || '', v.id || ''])),
    Advice: live.flatMap(p => (p.advice || []).map(a => [p.id, nm(p), d_(a.t), a.doctor || '', a.dx || '', a.rx || '', a.plan || '', a.follow || '', a.id || ''])),
    Tindakan: live.flatMap(p => (p.actions || []).map(a => [p.id, nm(p), d_(a.t), a.nurse || '', (a.items || []).join(', '), a.detail || '', a.id || '']))
  };
  Object.keys(sets).forEach(name => {
    const sh = sheet_(name), W = SHEETS[name].length, last = sh.getLastRow();
    const keep = last > 1 ? sh.getRange(2, 1, last - 1, W).getValues().filter(r => r[0] && !ids[r[0]]) : [];
    const rows = keep.concat(sets[name]).sort((a, b) => (time_(a) - time_(b)));
    if (last > 1) sh.getRange(2, 1, last - 1, W).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, W).setValues(rows);
  });
}
function time_(r) { for (let k = 0; k < r.length; k++) if (r[k] instanceof Date) return r[k].getTime(); return 0; }

/* ================= terjemahan ================= */
function translate_(b) {
  return (b.texts || []).slice(0, 300).map(t => {
    t = String(t || '');
    return t.trim() ? LanguageApp.translate(t, b.source || 'id', b.target || 'en') : '';
  });
}

/* ================= utilitas ================= */
function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  const def = SHEETS[name];
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, def.length).setValues([def.map(h => h[0])]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#17398C');
  sh.setFrozenRows(1);
  const rows = sh.getMaxRows() - 1;
  def.forEach((h, k) => sh.getRange(2, k + 1, rows, 1).setNumberFormat(h[1]));
  if (name === 'Pasien') {
    const j = P('JSON1') + 1;
    sh.hideColumns(2, 2); sh.hideColumns(j, 4);
  }
  sh.autoResizeColumns(1, def.length);
  return sh;
}
function log_(dev, aksi, kirim, terima) {
  const sh = sheet_('Log');
  sh.appendRow([new Date(), dev || '-', aksi, kirim, terima]);
  if (sh.getLastRow() > 5000) sh.deleteRows(2, 1000);
}
function d_(ms) { return ms ? new Date(Number(ms)) : ''; }
function n_(v) { return v === '' || v === null || v === undefined || isNaN(Number(v)) ? '' : Number(v); }
function age_(i) {
  if (i.dob) { const d = new Date(i.dob), n = new Date(); let a = n.getFullYear() - d.getFullYear(); if (n < new Date(n.getFullYear(), d.getMonth(), d.getDate())) a--; return isNaN(a) ? '' : a; }
  return n_(i.age);
}
function news_(v) {
  const f = { rr: x => x <= 8 ? 3 : x <= 11 ? 1 : x <= 20 ? 0 : x <= 24 ? 2 : 3, spo2: x => x <= 91 ? 3 : x <= 93 ? 2 : x <= 95 ? 1 : 0, sbp: x => x <= 90 ? 3 : x <= 100 ? 2 : x <= 110 ? 1 : x <= 219 ? 0 : 3, hr: x => x <= 40 ? 3 : x <= 50 ? 1 : x <= 90 ? 0 : x <= 110 ? 1 : x <= 130 ? 2 : 3, temp: x => x <= 35 ? 3 : x <= 36 ? 1 : x <= 38 ? 0 : x <= 39 ? 1 : 2 };
  let t = 0, n = 0;
  Object.keys(f).forEach(k => { const x = n_(v[k]); if (x !== '') { t += f[k](x); n++; } });
  if (!n) return '';
  if (v.o2) t += 2;
  if (v.avpu && v.avpu !== 'A') t += 3;
  return t;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ================= jalankan sekali ================= */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(SHEETS).forEach(sheet_);
  ['Sheet1', 'Lembar1', 'Sheet 1'].forEach(n => { const s = ss.getSheetByName(n); if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s); });
  Logger.log(LanguageApp.translate('Pasien mual dan muntah', 'id', 'en'));
  Logger.log('Siap. Spreadsheet: ' + ss.getUrl());
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Mitra 911').addItem('Siapkan / perbaiki sheet', 'setup').addToUi();
}
