/** SISTEM ADMINISTRASI KKI INSTIPER — Google Apps Script
 *  Pasang dari Google Sheet kosong: Extensions > Apps Script. Lalu jalankan setup() sekali. */
const TZ = 'Asia/Jakarta';
const TEMPLATE_DOC_ID_EN = '15slBC3hKSi1-4bS0GHAO8bjEIU2QP0O6Y5pkVSsyG7s'; // ← ID Google Doc template surat berbahasa Inggris
const TEMPLATE_DOC_ID = '1XvrZKL1s0ciR8kVgwiq9VyhfKtm08i8LbCS0FbZT2NQ'; // ← tempel ID Google Doc template surat (lihat panduan)
const P = PropertiesService.getScriptProperties();
const ROMAN = ['I','II','III','IV','V','VI','VII','VIII','IX','X','XI','XII'];
const BLN = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const ROLES = ['Admin','Pimpinan','Kontributor'];
const SESS_TTL = 21600;
const DEFAULT_CFG = {
  nomor_format: '{urut}/{kode}/{bulan_romawi}/{tahun}',
  nomor_kode: 'KKI/SB',
  nomor_reset: 'bulanan',
  nomor_pad: 3,
  nomor_mulai: 1,
  template_id: '',
  template_en: '',
  reminder_aktif: 'Ya',
  reminder_email: '',
  reminder_jam: 8,
  reminder_hari: '3,1,0',
  reminder_weekend: 'Tidak',
  reminder_mou_hari: '180,90,60,30,7',
  telegram_aktif: 'Tidak',
  telegram_token: '',
  telegram_chat_id: ''
};
const SHEETS = {
  pengguna: ['username','password_hash','role'],
  surat_masuk: ['id_surat_masuk','nomor_surat_asal','nomor_disposisi','tanggal_terima','pengirim','perihal','perlu_balasan','tenggat_balasan','status','file_scan_url'],
  surat_keluar: ['id_surat_keluar','nomor_surat','tanggal_surat','bahasa','tujuan_surat','perihal','ref_surat_masuk','status_ttd','file_docx_url','file_pdf_signed_url','isi','lampiran'],
  berita_acara_mom: ['id_dokumen','jenis_dokumen','nama_kegiatan','mitra_terlibat','tanggal_kegiatan','poin_pembahasan','foto_dokumentasi','file_generated_url'],
  reminder_followup: ['id_task','sumber_rujukan','what_deskripsi','who_pic','email_pic','when_deadline','status','catatan_progres'],
  portal_mou: ['id_mou','jenis_naskah','nama_mitra','negara','unit_pengusul','tanggal_mulai','tanggal_berakhir','status_dokumen','file_naskah_url']
};

/* ---------- SETUP (jalankan sekali) ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  P.setProperty('SS', ss.getId());
  Object.keys(SHEETS).forEach(n => {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (n === 'pengguna') {
      const h = s.getRange(1, 1, 1, Math.max(1, s.getLastColumn())).getValues()[0];
      if (h[0] !== 'username' || h[1] !== 'password_hash' || h[2] !== 'role') s.clearContents();
    }
    s.getRange(1, 1, 1, SHEETS[n].length).setValues([SHEETS[n]]).setFontWeight('bold');
    s.setFrozenRows(1);
  });
  const d = ss.getSheetByName('Sheet1'); if (d && ss.getSheets().length > 1) ss.deleteSheet(d);
  if (!P.getProperty('ROOT')) P.setProperty('ROOT', DriveApp.createFolder('KKI INSTIPER - Arsip Digital').getId());
  ensureUsersReady_();
  pasangTrigger();
  Logger.log('Setup selesai. Login default: admin / admin123 (segera ganti password setelah masuk).');
}
function pasangTrigger(hour) {
  const h = Math.min(23, Math.max(0, parseInt(hour, 10) || 8));
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'cekHarian') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('cekHarian').timeBased().atHour(h).nearMinute(0).everyDays(1).inTimezone(TZ).create();
}

/* ---------- WEB APP ---------- */
function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  const b = P.getProperty('SS') ? JSON.stringify({ ready: true }) : JSON.stringify({ err: 'Spreadsheet belum terhubung. Jalankan setup() di Apps Script.' });
  t.boot = b.replace(/</g, '\\u003c');
  return t.evaluate().setTitle('KKI INSTIPER')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- HELPER ---------- */
const ss_ = () => SpreadsheetApp.openById(P.getProperty('SS'));
const sh_ = n => ss_().getSheetByName(n);
function rows_(n) {
  const v = sh_(n).getDataRange().getValues(), h = v.shift();
  return v.filter(r => r[0] !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i] instanceof Date ? Utilities.formatDate(r[i], TZ, 'yyyy-MM-dd') : r[i]); return o; });
}
function add_(n, o) { sh_(n).appendRow(SHEETS[n].map(k => o[k] === undefined ? '' : o[k])); }
function upd_(n, idc, id, o) {
  const s = sh_(n), v = s.getDataRange().getValues(), h = v[0], c = h.indexOf(idc);
  for (let i = 1; i < v.length; i++) if (v[i][c] == id) { Object.keys(o).forEach(k => s.getRange(i + 1, h.indexOf(k) + 1).setValue(o[k])); return true; }
  return false;
}
function del_(n, idc, id) {
  const s = sh_(n), v = s.getDataRange().getValues(), c = v[0].indexOf(idc);
  let cnt = 0;
  for (let i = v.length - 1; i >= 1; i--) if (v[i][c] == id) { s.deleteRow(i + 1); cnt++; }
  return cnt;
}
function trashUrl_(url) {
  const id = extractDriveId_(url);
  if (!id) return;
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
}
function trashPrevPdf_(id, dtStr) {
  try {
    const y = new Date(dtStr || new Date()).getFullYear();
    const it = folder_(['Surat Keluar', y, '_preview']).getFilesByName(id + '.pdf');
    while (it.hasNext()) it.next().setTrashed(true);
  } catch (e) {}
}
function nextId_(prefix, n, col) {
  const y = new Date().getFullYear(), re = new RegExp('^' + prefix + '-' + y + '-(\\d+)$');
  const max = Math.max(0, ...rows_(n).map(r => { const m = re.exec(r[col]); return m ? +m[1] : 0; }));
  return prefix + '-' + y + '-' + String(max + 1).padStart(3, '0');
}
function sha256_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function hashPw_(pw) {
  const p = String(pw || '');
  if (p.length < 4) throw new Error('Password minimal 4 karakter.');
  const salt = Utilities.getUuid().slice(0, 12);
  return salt + '$' + sha256_(salt + ':' + p);
}
function verifyPw_(pw, stored) {
  const parts = String(stored || '').split('$');
  if (parts.length !== 2 || !parts[0] || !parts[1]) return false;
  return sha256_(parts[0] + ':' + String(pw || '')) === parts[1];
}
function normUser_(u) {
  const s = String(u || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,32}$/.test(s)) throw new Error('Username harus 3-32 karakter (huruf, angka, titik, _ atau -).');
  return s;
}
function ensureUsersReady_() {
  const s = sh_('pengguna');
  const h = s.getRange(1, 1, 1, Math.max(1, s.getLastColumn())).getValues()[0];
  if (h[0] !== 'username' || h[1] !== 'password_hash' || h[2] !== 'role') {
    lock_(() => {
      const h2 = s.getRange(1, 1, 1, Math.max(1, s.getLastColumn())).getValues()[0];
      if (h2[0] !== 'username' || h2[1] !== 'password_hash' || h2[2] !== 'role') {
        s.clearContents();
        s.getRange(1, 1, 1, SHEETS.pengguna.length).setValues([SHEETS.pengguna]).setFontWeight('bold');
        s.setFrozenRows(1);
      }
    });
  }
  if (!rows_('pengguna').length) {
    lock_(() => {
      if (!rows_('pengguna').length) add_('pengguna', { username: 'admin', password_hash: hashPw_('admin123'), role: 'Admin' });
    });
  }
}
function me_(tok) {
  const t = String(tok || '').trim();
  if (!t) throw new Error('AUTH_REQUIRED');
  const cache = CacheService.getScriptCache();
  const raw = cache.get('SESS_' + t);
  if (!raw) throw new Error('AUTH_REQUIRED');
  let sess;
  try { sess = JSON.parse(raw); } catch (e) { throw new Error('AUTH_REQUIRED'); }
  const u = sess && sess.username ? rows_('pengguna').find(x => String(x.username || '').trim().toLowerCase() === String(sess.username).toLowerCase()) : null;
  if (!u) { cache.remove('SESS_' + t); throw new Error('AUTH_REQUIRED'); }
  const role = ROLES.includes(String(u.role || '').trim()) ? String(u.role).trim() : 'Kontributor';
  cache.put('SESS_' + t, JSON.stringify({ username: String(u.username).trim(), role }), SESS_TTL);
  return { username: String(u.username).trim(), role };
}
function need_(tok, roles) { const m = me_(tok); if (!roles.includes(m.role)) throw new Error('Akses ditolak untuk peran ' + m.role); return m; }
function folder_(path) {
  let f = DriveApp.getFolderById(P.getProperty('ROOT'));
  path.forEach(n => { const it = f.getFoldersByName(String(n)); f = it.hasNext() ? it.next() : f.createFolder(String(n)); });
  return f;
}
function save_(f, path, name) {
  name = name.replace(/[\/\\:*?"<>|]/g, '-');
  return folder_(path).createFile(Utilities.newBlob(Utilities.base64Decode(f.data), f.type || 'application/pdf', name)).getUrl();
}
const days_ = d => { const t = new Date(); t.setHours(0, 0, 0, 0); return Math.round((new Date(d) - t) / 864e5); };
function statusMou_(end) { const d = days_(end); return d < 0 ? 'Expired' : d <= 180 ? 'Akan Berakhir (<6 Bulan)' : 'Aktif'; }
function lock_(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }

/* ---------- AUTENTIKASI & DATA UNTUK UI ---------- */
function login(username, password) {
  if (!P.getProperty('SS')) throw new Error('Spreadsheet belum terhubung. Jalankan setup() di Apps Script.');
  ensureUsersReady_();
  const uname = String(username || '').trim().toLowerCase();
  const u = uname ? rows_('pengguna').find(x => String(x.username || '').trim().toLowerCase() === uname) : null;
  if (!u || !verifyPw_(password, u.password_hash)) throw new Error('Username atau password salah.');
  const role = ROLES.includes(String(u.role || '').trim()) ? String(u.role).trim() : 'Kontributor';
  const tok = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  const me = { username: String(u.username).trim(), role };
  CacheService.getScriptCache().put('SESS_' + tok, JSON.stringify(me), SESS_TTL);
  return { token: tok, data: buildData_(me) };
}
function logout(tok) {
  const t = String(tok || '').trim();
  if (t) CacheService.getScriptCache().remove('SESS_' + t);
}
function buildData_(m) {
  const isA = m.role === 'Admin', full = isA || m.role === 'Pimpinan';
  const tasks = full ? rows_('reminder_followup').map(t => (t.when_deadline && t.status !== 'Done' && days_(t.when_deadline) < 0) ? Object.assign(t, { overdue: true }) : t) : [];
  let mou = rows_('portal_mou');
  if (!isA) mou = mou.filter(x => x.status_dokumen !== 'Draf Review');
  const users = isA ? rows_('pengguna').map(u => ({ username: String(u.username || '').trim(), role: ROLES.includes(String(u.role || '').trim()) ? String(u.role).trim() : 'Kontributor' })) : [];
  const cfg = isA ? getConfig_() : null;
  const next_nomor = isA ? nextNomor_().nomor : '';
  return { me: m, masuk: full ? rows_('surat_masuk').reverse() : [], keluar: full ? rows_('surat_keluar').reverse() : [], tasks, mou: mou.reverse(), users, cfg, next_nomor };
}
function getData(tok) {
  return buildData_(me_(tok));
}

/* ---------- SURAT MASUK ---------- */
function simpanSuratMasuk(tok, d, f) {
  const m = need_(tok, ['Admin']);
  return lock_(() => {
    const id = nextId_('SM', 'surat_masuk', 'id_surat_masuk'), t = d.tanggal_terima || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const url = f ? save_(f, ['Arsip Surat Masuk', t.slice(0, 4), t.slice(5, 7)], t + '_' + d.pengirim + '_' + d.perihal.slice(0, 40) + '.pdf') : '';
    add_('surat_masuk', Object.assign({}, d, { id_surat_masuk: id, status: 'Belum Diproses', file_scan_url: url }));
    if (d.perlu_balasan === 'Ya') {
      add_('reminder_followup', { id_task: nextId_('TASK', 'reminder_followup', 'id_task'), sumber_rujukan: id, what_deskripsi: 'Balas surat dari ' + d.pengirim + ': ' + d.perihal, who_pic: m.username, email_pic: '', when_deadline: d.tenggat_balasan, status: 'Pending' });
      try {
        const cfg = getConfig_();
        if (cfg.telegram_aktif === 'Ya' && cfg.telegram_token && cfg.telegram_chat_id) {
          const app = ScriptApp.getService().getUrl();
          const msg = '📬 <b>SURAT MASUK BARU (Perlu Balasan)</b>\n\n' +
            '• <b>ID:</b> ' + id + '\n' +
            (d.nomor_surat_asal ? '• <b>No. Asal:</b> ' + d.nomor_surat_asal + '\n' : '') +
            '• <b>Pengirim:</b> ' + d.pengirim + '\n' +
            '• <b>Perihal:</b> ' + d.perihal + '\n' +
            '• <b>Tenggat Balasan:</b> ' + d.tenggat_balasan + '\n' +
            '• <b>PIC:</b> ' + m.username + '\n\n' +
            '<a href="' + app + '">Buka Sistem Administrasi KKI</a>';
          kirimTelegram_(cfg.telegram_token, cfg.telegram_chat_id, msg);
        }
      } catch (te) { Logger.log('Telegram notif error: ' + te); }
    }
    return id;
  });
}
function editSuratMasuk(tok, id, d, f) {
  need_(tok, ['Admin']);
  return lock_(() => {
    const r = rows_('surat_masuk').find(x => x.id_surat_masuk === id);
    if (!r) throw new Error('Surat masuk tidak ditemukan.');
    const t = d.tanggal_terima || r.tanggal_terima || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const upd = {
      pengirim: d.pengirim || '', nomor_surat_asal: d.nomor_surat_asal || '', nomor_disposisi: d.nomor_disposisi || '',
      tanggal_terima: t, perihal: d.perihal || '', perlu_balasan: d.perlu_balasan || 'Tidak',
      tenggat_balasan: d.tenggat_balasan || '', status: d.status || r.status
    };
    if (f) {
      upd.file_scan_url = save_(f, ['Arsip Surat Masuk', t.slice(0, 4), t.slice(5, 7)], t + '_' + (d.pengirim || 'surat') + '_' + String(d.perihal || '').slice(0, 40) + '.pdf');
      trashUrl_(r.file_scan_url);
    }
    upd_('surat_masuk', 'id_surat_masuk', id, upd);
    return id;
  });
}
function hapusSuratMasuk(tok, id) {
  need_(tok, ['Admin']);
  return lock_(() => {
    const r = rows_('surat_masuk').find(x => x.id_surat_masuk === id);
    if (!r) throw new Error('Surat masuk tidak ditemukan.');
    trashUrl_(r.file_scan_url);
    del_('surat_masuk', 'id_surat_masuk', id);
    del_('reminder_followup', 'sumber_rujukan', id);
    return id;
  });
}

/* ---------- PENGATURAN & PENOMORAN ---------- */
function getConfig_() {
  let saved = {};
  try { saved = JSON.parse(P.getProperty('CFG') || '{}') || {}; } catch (e) {}
  const cfg = Object.assign({}, DEFAULT_CFG, saved);
  if (!cfg.template_id) cfg.template_id = TEMPLATE_DOC_ID || P.getProperty('TEMPLATE_DOC_ID') || '';
  if (!cfg.template_en) cfg.template_en = TEMPLATE_DOC_ID_EN || P.getProperty('TEMPLATE_DOC_ID_EN') || '';
  cfg.nomor_pad = Math.min(6, Math.max(1, parseInt(cfg.nomor_pad, 10) || 3));
  cfg.nomor_mulai = Math.max(1, parseInt(cfg.nomor_mulai, 10) || 1);
  if (!['bulanan', 'tahunan', 'lanjut'].includes(cfg.nomor_reset)) cfg.nomor_reset = 'bulanan';
  if (!String(cfg.nomor_format || '').includes('{urut}')) cfg.nomor_format = DEFAULT_CFG.nomor_format;
  cfg.reminder_aktif = cfg.reminder_aktif === 'Tidak' ? 'Tidak' : 'Ya';
  cfg.reminder_jam = Math.min(23, Math.max(0, parseInt(cfg.reminder_jam, 10) || 8));
  cfg.reminder_hari = String(cfg.reminder_hari || '3,1,0').trim();
  cfg.reminder_weekend = cfg.reminder_weekend === 'Ya' ? 'Ya' : 'Tidak';
  cfg.reminder_mou_hari = String(cfg.reminder_mou_hari || '180,90,60,30,7').trim();
  cfg.telegram_aktif = cfg.telegram_aktif === 'Ya' ? 'Ya' : 'Tidak';
  cfg.telegram_token = String(cfg.telegram_token || '').trim();
  cfg.telegram_chat_id = String(cfg.telegram_chat_id || '').trim();
  return cfg;
}
function formatNomor_(cfg, urut, dt) {
  const y = dt.getFullYear(), mo = dt.getMonth();
  return String(cfg.nomor_format)
    .replace(/\{urut\}/g, String(urut).padStart(cfg.nomor_pad, '0'))
    .replace(/\{kode\}/g, String(cfg.nomor_kode || '').trim())
    .replace(/\{bulan_romawi\}/g, ROMAN[mo])
    .replace(/\{bulan\}/g, String(mo + 1).padStart(2, '0'))
    .replace(/\{tahun\}/g, String(y));
}
function extractUrut_(nomorStr, cfg) {
  const s = String(nomorStr || '').trim();
  const escRe = v => v.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const pat = escRe(String(cfg.nomor_format))
    .replace('\\{urut\\}', '(\\d+)')
    .replace('\\{kode\\}', '.*?')
    .replace('\\{bulan_romawi\\}', '[IVX]+')
    .replace('\\{bulan\\}', '\\d{1,2}')
    .replace('\\{tahun\\}', '\\d{4}');
  const m = new RegExp('^' + pat + '$').exec(s);
  if (m) return parseInt(m[1], 10) || 0;
  const p = parseInt(s, 10);
  return isNaN(p) ? 0 : p;
}
function nextNomor_(dtStr) {
  const cfg = getConfig_();
  const dt = new Date(dtStr || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
  const y = dt.getFullYear(), mo = dt.getMonth();
  const existing = rows_('surat_keluar').filter(r => {
    if (cfg.nomor_reset === 'lanjut') return true;
    const x = new Date(r.tanggal_surat);
    return cfg.nomor_reset === 'tahunan' ? x.getFullYear() === y : (x.getFullYear() === y && x.getMonth() === mo);
  }).map(r => extractUrut_(r.nomor_surat, cfg));
  const urut = Math.max(cfg.nomor_mulai - 1, 0, ...existing) + 1;
  return { nomor: formatNomor_(cfg, urut, dt), urut, cfg };
}
function simpanPengaturan(tok, d) {
  need_(tok, ['Admin']);
  const fmt = String((d && d.nomor_format) || '').trim();
  if (!fmt.includes('{urut}')) throw new Error('Format nomor surat wajib mengandung placeholder {urut}.');
  const reset = String((d && d.nomor_reset) || 'bulanan').trim();
  if (!['bulanan', 'tahunan', 'lanjut'].includes(reset)) throw new Error('Siklus reset tidak valid.');
  const pad = Math.min(6, Math.max(1, parseInt(d && d.nomor_pad, 10) || 3));
  const mulai = Math.max(1, parseInt(d && d.nomor_mulai, 10) || 1);
  const pickId = v => { const s = String(v || '').trim(); return extractDriveId_(s) || s; };
  const jam = Math.min(23, Math.max(0, parseInt(d && d.reminder_jam, 10) || 8));
  const hari = String((d && d.reminder_hari) || '3,1,0').trim();
  const weekend = d && d.reminder_weekend === 'Ya' ? 'Ya' : 'Tidak';
  const mouHari = String((d && d.reminder_mou_hari) || '180,90,60,30,7').trim();
  return lock_(() => {
    let cur = {};
    try { cur = JSON.parse(P.getProperty('CFG') || '{}') || {}; } catch (e) {}
    const next = Object.assign({}, cur, {
      nomor_format: fmt,
      nomor_kode: String((d && d.nomor_kode) || '').trim(),
      nomor_reset: reset,
      nomor_pad: pad,
      nomor_mulai: mulai,
      template_id: pickId(d && d.template_id),
      template_en: pickId(d && d.template_en),
      reminder_aktif: d && d.reminder_aktif === 'Tidak' ? 'Tidak' : 'Ya',
      reminder_email: String((d && d.reminder_email) || '').trim(),
      reminder_jam: jam,
      reminder_hari: hari,
      reminder_weekend: weekend,
      reminder_mou_hari: mouHari,
      telegram_aktif: d && d.telegram_aktif === 'Ya' ? 'Ya' : 'Tidak',
      telegram_token: String((d && d.telegram_token) || '').trim(),
      telegram_chat_id: String((d && d.telegram_chat_id) || '').trim()
    });
    P.setProperty('CFG', JSON.stringify(next));
    try { pasangTrigger(jam); } catch (te) { Logger.log('Trigger error: ' + te); }
    return { cfg: getConfig_(), next_nomor: nextNomor_().nomor };
  });
}

/* ---------- SURAT KELUAR ---------- */
function ensureTemplateId_(en) {
  const cfg = getConfig_();
  const cfgId = en ? cfg.template_en : cfg.template_id;
  // Coba dari config dulu (Pengaturan)
  if (cfgId) {
    try { const f = DriveApp.getFileById(cfgId); if (!f.isTrashed()) return cfgId; } catch (e) {}
  }
  // Fallback ke konstanta di baris atas
  const constId = en ? TEMPLATE_DOC_ID_EN : TEMPLATE_DOC_ID;
  if (constId) {
    try { const f = DriveApp.getFileById(constId); if (!f.isTrashed()) return constId; } catch (e) {}
  }
  // Fallback ke ScriptProperties lama
  const propKey = en ? 'TEMPLATE_DOC_ID_EN' : 'TEMPLATE_DOC_ID';
  const propId = P.getProperty(propKey);
  if (propId) {
    try { const f = DriveApp.getFileById(propId); if (!f.isTrashed()) return propId; } catch (e) {}
  }
  // Auto-generate template sederhana jika tidak ada yang valid
  Logger.log('Auto-generating ' + (en ? 'EN' : 'ID') + ' letter template...');
  const doc = DocumentApp.create(en ? 'Template Surat KKI (EN)' : 'Template Surat KKI (ID)');
  const b = doc.getBody();
  b.appendParagraph('KANTOR KERJASAMA INTERNASIONAL (KKI)').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  b.appendParagraph('INSTITUT PERTANIAN STIPER (INSTIPER) YOGYAKARTA').editAsText().setBold(true);
  b.appendHorizontalRule();
  if (en) {
    b.appendParagraph('Ref. No.   : {{nomor}}');
    b.appendParagraph('Attachment : {{lampiran}}');
    b.appendParagraph('Subject    : {{perihal}}');
    b.appendParagraph('Date       : {{tanggal}}');
    b.appendParagraph('');
    b.appendParagraph('To:');
    b.appendParagraph('{{tujuan}}');
    b.appendParagraph('di Tempat');
    b.appendParagraph('');
    b.appendParagraph('{{isi}}');
    b.appendParagraph('');
    b.appendParagraph('Sincerely,');
    b.appendParagraph('');
    b.appendParagraph('');
    b.appendParagraph('________________________');
    b.appendParagraph('Kepala KKI INSTIPER');
  } else {
    b.appendParagraph('Nomor    : {{nomor}}');
    b.appendParagraph('Lampiran : {{lampiran}}');
    b.appendParagraph('Perihal  : {{perihal}}');
    b.appendParagraph('Tanggal  : {{tanggal}}');
    b.appendParagraph('');
    b.appendParagraph('Kepada Yth.');
    b.appendParagraph('{{tujuan}}');
    b.appendParagraph('di Tempat');
    b.appendParagraph('');
    b.appendParagraph('{{isi}}');
    b.appendParagraph('');
    b.appendParagraph('Hormat kami,');
    b.appendParagraph('');
    b.appendParagraph('');
    b.appendParagraph('________________________');
    b.appendParagraph('Kepala KKI INSTIPER');
  }
  doc.saveAndClose();
  const newId = doc.getId();
  DriveApp.getFileById(newId).moveTo(folder_(['Template Surat']));
  // Simpan ke properties dan config agar tersimpan permanen
  P.setProperty(propKey, newId);
  try {
    let cur = {};
    try { cur = JSON.parse(P.getProperty('CFG') || '{}') || {}; } catch (e2) {}
    cur[en ? 'template_en' : 'template_id'] = newId;
    P.setProperty('CFG', JSON.stringify(cur));
  } catch (e3) {}
  return newId;
}
function fillSuratDoc_(tplId, name, dir, map) {
  const copy = DriveApp.getFileById(tplId).makeCopy(name, dir);
  try {
    const doc = DocumentApp.openById(copy.getId()), body = doc.getBody();
    Object.keys(map).forEach(k => body.replaceText('\\{\\{' + k + '\\}\\}', String(map[k] || '').replace(/\$/g, '\\$')));
    doc.saveAndClose();
    const docId = copy.getId();
    const hdr = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } };
    return {
      exportAs: (fmt, ext) => UrlFetchApp.fetch('https://docs.google.com/document/d/' + docId + '/export?format=' + fmt, hdr).getBlob().setName(name + '.' + ext),
      cleanup: () => copy.setTrashed(true)
    };
  } catch (e) {
    copy.setTrashed(true);
    throw e;
  }
}
function previewSuratKeluar(tok, d) {
  need_(tok, ['Admin', 'Pimpinan']);
  const en = d && d.bahasa === 'Inggris';
  const dtStr = (d && d.tanggal_surat) || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const dt = new Date(dtStr), y = dt.getFullYear(), mo = dt.getMonth();
  const nomor = (d && d.nomor_surat) || nextNomor_(dtStr).nomor;
  const tgl = en ? Utilities.formatDate(dt, TZ, 'MMMM d, yyyy') : dt.getDate() + ' ' + BLN[mo] + ' ' + y;
  const name = d && d.nomor_surat ? 'Surat ' + d.nomor_surat.replace(/\//g, '-') : 'Preview_Surat';
  const id = d && d.id_surat_keluar ? String(d.id_surat_keluar) : '';
  if (id) {
    const pdir = folder_(['Surat Keluar', y, '_preview']);
    const it = pdir.getFilesByName(id + '.pdf');
    if (it.hasNext()) {
      const f = it.next();
      if (!f.isTrashed()) return { name: name + '.pdf', mime: 'application/pdf', b64: Utilities.base64Encode(f.getBlob().getBytes()) };
    }
  }
  const tplId = ensureTemplateId_(en), dir = folder_(['Template Surat']);
  const h = fillSuratDoc_(tplId, 'TMP_' + Date.now(), dir, { nomor, tanggal: tgl, tujuan: (d && d.tujuan_surat) || '', perihal: (d && d.perihal) || '', isi: (d && d.isi) || '', lampiran: (d && d.lampiran) || '-' });
  let blob;
  try { blob = h.exportAs('pdf', 'pdf'); } finally { h.cleanup(); }
  if (id) { try { folder_(['Surat Keluar', y, '_preview']).createFile(blob.copyBlob().setName(id + '.pdf')); } catch (e) {} }
  return { name: name + '.pdf', mime: 'application/pdf', b64: Utilities.base64Encode(blob.getBytes()) };
}
function buatSuratKeluar(tok, d) {
  need_(tok, ['Admin']);
  return lock_(() => {
    const dtStr = d.tanggal_surat || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const dt = new Date(dtStr), y = dt.getFullYear(), mo = dt.getMonth();
    const { nomor } = nextNomor_(dtStr);
    const id = nextId_('SK', 'surat_keluar', 'id_surat_keluar'), en = d.bahasa === 'Inggris';
    const tplId = ensureTemplateId_(en);
    const tgl = en ? Utilities.formatDate(dt, TZ, 'MMMM d, yyyy') : dt.getDate() + ' ' + BLN[mo] + ' ' + y;
    const name = 'Surat ' + nomor.replace(/\//g, '-'), dir = folder_(['Surat Keluar', y]);
    const h = fillSuratDoc_(tplId, name, dir, { nomor, tanggal: tgl, tujuan: d.tujuan_surat || '', perihal: d.perihal || '', isi: d.isi || '', lampiran: d.lampiran || '-' });
    let url;
    try {
      url = dir.createFile(h.exportAs('docx', 'docx')).getUrl();
      try { folder_(['Surat Keluar', y, '_preview']).createFile(h.exportAs('pdf', 'pdf').setName(id + '.pdf')); } catch (e) {}
    } finally { h.cleanup(); }
    add_('surat_keluar', { id_surat_keluar: id, nomor_surat: nomor, tanggal_surat: dtStr, bahasa: d.bahasa, tujuan_surat: d.tujuan_surat, perihal: d.perihal, ref_surat_masuk: d.ref_surat_masuk, status_ttd: 'Menunggu TTD', file_docx_url: url, isi: d.isi, lampiran: d.lampiran });
    if (d.ref_surat_masuk) upd_('surat_masuk', 'id_surat_masuk', d.ref_surat_masuk, { status: 'Dalam Review' });
    return nomor;
  });
}
function unggahTTD(tok, id, v, f) {
  need_(tok, ['Admin']);
  if (!f) throw new Error('Pilih file PDF bertanda tangan.');
  const r = rows_('surat_keluar').find(x => x.id_surat_keluar === id);
  upd_('surat_keluar', 'id_surat_keluar', id, { status_ttd: 'Terkirim / Selesai', file_pdf_signed_url: save_(f, ['Surat Keluar', String(r.tanggal_surat).slice(0, 4)], 'TTD_' + r.nomor_surat + '.pdf') });
  if (r.ref_surat_masuk) {
    upd_('surat_masuk', 'id_surat_masuk', r.ref_surat_masuk, { status: 'Selesai Dibalas' });
    rows_('reminder_followup').filter(t => t.sumber_rujukan === r.ref_surat_masuk).forEach(t => upd_('reminder_followup', 'id_task', t.id_task, { status: 'Done' }));
  }
}

function editSuratKeluar(tok, id, d) {
  need_(tok, ['Admin']);
  return lock_(() => {
    const r = rows_('surat_keluar').find(x => x.id_surat_keluar === id);
    if (!r) throw new Error('Surat keluar tidak ditemukan.');
    const dtStr = d.tanggal_surat || r.tanggal_surat || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const dt = new Date(dtStr), y = dt.getFullYear(), mo = dt.getMonth();
    const nomor = String(d.nomor_surat || r.nomor_surat || '').trim();
    const en = (d.bahasa || r.bahasa) === 'Inggris';
    const tplId = ensureTemplateId_(en);
    const tgl = en ? Utilities.formatDate(dt, TZ, 'MMMM d, yyyy') : dt.getDate() + ' ' + BLN[mo] + ' ' + y;
    const name = 'Surat ' + nomor.replace(/\//g, '-'), dir = folder_(['Surat Keluar', y]);
    const h = fillSuratDoc_(tplId, name, dir, { nomor, tanggal: tgl, tujuan: d.tujuan_surat || '', perihal: d.perihal || '', isi: d.isi || '', lampiran: d.lampiran || '-' });
    let url;
    try {
      url = dir.createFile(h.exportAs('docx', 'docx')).getUrl();
      trashUrl_(r.file_docx_url);
      trashPrevPdf_(id, r.tanggal_surat);
      try { folder_(['Surat Keluar', y, '_preview']).createFile(h.exportAs('pdf', 'pdf').setName(id + '.pdf')); } catch (e) {}
    } finally { h.cleanup(); }
    const status_ttd = d.status_ttd || r.status_ttd;
    upd_('surat_keluar', 'id_surat_keluar', id, {
      nomor_surat: nomor, tanggal_surat: dtStr, bahasa: d.bahasa || r.bahasa,
      tujuan_surat: d.tujuan_surat || '', perihal: d.perihal || '', lampiran: d.lampiran || '',
      isi: d.isi || '', ref_surat_masuk: d.ref_surat_masuk || '', status_ttd, file_docx_url: url
    });
    if (d.ref_surat_masuk) upd_('surat_masuk', 'id_surat_masuk', d.ref_surat_masuk, { status: status_ttd === 'Terkirim / Selesai' ? 'Selesai Dibalas' : 'Dalam Review' });
    return nomor;
  });
}
function hapusSuratKeluar(tok, id) {
  need_(tok, ['Admin']);
  return lock_(() => {
    const r = rows_('surat_keluar').find(x => x.id_surat_keluar === id);
    if (!r) throw new Error('Surat keluar tidak ditemukan.');
    trashUrl_(r.file_docx_url);
    trashUrl_(r.file_pdf_signed_url);
    trashPrevPdf_(id, r.tanggal_surat);
    del_('surat_keluar', 'id_surat_keluar', id);
    if (r.ref_surat_masuk && !rows_('surat_keluar').some(x => x.ref_surat_masuk === r.ref_surat_masuk)) {
      const sm = rows_('surat_masuk').find(x => x.id_surat_masuk === r.ref_surat_masuk);
      if (sm && sm.status === 'Dalam Review') upd_('surat_masuk', 'id_surat_masuk', r.ref_surat_masuk, { status: 'Belum Diproses' });
    }
    return id;
  });
}

/* ---------- TUGAS & MoU ---------- */
function simpanTask(tok, d) {
  const m = need_(tok, ['Admin']);
  if (!d.what_deskripsi || !d.when_deadline) throw new Error('Deskripsi tugas dan tenggat waktu wajib diisi.');
  return lock_(() => {
    const id = nextId_('TASK', 'reminder_followup', 'id_task');
    add_('reminder_followup', {
      id_task: id,
      sumber_rujukan: d.sumber_rujukan || 'Internal KKI',
      what_deskripsi: d.what_deskripsi,
      who_pic: d.who_pic || m.username,
      email_pic: d.email_pic || '',
      when_deadline: d.when_deadline,
      status: d.status || 'Pending',
      catatan_progres: d.catatan_progres || ''
    });
    try {
      const cfg = getConfig_();
      if (cfg.telegram_aktif === 'Ya' && cfg.telegram_token && cfg.telegram_chat_id) {
        const app = ScriptApp.getService().getUrl();
        const msg = '📌 <b>TUGAS BARU DITAMBAHKAN</b>\n\n' +
          '• <b>ID:</b> ' + id + '\n' +
          '• <b>Tugas:</b> ' + d.what_deskripsi + '\n' +
          '• <b>PIC:</b> ' + (d.who_pic || m.username) + (d.email_pic ? ' (' + d.email_pic + ')' : '') + '\n' +
          '• <b>Tenggat:</b> ' + d.when_deadline + '\n' +
          '• <b>Rujukan:</b> ' + (d.sumber_rujukan || 'Internal KKI') + '\n\n' +
          '<a href="' + app + '">Buka Sistem Administrasi KKI</a>';
        kirimTelegram_(cfg.telegram_token, cfg.telegram_chat_id, msg);
      }
    } catch (e) {}
    return id;
  });
}
function updateTask(tok, id, v) {
  need_(tok, ['Admin']);
  const upd = { status: v.status, catatan_progres: v.catatan_progres || '' };
  if (v.what_deskripsi !== undefined) upd.what_deskripsi = v.what_deskripsi;
  if (v.who_pic !== undefined) upd.who_pic = v.who_pic;
  if (v.email_pic !== undefined) upd.email_pic = v.email_pic;
  if (v.when_deadline !== undefined) upd.when_deadline = v.when_deadline;
  if (v.sumber_rujukan !== undefined) upd.sumber_rujukan = v.sumber_rujukan;
  upd_('reminder_followup', 'id_task', id, upd);
}
function hapusTask(tok, id) {
  need_(tok, ['Admin']);
  return lock_(() => {
    if (!del_('reminder_followup', 'id_task', id)) throw new Error('Tugas tidak ditemukan.');
    return id;
  });
}
function simpanMou(tok, d, f) {
  const m = me_(tok);
  if (!d.nama_mitra || !d.tanggal_mulai || !d.tanggal_berakhir) throw new Error('Mitra, tanggal mulai, dan tanggal berakhir wajib diisi.');
  if (d.tanggal_berakhir < d.tanggal_mulai) throw new Error('Tanggal berakhir tidak boleh sebelum tanggal mulai.');
  if (!f) throw new Error('Unggah file PDF naskah.');
  return lock_(() => {
    const id = nextId_('MOU', 'portal_mou', 'id_mou');
    const url = save_(f, ['Repositori MoU', d.tanggal_mulai.slice(0, 4)], id + '_' + d.nama_mitra + '.pdf');
    add_('portal_mou', Object.assign({}, d, { id_mou: id, status_dokumen: m.role === 'Admin' ? statusMou_(d.tanggal_berakhir) : 'Draf Review', file_naskah_url: url }));
    return id;
  });
}
function editMou(tok, id, d, f) {
  need_(tok, ['Admin']);
  if (!d.nama_mitra || !d.tanggal_mulai || !d.tanggal_berakhir) throw new Error('Mitra, tanggal mulai, dan tanggal berakhir wajib diisi.');
  if (d.tanggal_berakhir < d.tanggal_mulai) throw new Error('Tanggal berakhir tidak boleh sebelum tanggal mulai.');
  return lock_(() => {
    const r = rows_('portal_mou').find(x => x.id_mou === id);
    if (!r) throw new Error('Naskah MoU tidak ditemukan.');
    const upd = {
      jenis_naskah: d.jenis_naskah || r.jenis_naskah, nama_mitra: d.nama_mitra, negara: d.negara || '',
      unit_pengusul: d.unit_pengusul || '', tanggal_mulai: d.tanggal_mulai, tanggal_berakhir: d.tanggal_berakhir,
      status_dokumen: d.status_dokumen === 'Draf Review' ? 'Draf Review' : statusMou_(d.tanggal_berakhir)
    };
    if (f) {
      upd.file_naskah_url = save_(f, ['Repositori MoU', d.tanggal_mulai.slice(0, 4)], id + '_' + d.nama_mitra + '.pdf');
      trashUrl_(r.file_naskah_url);
    }
    upd_('portal_mou', 'id_mou', id, upd);
    return id;
  });
}
function hapusMou(tok, id) {
  need_(tok, ['Admin']);
  return lock_(() => {
    const r = rows_('portal_mou').find(x => x.id_mou === id);
    if (!r) throw new Error('Naskah MoU tidak ditemukan.');
    trashUrl_(r.file_naskah_url);
    del_('portal_mou', 'id_mou', id);
    return id;
  });
}
function verifikasiMou(tok, id) {
  need_(tok, ['Admin']);
  const r = rows_('portal_mou').find(x => x.id_mou === id);
  upd_('portal_mou', 'id_mou', id, { status_dokumen: statusMou_(r.tanggal_berakhir) });
}

/* ---------- MANAJEMEN PENGGUNA ---------- */
function simpanUser(tok, d) {
  need_(tok, ['Admin']);
  const username = normUser_(d && d.username), role = String(d && d.role || '').trim();
  if (!ROLES.includes(role)) throw new Error('Role tidak valid.');
  const password_hash = hashPw_(d && d.password);
  return lock_(() => {
    if (rows_('pengguna').some(u => String(u.username || '').trim().toLowerCase() === username))
      throw new Error('Username sudah terdaftar. Gunakan Ubah untuk memperbarui.');
    add_('pengguna', { username, password_hash, role });
    return username;
  });
}
function updateUser(tok, oldUsername, d) {
  const m = need_(tok, ['Admin']);
  const target = normUser_(oldUsername), username = normUser_(d && d.username), role = String(d && d.role || '').trim();
  if (!ROLES.includes(role)) throw new Error('Role tidak valid.');
  if (target === m.username.toLowerCase() && (role !== 'Admin' || username !== target))
    throw new Error('Tidak dapat mengubah username atau menurunkan role akun Anda sendiri.');
  const newHash = d && d.password ? hashPw_(d.password) : null;
  return lock_(() => {
    const s = sh_('pengguna'), v = s.getDataRange().getValues(), h = v[0];
    const cu = h.indexOf('username'), cp = h.indexOf('password_hash'), cr = h.indexOf('role');
    let idx = -1, admins = 0;
    for (let i = 1; i < v.length; i++) {
      const un = String(v[i][cu] || '').trim().toLowerCase();
      if (!un) continue;
      if (String(v[i][cr] || '').trim() === 'Admin') admins++;
      if (un === target) idx = i;
      else if (un === username) throw new Error('Username baru sudah digunakan pengguna lain.');
    }
    if (idx < 0) throw new Error('Pengguna tidak ditemukan.');
    if (String(v[idx][cr] || '').trim() === 'Admin' && role !== 'Admin' && admins <= 1)
      throw new Error('Minimal harus ada 1 Admin aktif.');
    s.getRange(idx + 1, cu + 1).setValue(username);
    if (newHash) s.getRange(idx + 1, cp + 1).setValue(newHash);
    s.getRange(idx + 1, cr + 1).setValue(role);
    return username;
  });
}
function hapusUser(tok, targetUsername) {
  const m = need_(tok, ['Admin']);
  const target = normUser_(targetUsername);
  if (target === m.username.toLowerCase()) throw new Error('Tidak dapat menghapus akun Anda sendiri.');
  return lock_(() => {
    const s = sh_('pengguna'), v = s.getDataRange().getValues(), h = v[0];
    const cu = h.indexOf('username'), cr = h.indexOf('role');
    let idx = -1, admins = 0;
    for (let i = 1; i < v.length; i++) {
      const un = String(v[i][cu] || '').trim().toLowerCase();
      if (!un) continue;
      if (String(v[i][cr] || '').trim() === 'Admin') admins++;
      if (un === target) idx = i;
    }
    if (idx < 0) throw new Error('Pengguna tidak ditemukan.');
    if (String(v[idx][cr] || '').trim() === 'Admin' && admins <= 1)
      throw new Error('Minimal harus ada 1 Admin aktif.');
    s.deleteRow(idx + 1);
    return target;
  });
}
function gantiPassword(tok, d) {
  const m = me_(tok);
  const oldPw = String(d && d.old_password || ''), newHash = hashPw_(d && d.new_password);
  return lock_(() => {
    const s = sh_('pengguna'), v = s.getDataRange().getValues(), h = v[0];
    const cu = h.indexOf('username'), cp = h.indexOf('password_hash');
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][cu] || '').trim().toLowerCase() === m.username.toLowerCase()) {
        if (!verifyPw_(oldPw, v[i][cp])) throw new Error('Password lama salah.');
        s.getRange(i + 1, cp + 1).setValue(newHash);
        return true;
      }
    }
    throw new Error('Pengguna tidak ditemukan.');
  });
}

/* ---------- PREVIEW DOKUMEN ---------- */
function extractDriveId_(u) {
  const s = String(u || '');
  const m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/) || s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  return m ? m[1] : '';
}
function previewDokumen(tok, url) {
  const m = me_(tok), isA = m.role === 'Admin', full = isA || m.role === 'Pimpinan';
  const fileId = extractDriveId_(url);
  if (!fileId) throw new Error('URL dokumen tidak valid.');
  const allowed = new Set();
  rows_('portal_mou').forEach(x => {
    if (!isA && x.status_dokumen === 'Draf Review') return;
    const id = extractDriveId_(x.file_naskah_url);
    if (id) allowed.add(id);
  });
  if (full) {
    rows_('surat_masuk').forEach(x => { const id = extractDriveId_(x.file_scan_url); if (id) allowed.add(id); });
    rows_('surat_keluar').forEach(x => {
      [x.file_docx_url, x.file_pdf_signed_url].forEach(u => { const id = extractDriveId_(u); if (id) allowed.add(id); });
    });
  }
  if (!allowed.has(fileId)) throw new Error('Akses pratinjau dokumen ditolak.');
  const f = DriveApp.getFileById(fileId);
  const name = f.getName(), mime = f.getMimeType(), size = f.getSize();
  const previewUrl = 'https://drive.google.com/file/d/' + fileId + '/preview';
  const downloadUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;
  const canInline = (mime === 'application/pdf' || mime.indexOf('image/') === 0) && size <= 10485760;
  const b64 = canInline ? Utilities.base64Encode(f.getBlob().getBytes()) : '';
  return { name, mime, size, b64, previewUrl, downloadUrl, viewUrl: url };
}

/* ---------- REMINDER HARIAN 08.00 WIB & TELEGRAM ---------- */
function kirimTelegram_(token, chatId, text) {
  if (!token || !chatId || !text) return { ok: false, err: 'Token, Chat ID, atau pesan kosong.' };
  try {
    const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
    const payload = {
      chat_id: chatId,
      text: text,
      parse_mode: 'HTML',
      disable_web_page_preview: true
    };
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    let body = {};
    try { body = JSON.parse(res.getContentText()); } catch (e) {}
    if (code === 200 && body.ok) return { ok: true };
    return { ok: false, err: body.description || ('HTTP error ' + code) };
  } catch (e) {
    return { ok: false, err: e.message || String(e) };
  }
}

function cekUpdateTelegram(tok, inputTok) {
  need_(tok, ['Admin']);
  const cfg = getConfig_();
  const token = String(inputTok || cfg.telegram_token || '').trim();
  if (!token) throw new Error('Isi Bot Token Telegram terlebih dahulu.');
  const url = 'https://api.telegram.org/bot' + token + '/getUpdates';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Token bot tidak valid atau tidak dapat terhubung ke Telegram.');
  let body = {};
  try { body = JSON.parse(res.getContentText()); } catch (e) {}
  if (!body.ok || !body.result || !body.result.length) {
    throw new Error('Belum ada pesan yang masuk ke bot. Buka bot Anda di Telegram (@' + (body.result && body.result.username || 'bot') + ') dan klik START, atau ketik pesan di grup yang sudah ada bot tersebut.');
  }
  for (let i = body.result.length - 1; i >= 0; i--) {
    const u = body.result[i];
    const m = u.message || u.channel_post || u.my_chat_member;
    if (m && m.chat && m.chat.id) {
      const c = m.chat;
      const title = c.title || ((c.first_name || '') + ' ' + (c.last_name || '')).trim() || 'Chat';
      return { chat_id: String(c.id), title, type: c.type };
    }
  }
  throw new Error('Tidak ditemukan chat aktif pada bot ini.');
}

function testTelegram(tok, inputTok, inputCid) {
  need_(tok, ['Admin']);
  const cfg = getConfig_();
  const token = String(inputTok || cfg.telegram_token || '').trim();
  const chatId = String(inputCid || cfg.telegram_chat_id || '').trim();
  if (!token || !chatId) {
    throw new Error('Bot Token dan Chat ID Telegram wajib diisi terlebih dahulu.');
  }
  const text = '<b>✅ Notifikasi Telegram KKI INSTIPER Aktif</b>\n\n' +
    'Sistem Administrasi & Kerjasama Internasional KKI INSTIPER berhasil terhubung dengan bot ini.\n\n' +
    'Pengingat harian dan tugas surat masuk akan dikirimkan ke chat/grup ini.';
  const r = kirimTelegram_(token, chatId, text);
  if (!r.ok) {
    throw new Error('Telegram: ' + r.err + (r.err.includes('chat not found') ? ' (Pastikan Chat ID benar dan bot sudah Anda klik /start, atau sudah diundang ke grup)' : ''));
  }
  return true;
}

function kirimLaporanHarianTelegram_(cfg, appUrl, taskAlertDays, mouAlertDays) {
  if (!cfg.telegram_token || !cfg.telegram_chat_id) return;
  const tDays = taskAlertDays || [3, 2, 1, 0];
  const mDays = mouAlertDays || [180, 90, 60, 30, 7];

  const tasks = rows_('reminder_followup').filter(t => {
    if (t.status === 'Done' || !t.when_deadline) return false;
    const d = days_(t.when_deadline);
    return tDays.includes(d) || d < 0;
  });
  const expiringMou = rows_('portal_mou').filter(x => {
    if (!['Aktif', 'Akan Berakhir (<6 Bulan)'].includes(x.status_dokumen)) return false;
    const d = days_(x.tanggal_berakhir);
    return mDays.includes(d) || (d < 0 && d >= -14);
  });

  if (!tasks.length && !expiringMou.length) return;

  let msg = '<b>🔔 PENGINGAT HARIAN KKI INSTIPER</b>\n' +
    '<i>' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') + '</i>\n\n';

  if (tasks.length) {
    msg += '<b>📋 Tugas Tindak Lanjut Perlu Perhatian (' + tasks.length + '):</b>\n';
    tasks.forEach(t => {
      const d = days_(t.when_deadline);
      const tag = d < 0 ? '🔴 [TERLAMBAT ' + (-d) + ' hari]' : (d === 0 ? '🟡 [HARI INI]' : '🟡 [H-' + d + ']');
      msg += tag + ' <b>' + t.what_deskripsi + '</b>\n' +
        '  • PIC: ' + t.who_pic + (t.email_pic ? ' (' + t.email_pic + ')' : '') + ' · Tenggat: ' + t.when_deadline + '\n' +
        '  • Rujukan: ' + t.sumber_rujukan + '\n';
    });
    msg += '\n';
  }

  if (expiringMou.length) {
    msg += '<b>📑 Naskah Kerja Sama Mendekati Berakhir (' + expiringMou.length + '):</b>\n';
    expiringMou.forEach(m => {
      const d = days_(m.tanggal_berakhir);
      const tag = d < 0 ? '🔴 [Kedaluwarsa]' : '⚠️ [Sisa ' + d + ' hari]';
      msg += tag + ' ' + m.jenis_naskah + ' - ' + m.nama_mitra + '\n' +
        '  • Berakhir: ' + m.tanggal_berakhir + ' · Pengusul: ' + m.unit_pengusul + '\n';
    });
    msg += '\n';
  }

  msg += '<a href="' + appUrl + '">Buka Sistem Administrasi KKI</a>';
  kirimTelegram_(cfg.telegram_token, cfg.telegram_chat_id, msg);
}

function cekHarian() {
  const cfg = getConfig_();
  if (cfg.reminder_aktif === 'Tidak' && cfg.telegram_aktif === 'Tidak') return;
  const wd = new Date().getDay();
  if (cfg.reminder_weekend === 'Tidak' && (wd === 0 || wd === 6)) return;
  const app = ScriptApp.getService().getUrl();

  const taskAlertDays = String(cfg.reminder_hari || '3,1,0')
    .split(',')
    .map(x => parseInt(x.trim(), 10))
    .filter(x => !isNaN(x));
  const mouAlertDays = String(cfg.reminder_mou_hari || '180,90,60,30,7')
    .split(',')
    .map(x => parseInt(x.trim(), 10))
    .filter(x => !isNaN(x));

  if (cfg.reminder_aktif === 'Ya') {
    rows_('reminder_followup').forEach(t => {
      if (t.status === 'Done' || !t.when_deadline || !t.email_pic) return;
      const d = days_(t.when_deadline);
      if (taskAlertDays.includes(d) || d < 0) {
        MailApp.sendEmail(t.email_pic, (d < 0 ? '[OVERDUE] ' : '[KKI] ') + 'Pengingat tugas: ' + t.what_deskripsi,
          'Tugas: ' + t.what_deskripsi + '\nPIC: ' + t.who_pic + '\nTenggat: ' + t.when_deadline + (d < 0 ? ' (terlambat ' + (-d) + ' hari)' : ' (H-' + d + ')') + '\nRujukan: ' + t.sumber_rujukan + '\nBuka aplikasi: ' + app);
      }
    });
    const admins = String(cfg.reminder_email || '').trim() || Session.getEffectiveUser().getEmail();
    rows_('portal_mou').forEach(x => {
      if (!['Aktif', 'Akan Berakhir (<6 Bulan)'].includes(x.status_dokumen)) return;
      const ns = statusMou_(x.tanggal_berakhir);
      if (ns !== x.status_dokumen) upd_('portal_mou', 'id_mou', x.id_mou, { status_dokumen: ns });
      const d = days_(x.tanggal_berakhir);
      if (admins && mouAlertDays.includes(d)) {
        MailApp.sendEmail(admins, '[KKI] ' + x.jenis_naskah + ' dengan ' + x.nama_mitra + ' berakhir ' + d + ' hari lagi', 'Berakhir: ' + x.tanggal_berakhir + '\nSiapkan draf pembaruan (renewal).\nNaskah: ' + x.file_naskah_url);
      }
    });
  }

  if (cfg.telegram_aktif === 'Ya') {
    kirimLaporanHarianTelegram_(cfg, app, taskAlertDays, mouAlertDays);
  }
}
