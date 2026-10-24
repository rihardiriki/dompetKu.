/** DompetKu — backend Google Apps Script (database: Google Sheets) */

const SHEETS = {
  Users: ['username', 'salt', 'hash', 'dibuat'],
  Sesi: ['token_hash', 'username', 'kadaluarsa'],
  Transaksi: ['id', 'username', 'tanggal', 'jenis', 'jumlah', 'kategori', 'catatan', 'dibuat'],
  Anggaran: ['username', 'kategori', 'jumlah']
};
const CATS = {
  out: ['Makan', 'Transport', 'Belanja', 'Tagihan', 'Hiburan', 'Kesehatan', 'Lainnya'],
  in: ['Gaji', 'Bonus', 'Usaha', 'Lainnya']
};
const SESSION_DAYS = 30;

function doGet() {
  setup_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('DompetKu')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Helper ---------- */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) { return ss_().getSheetByName(name); }

function setup_() {
  Object.keys(SHEETS).forEach(function (name) {
    if (sh_(name)) return;
    const s = ss_().insertSheet(name);
    s.appendRow(SHEETS[name]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, SHEETS[name].length).setFontWeight('bold');
    if (name === 'Users') s.getRange('A:A').setNumberFormat('@');
    if (name === 'Transaksi') s.getRange('B:C').setNumberFormat('@');
    if (name === 'Anggaran') s.getRange('A:B').setNumberFormat('@');
  });
}

function sha_(text) {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8)
  );
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function cleanUser_(u) { return String(u || '').trim().toLowerCase(); }

/* ---------- Akun & sesi ---------- */

function register(username, password) {
  username = cleanUser_(username);
  password = String(password || '');
  if (!/^[a-z0-9_.]{3,20}$/.test(username)) {
    throw new Error('Nama pengguna 3–20 karakter: huruf, angka, titik, atau garis bawah.');
  }
  if (password.length < 6) throw new Error('Kata sandi minimal 6 karakter.');

  return withLock_(function () {
    const users = sh_('Users');
    const rows = users.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === username) throw new Error('Nama pengguna sudah dipakai.');
    }
    const salt = Utilities.getUuid();
    users.appendRow([username, salt, sha_(salt + password), new Date()]);
    return newSession_(username);
  });
}

function login(username, password) {
  username = cleanUser_(username);
  password = String(password || '');
  const cache = CacheService.getScriptCache();
  const key = 'gagal_' + username;
  const gagal = Number(cache.get(key) || 0);
  if (gagal >= 5) throw new Error('Terlalu banyak percobaan. Coba lagi dalam 10 menit.');

  const rows = sh_('Users').getDataRange().getValues();
  let row = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === username) { row = rows[i]; break; }
  }
  if (!row || sha_(row[1] + password) !== row[2]) {
    cache.put(key, String(gagal + 1), 600);
    throw new Error('Nama pengguna atau kata sandi salah.');
  }
  cache.remove(key);
  return withLock_(function () {
    purgeSessions_();
    return newSession_(username);
  });
}

function newSession_(username) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  const exp = new Date(Date.now() + SESSION_DAYS * 86400000);
  sh_('Sesi').appendRow([sha_(token), username, exp]);
  return { token: token, username: username };
}

function purgeSessions_() {
  const s = sh_('Sesi');
  const rows = s.getDataRange().getValues();
  const now = new Date();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (new Date(rows[i][2]) < now) s.deleteRow(i + 1);
  }
}

// Mengembalikan username pemilik token, atau melempar error 'SESI'
function auth_(token) {
  if (!token) throw new Error('SESI');
  const h = sha_(token);
  const cache = CacheService.getScriptCache();
  const hit = cache.get('s_' + h);
  if (hit) return hit;
  const rows = sh_('Sesi').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === h && new Date(rows[i][2]) > new Date()) {
      cache.put('s_' + h, String(rows[i][1]), 600);
      return String(rows[i][1]);
    }
  }
  throw new Error('SESI');
}

function logout(token) {
  if (!token) return true;
  const h = sha_(token);
  CacheService.getScriptCache().remove('s_' + h);
  withLock_(function () {
    const s = sh_('Sesi');
    const rows = s.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (rows[i][0] === h) s.deleteRow(i + 1);
    }
  });
  return true;
}

/* ---------- Transaksi (setiap pengguna hanya melihat datanya sendiri) ---------- */

function getData(token) {
  const user = auth_(token);
  const tz = ss_().getSpreadsheetTimeZone();
  const rows = sh_('Transaksi').getDataRange().getValues();
  const tx = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[1]) !== user) continue;
    tx.push({
      id: String(r[0]),
      tanggal: r[2] instanceof Date ? Utilities.formatDate(r[2], tz, 'yyyy-MM-dd') : String(r[2]),
      jenis: String(r[3]),
      jumlah: Number(r[4]),
      kategori: String(r[5]),
      catatan: String(r[6] || '')
    });
  }
  const budgets = {};
  sh_('Anggaran').getDataRange().getValues().slice(1).forEach(function (r) {
    if (String(r[0]) === user) budgets[String(r[1])] = Number(r[2]);
  });
  return { username: user, tx: tx, cats: CATS, budgets: budgets };
}

function addTx(token, d) {
  const user = auth_(token);
  if (d.jenis !== 'in' && d.jenis !== 'out') throw new Error('Jenis transaksi tidak valid.');
  const jumlah = Math.round(Number(d.jumlah));
  if (!(jumlah > 0)) throw new Error('Jumlah harus lebih dari 0.');
  if (CATS[d.jenis].indexOf(d.kategori) === -1) throw new Error('Kategori tidak valid.');
  const tanggal = /^\d{4}-\d{2}-\d{2}$/.test(d.tanggal)
    ? d.tanggal
    : Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const catatan = String(d.catatan || '').trim().slice(0, 60);

  return withLock_(function () {
    const id = String(Date.now());
    sh_('Transaksi').appendRow([id, user, tanggal, d.jenis, jumlah, d.kategori, catatan, new Date()]);
    return id;
  });
}

function deleteTx(token, id) {
  const user = auth_(token);
  return withLock_(function () {
    const s = sh_('Transaksi');
    const rows = s.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(id) && String(rows[i][1]) === user) {
        s.deleteRow(i + 1);
        return true;
      }
    }
    throw new Error('Transaksi tidak ditemukan.');
  });
}

/* ---------- Anggaran bulanan per kategori pengeluaran ---------- */

// jumlah = 0 menghapus anggaran kategori tersebut
function setBudget(token, kategori, jumlah) {
  const user = auth_(token);
  if (CATS.out.indexOf(kategori) === -1) throw new Error('Kategori tidak valid.');
  jumlah = Math.round(Number(jumlah) || 0);
  if (jumlah < 0) throw new Error('Jumlah tidak valid.');
  return withLock_(function () {
    const s = sh_('Anggaran');
    const rows = s.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === user && String(rows[i][1]) === kategori) {
        if (jumlah === 0) s.deleteRow(i + 1); else s.getRange(i + 1, 3).setValue(jumlah);
        return true;
      }
    }
    if (jumlah > 0) s.appendRow([user, kategori, jumlah]);
    return true;
  });
}
