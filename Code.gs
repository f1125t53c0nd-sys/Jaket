// ============================================================
// Code.gs - Survei Jaket
// Membutuhkan: App.gs (SPREADSHEET_URL, SHEET_NAME, APP_TITLE),
//              Form.html, Hasil.html
// ============================================================

const HEADERS = ["Timestamp", "Nama Panggilan", "Brand", "Jenis Jaket", "Model/Tipe",
  "Lama Memiliki", "Rentang Harga", "Alasan Memilih", "Faktor Penting",
  "Frekuensi Pakai", "Frekuensi Beli", "Tempat Beli", "Kepuasan", "Beli Lagi"];

const SKOR_PAKAI = { "Setiap hari": 5, "Beberapa kali seminggu": 4, "Seminggu sekali": 3,
  "Beberapa kali sebulan": 2, "Jarang sekali": 1 };
const LABEL_PAKAI = ["-", "Jarang sekali", "Beberapa kali sebulan", "Seminggu sekali",
  "Beberapa kali seminggu", "Setiap hari"];

// Menjalankan web app. ?page=hasil -> halaman hasil, selain itu -> form
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || "form";
  const t = HtmlService.createTemplateFromFile(page === "hasil" ? "Hasil" : "Form");
  t.baseUrl = ScriptApp.getService().getUrl();
  t.judul = APP_TITLE;
  return t.evaluate().setTitle(APP_TITLE)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// Membuka spreadsheet & sheet (membuat header otomatis jika kosong)
function getSheet_() {
  if (SPREADSHEET_URL.indexOf("MASUKKAN") === 0) {
    throw new Error("SPREADSHEET_URL di App.gs belum diisi.");
  }
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#ffd6e7");
    sh.setFrozenRows(1);
  }
  return sh;
}

// Jalankan sekali dari editor untuk memberi izin & membuat sheet
function setupSheet() {
  getSheet_();
  Logger.log("Sheet siap: " + SHEET_NAME);
}

function bersih_(s) {
  s = String(s || "").trim().substring(0, 60);
  return /^[=+\-@]/.test(s) ? "'" + s : s; // cegah formula injection
}

// Dipanggil dari Form.html
function simpanJawaban(d) {
  const wajib = ["brand", "jenis", "lama", "harga", "pakai", "beli", "kepuasan", "belilagi"];
  for (const k of wajib) {
    if (!d[k]) return { ok: false, pesan: "Mohon lengkapi semua pertanyaan bertanda *" };
  }
  const nilai = Number(d.kepuasan);
  if (!(nilai >= 1 && nilai <= 5)) return { ok: false, pesan: "Nilai kepuasan tidak valid." };
  const gabung = v => (Array.isArray(v) ? v : []).map(x => String(x).replace(/;/g, ",")).join("; ");

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    getSheet_().appendRow([new Date(), bersih_(d.nama), bersih_(d.brand), bersih_(d.jenis),
      bersih_(d.model), bersih_(d.lama), bersih_(d.harga), gabung(d.alasan), gabung(d.faktor),
      bersih_(d.pakai), bersih_(d.beli), gabung(d.tempat), nilai, bersih_(d.belilagi)]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// Dipanggil dari Hasil.html: statistik terbaru
function ambilStatistik() {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return { total: 0 };
  const rows = sh.getRange(2, 1, n, HEADERS.length).getValues();
  const I = h => HEADERS.indexOf(h);
  const one = h => rows.map(r => r[I(h)]);
  const many = h => rows.reduce((a, r) => a.concat(String(r[I(h)]).split(";")), []);
  const rata = a => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0);
  const cap = s => s.charAt(0).toUpperCase() + s.slice(1);

  const hitung = (vals, base) => {
    const m = {}, nama = {};
    vals.forEach(v => {
      const s = String(v).trim();
      if (!s) return;
      const k = s.toLowerCase();
      m[k] = (m[k] || 0) + 1;
      if (!nama[k]) nama[k] = cap(s);
    });
    return Object.keys(m)
      .map(k => ({ label: nama[k], jumlah: m[k], persen: Math.round(m[k] / base * 100) }))
      .sort((a, b) => b.jumlah - a.jumlah);
  };

  const puas = one("Kepuasan").map(Number).filter(x => x > 0);
  const skor = one("Frekuensi Pakai").map(v => SKOR_PAKAI[v] || 0).filter(x => x > 0);
  const ya = rows.filter(r => r[I("Beli Lagi")] === "Ya, pasti");

  const perBrand = {};
  rows.forEach(r => {
    const b = String(r[I("Brand")]).trim();
    if (!b) return;
    const k = b.toLowerCase();
    (perBrand[k] = perBrand[k] || { label: cap(b), s: [] }).s.push(Number(r[I("Kepuasan")]) || 0);
  });
  const brandPuas = Object.values(perBrand)
    .map(o => ({ label: o.label, jumlah: o.s.length, rata: Math.round(rata(o.s) * 10) / 10 }))
    .sort((a, b) => b.jumlah - a.jumlah).slice(0, 8);

  const rataPakai = rata(skor);
  return {
    total: n,
    rataKepuasan: Math.round(rata(puas) * 10) / 10,
    rataPakai: Math.round(rataPakai * 10) / 10,
    labelPakai: LABEL_PAKAI[Math.round(rataPakai)] || "-",
    persenBeliLagi: Math.round(ya.length / n * 100),
    brand: hitung(one("Brand"), n),
    jenis: hitung(one("Jenis Jaket"), n),
    alasan: hitung(many("Alasan Memilih"), n),
    harga: hitung(one("Rentang Harga"), n),
    faktor: hitung(many("Faktor Penting"), n),
    tempat: hitung(many("Tempat Beli"), n),
    pakai: hitung(one("Frekuensi Pakai"), n),
    beli: hitung(one("Frekuensi Beli"), n),
    lama: hitung(one("Lama Memiliki"), n),
    beliLagiBrand: hitung(ya.map(r => r[I("Brand")]), Math.max(ya.length, 1)),
    brandPuas: brandPuas,
    diperbarui: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm")
  };
}
