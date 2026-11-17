/**
 * Personel Takip — Google E-Tablo sunucusu (Apps Script web uygulaması).
 * Bütün kayıtlar bu betiğin bağlı olduğu E-Tablonun sayfalarında durur.
 * Uygulamalar buraya JSON gönderir: { kod, hafta, tz, ops: [...] }
 * Cevap: { ok, rol, sonuclar, veri } — veri, o kodun görmeye yetkili olduğu kısım.
 */

// Yönetici kodu yalnızca Google'daki Ayar.js dosyasında durur (YONETICI_KODU).
// Bu dosya telefonlardaki programa da "ortak.js" olarak kopyalanır; içinde gizli bilgi olmamalı.
function yoneticiKodu() { return typeof YONETICI_KODU !== 'undefined' ? YONETICI_KODU : ''; }

var SAYFALAR = {
  Personel: ['id', 'ad', 'saatUcreti', 'kod', 'aktif'],
  Saatler: ['pid', 'hafta', 'gun', 'giris', 'cikis', 'onayGiris', 'onayCikis', 'elle', 'guncelleme'],
  Kiyafetler: ['id', 'pid', 'hafta', 'ad', 'fiyat', 'indirimli', 'onay', 'eklendi'],
  Haftalar: ['pid', 'hafta', 'odendi', 'saatUcreti']
};

// ---------------- Google tarafı ----------------

// İlk kurulumda düzenleyiciden bir kez çalıştırılır: izni verir ve tablonun sayfalarını oluşturur.
function kurulum() {
  dbYukle();
  Logger.log('Hazır. Personel Takip tablosu kuruldu.');
}

function doGet() {
  return cevap({ ok: true, mesaj: 'Personel Takip sunucusu çalışıyor.' });
}

function doPost(e) {
  var kilit = LockService.getScriptLock();
  kilit.waitLock(25000);
  try {
    var istek = JSON.parse(e.postData.contents);
    var db = dbYukle();
    var sonuc = isle(db, istek, googleSimdi);
    dbKaydet(db);
    return cevap(sonuc);
  } catch (err) {
    return cevap({ ok: false, hata: String((err && err.message) || err) });
  } finally {
    kilit.releaseLock();
  }
}

function cevap(nesne) {
  return ContentService.createTextOutput(JSON.stringify(nesne)).setMimeType(ContentService.MimeType.JSON);
}

function googleSimdi(tz) {
  var bolge = tz || Session.getScriptTimeZone();
  var s = Utilities.formatDate(new Date(), bolge, "yyyy-MM-dd'|'HH:mm'|'u").split('|');
  return { tarih: s[0], saat: s[1], gunNo: Number(s[2]) - 1 }; // 0 = Pazartesi
}

function dbYukle() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var db = { tablolar: {}, kirli: {} };
  Object.keys(SAYFALAR).forEach(function (ad) {
    var basliklar = SAYFALAR[ad];
    var sayfa = ss.getSheetByName(ad);
    if (!sayfa) {
      sayfa = ss.insertSheet(ad);
      sayfa.getRange(1, 1, sayfa.getMaxRows(), basliklar.length).setNumberFormat('@');
      sayfa.getRange(1, 1, 1, basliklar.length).setValues([basliklar]).setFontWeight('bold');
      sayfa.setFrozenRows(1);
    }
    var degerler = sayfa.getDataRange().getDisplayValues();
    db.tablolar[ad] = degerler.slice(1).filter(function (r) { return r.join('') !== ''; }).map(function (r) {
      var o = {};
      basliklar.forEach(function (b, i) { o[b] = r[i] == null ? '' : String(r[i]); });
      return o;
    });
  });
  var bos = ss.getSheetByName('Sayfa1') || ss.getSheetByName('Sheet1');
  if (bos && ss.getSheets().length > 1 && bos.getLastRow() === 0) ss.deleteSheet(bos);
  return db;
}

function dbKaydet(db) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(db.kirli).forEach(function (ad) {
    var basliklar = SAYFALAR[ad];
    var sayfa = ss.getSheetByName(ad);
    var satirlar = db.tablolar[ad].map(function (o) { return basliklar.map(function (b) { return String(o[b] == null ? '' : o[b]); }); });
    var eski = sayfa.getLastRow();
    if (eski > 1) sayfa.getRange(2, 1, eski - 1, basliklar.length).clearContent();
    if (satirlar.length) {
      if (sayfa.getMaxRows() < satirlar.length + 1) sayfa.insertRowsAfter(sayfa.getMaxRows(), satirlar.length + 1 - sayfa.getMaxRows());
      var alan = sayfa.getRange(2, 1, satirlar.length, basliklar.length);
      alan.setNumberFormat('@');
      alan.setValues(satirlar);
    }
  });
}

// ---------------- İş mantığı (Google'dan bağımsız, test edilebilir) ----------------

var SAAT_RE = /^([01]\d|2[0-3]):[0-5]\d$/;
var HAFTA_RE = /^\d{4}-\d{2}-\d{2}$/;
var KOD_HARFLER = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

function hata(m) { throw new Error(m); }
function normKod(k) { return String(k || '').trim().toUpperCase().replace(/\s+/g, ''); }
function rastgeleKod(db, uzunluk) {
  var kod;
  do {
    kod = '';
    for (var i = 0; i < uzunluk; i++) kod += KOD_HARFLER.charAt(Math.floor(Math.random() * KOD_HARFLER.length));
  } while (db.tablolar.Personel.some(function (p) { return p.kod === kod; }));
  return kod;
}
function yeniId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

function haftaKontrol(h) { if (!HAFTA_RE.test(String(h || ''))) hata('Geçersiz hafta.'); return String(h); }
function gunKontrol(g) { var n = Number(g); if (!(n >= 0 && n <= 6) || n !== Math.floor(n)) hata('Geçersiz gün.'); return String(n); }
function saatKontrol(s) { s = String(s || ''); if (s && !SAAT_RE.test(s)) hata('Geçersiz saat: ' + s); return s; }
function fiyatKontrol(f) { var n = Number(String(f).replace(',', '.')); if (!(n >= 0) || n > 100000) hata('Geçersiz fiyat.'); return String(Math.round(n * 100) / 100); }
function metin(s, en) { return String(s == null ? '' : s).trim().slice(0, en || 80); }

function pazartesiOf(tarih, gunNo) {
  var p = tarih.split('-').map(Number);
  var d = new Date(Date.UTC(p[0], p[1] - 1, p[2] - gunNo));
  return d.toISOString().slice(0, 10);
}

function kirlet(db, ad) { db.kirli[ad] = true; }

function personelBul(db, id) { return db.tablolar.Personel.filter(function (p) { return p.id === id; })[0]; }
function haftaSatiri(db, pid, hafta, olustur) {
  var h = db.tablolar.Haftalar.filter(function (x) { return x.pid === pid && x.hafta === hafta; })[0];
  if (!h && olustur) {
    var p = personelBul(db, pid);
    h = { pid: pid, hafta: hafta, odendi: '0', saatUcreti: p ? p.saatUcreti : '' };
    db.tablolar.Haftalar.push(h); kirlet(db, 'Haftalar');
  }
  return h;
}
function odendiMi(db, pid, hafta) { var h = haftaSatiri(db, pid, hafta, false); return !!(h && h.odendi === '1'); }
function saatSatiri(db, pid, hafta, gun, olustur) {
  var s = db.tablolar.Saatler.filter(function (x) { return x.pid === pid && x.hafta === hafta && x.gun === gun; })[0];
  if (!s && olustur) {
    s = { pid: pid, hafta: hafta, gun: gun, giris: '', cikis: '', onayGiris: '', onayCikis: '', elle: '0', guncelleme: '' };
    db.tablolar.Saatler.push(s);
  }
  if (olustur) kirlet(db, 'Saatler');
  return s;
}
function kiyafetBul(db, id) { return db.tablolar.Kiyafetler.filter(function (k) { return k.id === id; })[0]; }

function kimlik(db, kod) {
  kod = normKod(kod);
  if (!kod) return null;
  if (yoneticiKodu() && kod === normKod(yoneticiKodu())) return { rol: 'yonetici' };
  var p = db.tablolar.Personel.filter(function (x) { return x.kod === kod && x.aktif !== '0'; })[0];
  return p ? { rol: 'personel', pid: p.id } : null;
}

function isle(db, istek, simdiFn) {
  var kim = kimlik(db, istek && istek.kod);
  if (!kim) return { ok: false, kodHatasi: true, hata: 'Kod geçersiz.' };
  var sonuclar = [];
  (istek.ops || []).forEach(function (op) {
    try {
      var r = kim.rol === 'yonetici' ? yoneticiOp(db, op, istek, simdiFn) : personelOp(db, kim.pid, op, istek, simdiFn);
      sonuclar.push({ opId: op.opId, ok: true, sonuc: r || null });
    } catch (err) {
      sonuclar.push({ opId: op.opId, ok: false, hata: String(err.message || err) });
    }
  });
  var hafta = HAFTA_RE.test(String(istek.hafta || '')) ? String(istek.hafta) : null;
  return { ok: true, rol: kim.rol, sonuclar: sonuclar, veri: gorunum(db, kim, hafta) };
}

function gorunum(db, kim, hafta) {
  var T = db.tablolar;
  var bu = function (x) { return !hafta || x.hafta === hafta; };
  if (kim.rol === 'yonetici') {
    // Onay bekleyenler bütün haftalardan sayılır ki eski haftada kalan unutulmasın
    var bekleyen = {};
    T.Saatler.forEach(function (s) { if (s.giris !== s.onayGiris || s.cikis !== s.onayCikis) bekleyen[s.pid + '|' + s.hafta] = (bekleyen[s.pid + '|' + s.hafta] || 0) + 1; });
    T.Kiyafetler.forEach(function (k) { if (k.onay !== '1') bekleyen[k.pid + '|' + k.hafta] = (bekleyen[k.pid + '|' + k.hafta] || 0) + 1; });
    return {
      personel: T.Personel,
      saatler: T.Saatler.filter(bu),
      kiyafetler: T.Kiyafetler.filter(bu),
      haftalar: T.Haftalar.filter(bu),
      bekleyen: bekleyen
    };
  }
  var p = personelBul(db, kim.pid);
  var mi = function (x) { return x.pid === kim.pid && bu(x); };
  return {
    ben: { id: p.id, ad: p.ad },
    saatler: T.Saatler.filter(mi).map(function (s) { return { pid: s.pid, hafta: s.hafta, gun: s.gun, giris: s.giris, cikis: s.cikis, onayGiris: s.onayGiris, onayCikis: s.onayCikis, elle: s.elle }; }),
    kiyafetler: T.Kiyafetler.filter(mi).map(function (k) { return { id: k.id, pid: k.pid, hafta: k.hafta, ad: k.ad, fiyat: k.fiyat, indirimli: k.indirimli, onay: k.onay }; }),
    haftalar: T.Haftalar.filter(mi).map(function (h) { return { pid: h.pid, hafta: h.hafta, odendi: h.odendi }; })
  };
}

function personelOp(db, pid, op, istek, simdiFn) {
  var simdi = simdiFn(istek.tz);
  if (op.tur === 'damga') {
    var hafta = pazartesiOf(simdi.tarih, simdi.gunNo), gun = String(simdi.gunNo);
    if (op.yon === 'geldim') {
      if (odendiMi(db, pid, hafta)) hata('Bu hafta ödendi, değiştirilemez.');
      var s = saatSatiri(db, pid, hafta, gun, true);
      if (s.giris && !s.cikis) hata('Bugün zaten giriş yapılmış (' + s.giris + ').');
      if (s.giris && s.cikis) hata('Bugünün giriş ve çıkışı zaten var. Değiştirmek için saati elle düzelt.');
      s.giris = simdi.saat; s.guncelleme = simdi.tarih + ' ' + simdi.saat;
      return { saat: simdi.saat };
    }
    if (op.yon === 'ciktim') {
      var bugun = saatSatiri(db, pid, hafta, gun, false);
      var hedef = bugun && bugun.giris && !bugun.cikis ? bugun : null;
      if (!hedef) {
        // Gece yarısını geçen vardiya: dünün açık girişi
        var dHafta = simdi.gunNo === 0 ? pazartesiOf(simdi.tarih, 7) : hafta;
        var dGun = String(simdi.gunNo === 0 ? 6 : simdi.gunNo - 1);
        var dun = saatSatiri(db, pid, dHafta, dGun, false);
        if (dun && dun.giris && !dun.cikis) hedef = dun;
      }
      if (!hedef) hata('Önce "Geldim" tuşuna basılmalı.');
      if (odendiMi(db, pid, hedef.hafta)) hata('Bu hafta ödendi, değiştirilemez.');
      kirlet(db, 'Saatler');
      hedef.cikis = simdi.saat; hedef.guncelleme = simdi.tarih + ' ' + simdi.saat;
      return { saat: simdi.saat };
    }
    hata('Geçersiz işlem.');
  }
  if (op.tur === 'saat') {
    var h = haftaKontrol(op.hafta), g = gunKontrol(op.gun);
    if (odendiMi(db, pid, h)) hata('Bu hafta ödendi, değiştirilemez.');
    var giris = saatKontrol(op.giris), cikis = saatKontrol(op.cikis);
    var r = saatSatiri(db, pid, h, g, true);
    if (r.giris === giris && r.cikis === cikis) return null;
    r.giris = giris; r.cikis = cikis; r.elle = '1';
    r.guncelleme = simdi.tarih + ' ' + simdi.saat;
    return null;
  }
  if (op.tur === 'kiyafet') {
    var hk = haftaKontrol(op.hafta);
    if (odendiMi(db, pid, hk)) hata('Bu hafta ödendi, değiştirilemez.');
    var ad = metin(op.ad); if (!ad) hata('Ürün adı boş olamaz.');
    var id = metin(op.id, 40) || yeniId();
    var k = kiyafetBul(db, id);
    if (k && (k.pid !== pid || k.onay === '1')) hata('Onaylanmış kıyafet değiştirilemez.');
    if (!k) { k = { id: id, pid: pid, hafta: hk, eklendi: simdi.tarih + ' ' + simdi.saat }; db.tablolar.Kiyafetler.push(k); }
    k.ad = ad; k.fiyat = fiyatKontrol(op.fiyat); k.indirimli = op.indirimli ? '1' : '0'; k.onay = '0'; k.hafta = hk;
    kirlet(db, 'Kiyafetler');
    return { id: id };
  }
  if (op.tur === 'kiyafetSil') {
    var ks = kiyafetBul(db, op.id);
    if (!ks) return null;
    if (ks.pid !== pid || ks.onay === '1') hata('Onaylanmış kıyafet silinemez.');
    db.tablolar.Kiyafetler = db.tablolar.Kiyafetler.filter(function (x) { return x !== ks; });
    kirlet(db, 'Kiyafetler');
    return null;
  }
  hata('Bu işleme yetkin yok.');
}

function yoneticiOp(db, op, istek, simdiFn) {
  var T = db.tablolar, simdi = simdiFn(istek.tz);
  switch (op.tur) {
    case 'personel': {
      var p = op.id ? personelBul(db, op.id) : null;
      if (!p) { p = { id: metin(op.id, 40) || yeniId(), ad: '', saatUcreti: '', kod: rastgeleKod(db, 6), aktif: '1' }; T.Personel.push(p); }
      if (op.ad !== undefined) p.ad = metin(op.ad);
      if (op.saatUcreti !== undefined) p.saatUcreti = op.saatUcreti === '' ? '' : fiyatKontrol(op.saatUcreti);
      kirlet(db, 'Personel');
      return { id: p.id, kod: p.kod };
    }
    case 'kodYenile': {
      var pk = personelBul(db, op.id); if (!pk) hata('Personel bulunamadı.');
      pk.kod = rastgeleKod(db, 6); kirlet(db, 'Personel');
      return { kod: pk.kod };
    }
    case 'personelSil': {
      ['Personel', 'Saatler', 'Kiyafetler', 'Haftalar'].forEach(function (ad) {
        T[ad] = T[ad].filter(function (x) { return (ad === 'Personel' ? x.id : x.pid) !== op.id; });
        kirlet(db, ad);
      });
      return null;
    }
    case 'saat': {
      var h = haftaKontrol(op.hafta), g = gunKontrol(op.gun);
      if (!personelBul(db, op.pid)) hata('Personel bulunamadı.');
      var yg = saatKontrol(op.giris), yc = saatKontrol(op.cikis);
      var s = saatSatiri(db, op.pid, h, g, true);
      s.giris = s.onayGiris = yg;
      s.cikis = s.onayCikis = yc;
      s.elle = '0'; s.guncelleme = simdi.tarih + ' ' + simdi.saat;
      return null;
    }
    case 'ucret': {
      var hu = haftaSatiri(db, op.pid, haftaKontrol(op.hafta), true);
      var u = op.saatUcreti === '' ? '' : fiyatKontrol(op.saatUcreti);
      hu.saatUcreti = u;
      var pu = personelBul(db, op.pid); if (pu) { pu.saatUcreti = u; kirlet(db, 'Personel'); }
      kirlet(db, 'Haftalar');
      return null;
    }
    case 'kiyafet': {
      var hk = haftaKontrol(op.hafta);
      var id = metin(op.id, 40) || yeniId();
      var k = kiyafetBul(db, id);
      if (!k) { k = { id: id, pid: op.pid, hafta: hk, eklendi: simdi.tarih + ' ' + simdi.saat }; T.Kiyafetler.push(k); }
      k.ad = metin(op.ad); k.fiyat = op.fiyat === '' || op.fiyat == null ? '' : fiyatKontrol(op.fiyat);
      k.indirimli = op.indirimli ? '1' : '0'; k.onay = '1';
      kirlet(db, 'Kiyafetler');
      return { id: id };
    }
    case 'kiyafetSil': {
      T.Kiyafetler = T.Kiyafetler.filter(function (x) { return x.id !== op.id; });
      kirlet(db, 'Kiyafetler');
      return null;
    }
    case 'onayla':
    case 'reddet': {
      var onay = op.tur === 'onayla';
      var hafta = haftaKontrol(op.hafta);
      var tumu = op.gun === undefined && op.kiyafetId === undefined;
      T.Saatler.forEach(function (x) {
        if (x.pid !== op.pid || x.hafta !== hafta) return;
        if (!tumu && (op.gun === undefined || String(op.gun) !== x.gun)) return;
        if (onay) { x.onayGiris = x.giris; x.onayCikis = x.cikis; } else { x.giris = x.onayGiris; x.cikis = x.onayCikis; }
        x.elle = '0';
      });
      kirlet(db, 'Saatler');
      var hedefKiy = function (x) { return x.pid === op.pid && x.hafta === hafta && x.onay !== '1' && (tumu || x.id === op.kiyafetId); };
      if (onay) T.Kiyafetler.forEach(function (x) { if (hedefKiy(x)) x.onay = '1'; });
      else T.Kiyafetler = T.Kiyafetler.filter(function (x) { return !hedefKiy(x); });
      kirlet(db, 'Kiyafetler');
      return null;
    }
    case 'odendi': {
      var ho = haftaSatiri(db, op.pid, haftaKontrol(op.hafta), true);
      ho.odendi = op.deger ? '1' : '0';
      if (!ho.saatUcreti) { var po = personelBul(db, op.pid); if (po) ho.saatUcreti = po.saatUcreti; }
      kirlet(db, 'Haftalar');
      return null;
    }
    case 'iceAktar': {
      // Eski (telefonda/bilgisayarda tutulan) kayıtları bir kez tabloya taşır
      var v = op.veri || {};
      (v.personel || []).forEach(function (ep) {
        var id = metin(ep.id, 40); if (!id || personelBul(db, id)) return;
        T.Personel.push({ id: id, ad: metin(ep.ad), saatUcreti: ep.saatUcreti === '' || ep.saatUcreti == null ? '' : fiyatKontrol(ep.saatUcreti), kod: rastgeleKod(db, 6), aktif: '1' });
      });
      Object.keys(v.haftalar || {}).forEach(function (hafta) {
        if (!HAFTA_RE.test(hafta)) return;
        var kayitlar = v.haftalar[hafta];
        Object.keys(kayitlar).forEach(function (pid) {
          if (!personelBul(db, pid)) return;
          var kr = kayitlar[pid];
          var hs = haftaSatiri(db, pid, hafta, true);
          hs.odendi = kr.odendi ? '1' : '0';
          if (kr.saatUcreti !== undefined && kr.saatUcreti !== '') hs.saatUcreti = fiyatKontrol(kr.saatUcreti);
          Object.keys(kr.gunler || {}).forEach(function (g) {
            var gg = kr.gunler[g] || {};
            var gi = SAAT_RE.test(gg.giris || '') ? gg.giris : '', ci = SAAT_RE.test(gg.cikis || '') ? gg.cikis : '';
            if (!gi && !ci) return;
            var s = saatSatiri(db, pid, hafta, String(Number(g)), true);
            s.giris = s.onayGiris = gi; s.cikis = s.onayCikis = ci; s.elle = '0';
          });
          (kr.kiyafetler || []).forEach(function (ek) {
            var kid = metin(ek.id, 40) || yeniId(); if (kiyafetBul(db, kid)) return;
            T.Kiyafetler.push({ id: kid, pid: pid, hafta: hafta, ad: metin(ek.ad), fiyat: ek.fiyat === '' || ek.fiyat == null ? '' : fiyatKontrol(ek.fiyat), indirimli: ek.indirimli ? '1' : '0', onay: '1', eklendi: '' });
          });
        });
      });
      ['Personel', 'Saatler', 'Kiyafetler', 'Haftalar'].forEach(function (ad) { kirlet(db, ad); });
      return null;
    }
  }
  hata('Bilinmeyen işlem: ' + op.tur);
}

// Node testleri için (Google'da bu satırın etkisi yok)
if (typeof module !== 'undefined') module.exports = { isle: isle, kimlik: kimlik, pazartesiOf: pazartesiOf, SAYFALAR: SAYFALAR, ayarla: function (k) { global.YONETICI_KODU = k; } };
