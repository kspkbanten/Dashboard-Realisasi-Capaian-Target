/**
 * DASHBOARD PROGRAM PRIORITAS PEMBANGUNAN KELUARGA
 * Google Sheets -> Google Apps Script Web App -> GitHub Pages
 * READ ONLY: endpoint hanya melayani HTTP GET.
 */

const SPREADSHEET_ID = '1B9IM4H4Mi5ToMW8oJAZceQV2NLMEh0H8';

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const payload = buildPayload_(ss);

    // JSONP untuk dashboard GitHub Pages.
    // Mendukung ?prefix=namaFungsi dan ?callback=namaFungsi.
    const params = (e && e.parameter) ? e.parameter : {};
    const prefix = String(params.prefix || params.callback || '').trim();

    if (prefix) {
      // Hanya izinkan nama fungsi JavaScript yang valid.
      if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(prefix)) {
        return ContentService
          .createTextOutput('Invalid prefix')
          .setMimeType(ContentService.MimeType.TEXT);
      }

      return ContentService
        .createTextOutput(prefix + '(' + JSON.stringify(payload) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    // Akses /exec biasa tetap mengembalikan JSON.
    return json_(payload);

  } catch (err) {
    const errorPayload = {
      success: false,
      error: String(err && err.message ? err.message : err)
    };

    const params = (e && e.parameter) ? e.parameter : {};
    const prefix = String(params.prefix || params.callback || '').trim();

    if (prefix && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(prefix)) {
      return ContentService
        .createTextOutput(prefix + '(' + JSON.stringify(errorPayload) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return json_(errorPayload);
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildPayload_(ss) {
  const tamasya = readRows_(ss, 'TAMASYA', 5, 8, [2, 3, 4, 6, 7, 10, 11], function(v) {
    return {
      kab: normalizeKab_(v[0]),
      jumlah: num_(v[1]),
      dampingan: num_(v[2]),
      target4: num_(v[3]),
      capaian4: num_(v[4]),
      inisiasiTarget: num_(v[5]),
      inisiasiCapaian: num_(v[6])
    };
  });

  const hpk = readRows_(ss, '1000 HPK', 8, 8, [2, 3, 4, 6, 7, 8, 9, 10, 11, 12], function(v) {
    return {
      kab: normalizeKab_(v[0]),
      targetBumil: num_(v[1]),
      targetBaduta: num_(v[2]),
      jan: num_(v[3]),
      feb: num_(v[4]),
      mar: num_(v[5]),
      apr: num_(v[6]),
      mei: num_(v[7]),
      jun: num_(v[8]),
      jul: num_(v[9])
    };
  });

  const gati = readRows_(ss, 'GATI', 6, 8, [2, 3, 4, 5, 6, 7], function(v) {
    return {
      kab: normalizeKab_(v[0]),
      target: num_(v[1]),
      kompak: num_(v[2]),
      sebaya: num_(v[3]),
      dekat: num_(v[4]),
      totalCapaian: num_(v[5])
    };
  });

  const sidaya = {
    bkl: readSimpleBlock_(ss, 'SIDAYA', 7, 8),
    lansia: readSimpleBlock_(ss, 'SIDAYA', 23, 8),
    kader: readSimpleBlock_(ss, 'SIDAYA', 39, 8),
    sekolah: readSimpleBlock_(ss, 'SIDAYA', 53, 8)
  };

  const satyagatra = {
    terjangkau: readSimpleBlock_(ss, 'SATYAGATRA', 7, 8),
    konseling: readSimpleBlock_(ss, 'SATYAGATRA', 23, 8)
  };

  const uppka = readRows_(ss, 'UPPKA', 8, 8, [3, 4, 5, 6, 7], function(v) {
    return {
      kab: normalizeKab_(v[0]),
      jumlah: num_(v[1]),
      nib: num_(v[2]),
      target2026: num_(v[3]),
      capaian2026: num_(v[4])
    };
  });

  const pklKecamatan = readAllRows_(ss, 'PKL_KECAMATAN');
  const slKecamatan = readAllRows_(ss, 'SL_KECAMATAN');

  const spreadsheetTitle = ss.getName();
  const compiledDate = extractDataDateFromTitle_(spreadsheetTitle);

  return {
    success: true,
    generatedAt: new Date().toISOString(),
    source: 'Google Spreadsheet',
    spreadsheetTitle: spreadsheetTitle,
    compiledDate: compiledDate,
    tamasya: tamasya,
    hpk: hpk,
    gati: gati,
    sidaya: sidaya,
    satyagatra: satyagatra,
    uppka: uppka,
    kecamatan: {
      pkl: pklKecamatan,
      sekolahLansia: slKecamatan
    }
  };
}


function extractDataDateFromTitle_(title) {
  if (!title) return null;
  const months = [
    'januari','februari','maret','april','mei','juni',
    'juli','agustus','september','oktober','november','desember'
  ];
  const t = String(title).replace(/\.xlsx/ig, ' ');
  let m = t.match(/(\d{1,2})[\s\-\/\.]*?(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember)[\s\-\/\.]*?(\d{4})/i);
  if (m) {
    const day = parseInt(m[1], 10);
    const monthName = String(m[2]).toLowerCase();
    const year = parseInt(m[3], 10);
    const monthIndex = months.indexOf(monthName);
    if (monthIndex < 0 || day < 1 || day > 31 || year < 2020 || year > 2100) return null;
    const daysInMonth = [31, (year % 4 === 0 ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if (day > daysInMonth[monthIndex]) return null;
    return day + ' ' + monthName.charAt(0).toUpperCase() + monthName.slice(1) + ' ' + year;
  }
  m = t.match(/\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\b/);
  if (m) {
    const day = parseInt(m[1], 10);
    const month = parseInt(m[2], 10);
    const year = parseInt(m[3], 10);
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    return day + ' ' + months[month - 1].charAt(0).toUpperCase() + months[month - 1].slice(1) + ' ' + year;
  }
  m = t.match(/\b(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\b/);
  if (m) {
    const year = parseInt(m[1], 10);
    const month = parseInt(m[2], 10);
    const day = parseInt(m[3], 10);
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    return day + ' ' + months[month - 1].charAt(0).toUpperCase() + months[month - 1].slice(1) + ' ' + year;
  }
  return null;
}

function readSimpleBlock_(ss, sheetName, startRow, count) {
  const rows = readRows_(ss, sheetName, startRow, count, [3, 4, 5], function(v) {
    return {
      kab: normalizeKab_(v[0]),
      target: num_(v[1]),
      capaian: num_(v[2])
    };
  });
  return rows;
}

function readRows_(ss, sheetName, startRow, count, cols, mapper) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + sheetName);

  const minCol = Math.min.apply(null, cols);
  const maxCol = Math.max.apply(null, cols);
  const values = sh
    .getRange(startRow, minCol, count, maxCol - minCol + 1)
    .getValues();

  return values
    .map(function(row) {
      const selected = cols.map(function(col) {
        return row[col - minCol];
      });
      return mapper(selected);
    })
    .filter(function(x) {
      return x.kab && x.kab !== 'PROVINSI BANTEN';
    });
}

function readAllRows_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + sheetName);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(function(h) {
    return String(h || '').trim();
  });

  return values
    .slice(1)
    .filter(function(row) {
      return row.some(function(v) {
        return v !== '' && v !== null;
      });
    })
    .map(function(row) {
      const obj = {};

      headers.forEach(function(h, i) {
        if (!h) return;
        obj[normalizeKey_(h)] = value_(row[i]);
      });

      if (obj.kota_kabupaten) {
        obj.kab = normalizeKab_(obj.kota_kabupaten);
      }

      if (obj.kecamatan) {
        obj.kecamatan = String(obj.kecamatan).trim();
      }

      return obj;
    });
}

function value_(v) {
  if (v instanceof Date) return v.toISOString();
  return v === '' ? null : v;
}

function num_(v) {
  if (v === null || v === '' || v === undefined) return 0;
  if (typeof v === 'number') return v;

  const n = Number(
    String(v)
      .replace(/,/g, '')
      .replace('%', '')
  );

  return isNaN(n) ? 0 : n;
}

function normalizeKab_(v) {
  let s = String(v || '').trim().toUpperCase();

  s = s
    .replace(/^KAB\.\s*/, '')
    .replace(/^KABUPATEN\s*/, '');

  if (s === 'TANGERANG SELATAN' || s === 'KOTA TANGERANG SELATAN') {
    return 'KOTA TANGSEL';
  }

  if (s.indexOf('KOTA ') === 0) return s;
  return s;
}

function normalizeKey_(v) {
  return String(v || '')
    .trim()
    .toLowerCase()
    .replace(/[\\/-]+/g, '_')
    .replace(/\s+/g, '_')
    .replace(/[^a-z0-9_]/g, '')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function testApi() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log(JSON.stringify(buildPayload_(ss), null, 2));
}
