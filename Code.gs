/*******************************************************
 * SISTEM ABSENSI QR SISWA
 * Google Apps Script + Google Spreadsheet
 *
 * FILE : Code.gs
 * VERSI : 1.0
 *******************************************************/

const SPREADSHEET_ID = "1cPkf86_pLVUbLrLQVMBwAdelCmLN-itOuXWANG9o3Bg";

/*
 * Jika Code.gs dibuat dari Spreadsheet:
 * SPREADSHEET_ID boleh dikosongkan:
 *
 * const SPREADSHEET_ID = "";
 *
 * Sistem akan menggunakan Spreadsheet yang terikat.
 */

const APP_NAME = "SkyAttend QR";
const TIMEZONE = "Asia/Jakarta";
const SESSION_SECONDS = 21600; // 6 jam

const SHEETS = {
  USERS: "USERS",
  SISWA: "SISWA",
  GURU: "GURU",
  KELAS: "KELAS",
  JADWAL: "JADWAL",
  ABSENSI: "ABSENSI",
  QR_SISWA: "QR_SISWA",
  NOTIFIKASI: "NOTIFIKASI",
  PENGATURAN: "PENGATURAN",
  LOG: "LOG_AKTIVITAS"
};

const HEADERS = {
  USERS: [
    "ID_USER",
    "USERNAME",
    "PASSWORD",
    "NAMA",
    "ROLE",
    "ID_GURU",
    "STATUS",
    "CREATED_AT"
  ],

  SISWA: [
    "ID_SISWA",
    "NIS",
    "NISN",
    "NAMA_SISWA",
    "JK",
    "TEMPAT_LAHIR",
    "TANGGAL_LAHIR",
    "ALAMAT",
    "NO_HP_ORANG_TUA",
    "NAMA_ORANG_TUA",
    "ID_KELAS",
    "NAMA_KELAS",
    "STATUS",
    "QR_TOKEN",
    "CREATED_AT",
    "UPDATED_AT"
  ],

  GURU: [
    "ID_GURU",
    "NIP",
    "NAMA_GURU",
    "NO_HP",
    "EMAIL",
    "JABATAN",
    "ID_KELAS_WALI",
    "STATUS",
    "CREATED_AT"
  ],

  KELAS: [
    "ID_KELAS",
    "NAMA_KELAS",
    "TINGKAT",
    "JURUSAN",
    "WALI_KELAS",
    "RUANG",
    "STATUS",
    "CREATED_AT"
  ],

  JADWAL: [
    "ID_JADWAL",
    "ID_KELAS",
    "NAMA_KELAS",
    "JAM_MASUK",
    "BATAS_TERLAMBAT",
    "JAM_PULANG",
    "HARI",
    "STATUS"
  ],

  ABSENSI: [
    "ID_ABSENSI",
    "TANGGAL",
    "WAKTU",
    "ID_SISWA",
    "NIS",
    "NAMA_SISWA",
    "ID_KELAS",
    "NAMA_KELAS",
    "STATUS",
    "KETERANGAN",
    "METODE",
    "ID_USER",
    "NAMA_USER",
    "UPDATED_AT"
  ],

  QR_SISWA: [
    "ID_QR",
    "ID_SISWA",
    "NAMA_SISWA",
    "QR_TOKEN",
    "QR_DATA",
    "STATUS",
    "CREATED_AT"
  ],

  NOTIFIKASI: [
    "ID_NOTIFIKASI",
    "TANGGAL",
    "ID_SISWA",
    "NAMA_SISWA",
    "NO_HP_ORANG_TUA",
    "JENIS",
    "PESAN",
    "CHANNEL",
    "STATUS",
    "RESPONSE",
    "CREATED_AT"
  ],

  PENGATURAN: [
    "KEY",
    "VALUE",
    "KETERANGAN"
  ],

  LOG: [
    "ID_LOG",
    "WAKTU",
    "ID_USER",
    "NAMA_USER",
    "ROLE",
    "AKSI",
    "KETERANGAN"
  ]
};


/* =====================================================
 * WEB APP
 * ===================================================== */

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* =====================================================
 * DATABASE
 * ===================================================== */

function getSpreadsheet_() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== "GANTI_DENGAN_ID_SPREADSHEET_ANDA") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      "Spreadsheet tidak ditemukan. Masukkan SPREADSHEET_ID pada Code.gs."
    );
  }

  return ss;
}


function setupDatabase() {
  const ss = getSpreadsheet_();

  Object.keys(SHEETS).forEach(function(key) {
    const name = SHEETS[key];
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    const headers = HEADERS[key];

    if (sheet.getLastRow() === 0) {
      sheet
        .getRange(1, 1, 1, headers.length)
        .setValues([headers]);

      sheet.setFrozenRows(1);

      sheet
        .getRange(1, 1, 1, headers.length)
        .setFontWeight("bold");
    }
  });

  // Migrasi database lama/XLSX ke struktur aplikasi.
  normalizeDatabase_();

  // Isi data demo hanya jika sheet benar-benar kosong.
  seedDatabase_();

  // Sangat penting: buat QR untuk siswa yang sudah ada di database.
  ensureQrRecords_();

  return {
    success: true,
    message: "Database berhasil disiapkan dan disinkronkan."
  };
}



/* =====================================================
 * DATABASE COMPATIBILITY / MIGRATION
 * Menyamakan database XLSX lama dengan struktur aplikasi.
 * Data lama dipertahankan berdasarkan nama kolom/alias.
 * ===================================================== */

function normalizeDatabase_() {
  const ss = getSpreadsheet_();

  Object.keys(SHEETS).forEach(function(key) {
    const sheetName = SHEETS[key];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const expected = HEADERS[key];
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();

    if (lastCol === expected.length && lastRow > 0) {
      const current = sheet.getRange(1, 1, 1, lastCol)
        .getDisplayValues()[0]
        .map(function(v) { return String(v).trim(); });

      if (JSON.stringify(current) === JSON.stringify(expected)) {
        return;
      }
    }

    const oldHeaders = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol)
          .getDisplayValues()[0]
          .map(function(v) { return String(v).trim(); })
      : [];

    const oldRows = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
      : [];

    const aliases = getHeaderAliases_(key);

    const normalizedRows = oldRows.map(function(row) {
      return expected.map(function(header) {
        const candidates = [header].concat(aliases[header] || []);
        let idx = -1;

        for (let i = 0; i < candidates.length; i++) {
          idx = oldHeaders.indexOf(candidates[i]);
          if (idx !== -1) break;
        }

        return idx !== -1 ? row[idx] : "";
      });
    });

    sheet.clearContents();
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);

    if (normalizedRows.length) {
      sheet.getRange(2, 1, normalizedRows.length, expected.length)
        .setValues(normalizedRows);
    }

    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, expected.length).setFontWeight("bold");
  });

  repairDerivedData_();
}

function getHeaderAliases_(key) {
  const common = {
    SISWA: {
      NIS: ["NOMOR_INDUK", "NO_INDUK"],
      NISN: ["NO_NISN"],
      ID_KELAS: ["KELAS_ID"],
      NAMA_KELAS: ["KELAS", "NAMA_KELAS_SISWA"],
      QR_TOKEN: ["QR_CODE", "TOKEN_QR"]
    },
    GURU: {
      ID_KELAS_WALI: ["KELAS_WALI", "ID_KELAS"],
      JABATAN: ["POSISI"]
    },
    KELAS: {
      WALI_KELAS: ["NAMA_WALI", "WALI"],
      ID_GURU_WALI: ["GURU_WALI"]
    },
    JADWAL: {
      BATAS_TERLAMBAT: ["BATAS_TERLAMBAT", "JAM_TERLAMBAT"],
      HARI: ["NAMA_HARI"],
      NAMA_KELAS: ["KELAS"]
    },
    ABSENSI: {
      ID_USER: ["DICATAT_OLEH_ID"],
      NAMA_USER: ["DICATAT_OLEH", "OPERATOR"],
      METODE: ["CARA"],
      KETERANGAN: ["CATATAN"],
      NIS: ["NO_INDUK"]
    },
    NOTIFIKASI: {
      ID_NOTIFIKASI: ["ID_NOTIF"],
      NO_HP_ORANG_TUA: ["TUJUAN", "NO_HP"],
      CHANNEL: ["SALURAN"]
    },
    PENGATURAN: {
      KEY: ["KUNCI"],
      VALUE: ["NILAI"]
    }
  };

  return common[key] || {};
}

function repairDerivedData_() {
  const ss = getSpreadsheet_();

  const kelas = getRecords_(SHEETS.KELAS);
  const guru = getRecords_(SHEETS.GURU);
  const siswaSheet = getSheet_(SHEETS.SISWA);
  const siswa = getRecords_(SHEETS.SISWA);

  // Lengkapi nama kelas siswa dan timestamp/QR.
  siswa.forEach(function(s) {
    let changed = false;

    if (!s.NAMA_KELAS && s.ID_KELAS) {
      const k = kelas.find(function(x) {
        return String(x.ID_KELAS) === String(s.ID_KELAS);
      });
      if (k) {
        s.NAMA_KELAS = k.NAMA_KELAS || "";
        changed = true;
      }
    }

    if (!s.QR_TOKEN) {
      s.QR_TOKEN = generateToken_();
      changed = true;
    }

    if (!s.CREATED_AT) {
      s.CREATED_AT = new Date();
      changed = true;
    }

    if (!s.UPDATED_AT) {
      s.UPDATED_AT = new Date();
      changed = true;
    }

    if (changed) {
      updateRecord_(SHEETS.SISWA, "ID_SISWA", s.ID_SISWA, s);
    }
  });

  // Lengkapi nama wali kelas jika database lama hanya menyimpan ID guru.
  const kelasRows = getRecords_(SHEETS.KELAS);
  kelasRows.forEach(function(k) {
    let changed = false;

    // Database XLSX lama dapat menyimpan ID guru di kolom wali.
    if (k.WALI_KELAS) {
      const wali = guru.find(function(g) {
        return String(g.ID_GURU) === String(k.WALI_KELAS);
      });

      if (wali) {
        k.WALI_KELAS = wali.NAMA_GURU || k.WALI_KELAS;
        changed = true;
      }
    }

    if (changed) {
      updateRecord_(SHEETS.KELAS, "ID_KELAS", k.ID_KELAS, k);
    }
  });

  // Lengkapi nama kelas pada jadwal.
  const jadwalRows = getRecords_(SHEETS.JADWAL);
  jadwalRows.forEach(function(j) {
    let changed = false;

    if (!j.NAMA_KELAS && j.ID_KELAS) {
      const k = kelas.find(function(x) {
        return String(x.ID_KELAS) === String(j.ID_KELAS);
      });
      if (k) {
        j.NAMA_KELAS = k.NAMA_KELAS || "";
        changed = true;
      }
    }

    // Database XLSX menggunakan TOLERANSI_MENIT.
    // Jika BATAS_TERLAMBAT kosong, hitung dari JAM_MASUK + 15 menit.
    if (!j.BATAS_TERLAMBAT && j.JAM_MASUK) {
      j.BATAS_TERLAMBAT = addMinutesToTime_(
        normalizeTime_(j.JAM_MASUK),
        15
      );
      changed = true;
    }

    if (changed) {
      updateRecord_(SHEETS.JADWAL, "ID_JADWAL", j.ID_JADWAL, j);
    }
  });

  // Lengkapi data absensi dari master siswa/kelas.
  const absensi = getRecords_(SHEETS.ABSENSI);
  absensi.forEach(function(a) {
    const s = siswa.find(function(x) {
      return String(x.ID_SISWA) === String(a.ID_SISWA);
    });

    if (!s) return;

    let changed = false;

    if (!a.NIS && s.NIS) {
      a.NIS = s.NIS;
      changed = true;
    }

    if (!a.NAMA_KELAS && s.NAMA_KELAS) {
      a.NAMA_KELAS = s.NAMA_KELAS;
      changed = true;
    }

    if (changed) {
      updateRecord_(SHEETS.ABSENSI, "ID_ABSENSI", a.ID_ABSENSI, a);
    }
  });

  ensureQrRecords_();
}

function normalizeTime_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, TIMEZONE, "HH:mm");
  }

  const s = String(value || "").trim();

  if (!s) return "";

  if (/^\d{1,2}:\d{2}$/.test(s)) {
    const p = s.split(":");
    return String(p[0]).padStart(2, "0") + ":" + p[1];
  }

  if (/^\d{1,2}:\d{2}:\d{2}$/.test(s)) {
    return s.substring(0, 5);
  }

  return s.substring(0, 5);
}

function addMinutesToTime_(time, minutes) {
  const parts = String(time || "00:00").split(":");
  let h = Number(parts[0] || 0);
  let m = Number(parts[1] || 0);

  m += Number(minutes || 0);
  h += Math.floor(m / 60);
  m = m % 60;
  h = h % 24;

  return String(h).padStart(2, "0") + ":" +
         String(m).padStart(2, "0");
}

function ensureQrRecords_() {
  const siswa = getRecords_(SHEETS.SISWA);
  const qrRows = getRecords_(SHEETS.QR_SISWA);

  siswa.forEach(function(s) {
    if (!s.ID_SISWA) return;

    const token = s.QR_TOKEN || generateToken_();
    const qrData = "SKYATTEND|" + s.ID_SISWA + "|" + token;

    if (!s.QR_TOKEN) {
      s.QR_TOKEN = token;
      s.UPDATED_AT = new Date();
      updateRecord_(SHEETS.SISWA, "ID_SISWA", s.ID_SISWA, s);
    }

    const existing = qrRows.find(function(q) {
      return String(q.ID_SISWA) === String(s.ID_SISWA);
    });

    if (!existing) {
      appendRecord_(SHEETS.QR_SISWA, {
        ID_QR: generateId_("QR"),
        ID_SISWA: s.ID_SISWA,
        NAMA_SISWA: s.NAMA_SISWA,
        QR_TOKEN: token,
        QR_DATA: qrData,
        STATUS: "AKTIF",
        CREATED_AT: new Date()
      });
    } else if (
      String(existing.QR_TOKEN) !== String(token) ||
      String(existing.QR_DATA) !== String(qrData) ||
      String(existing.NAMA_SISWA) !== String(s.NAMA_SISWA)
    ) {
      updateRecord_(SHEETS.QR_SISWA, "ID_SISWA", s.ID_SISWA, {
        ID_QR: existing.ID_QR,
        ID_SISWA: s.ID_SISWA,
        NAMA_SISWA: s.NAMA_SISWA,
        QR_TOKEN: token,
        QR_DATA: qrData,
        STATUS: String(s.STATUS).toUpperCase() === "AKTIF"
          ? "AKTIF"
          : "NONAKTIF",
        CREATED_AT: existing.CREATED_AT || new Date()
      });
    }
  });
}


/* =====================================================
 * HELPER SHEET
 * ===================================================== */

function getSheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);

  if (!sheet) {
    throw new Error("Sheet " + name + " belum tersedia. Jalankan setupDatabase().");
  }

  return sheet;
}


function getRecords_(sheetName) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const values = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();

  const headers = HEADERS[sheetName];

  return values.map(function(row) {
    const obj = {};

    headers.forEach(function(header, i) {
      obj[header] = formatValue_(row[i], header);
    });

    return obj;
  });
}


function formatValue_(value, fieldName) {
  if (value instanceof Date) {
    if (
      fieldName === "TANGGAL" ||
      fieldName === "TANGGAL_LAHIR"
    ) {
      return Utilities.formatDate(
        value,
        TIMEZONE,
        "yyyy-MM-dd"
      );
    }

    if (
      fieldName === "JAM_MASUK" ||
      fieldName === "JAM_PULANG" ||
      fieldName === "BATAS_TERLAMBAT"
    ) {
      return Utilities.formatDate(
        value,
        TIMEZONE,
        "HH:mm"
      );
    }

    if (fieldName === "WAKTU") {
      return Utilities.formatDate(
        value,
        TIMEZONE,
        "HH:mm:ss"
      );
    }

    return Utilities.formatDate(
      value,
      TIMEZONE,
      "yyyy-MM-dd HH:mm:ss"
    );
  }

  return value;
}


function appendRecord_(sheetName, data) {
  const sheet = getSheet_(sheetName);
  const headers = HEADERS[sheetName];

  const row = headers.map(function(header) {
    return data[header] !== undefined ? data[header] : "";
  });

  sheet.appendRow(row);

  return true;
}


function findRowById_(sheetName, field, value) {
  const sheet = getSheet_(sheetName);
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return -1;

  const headers = data[0];
  const index = headers.indexOf(field);

  if (index === -1) return -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][index]) === String(value)) {
      return i + 1;
    }
  }

  return -1;
}


function updateRecord_(sheetName, field, value, data) {
  const sheet = getSheet_(sheetName);
  const rowNumber = findRowById_(sheetName, field, value);

  if (rowNumber === -1) {
    throw new Error("Data tidak ditemukan.");
  }

  const headers = HEADERS[sheetName];

  const row = headers.map(function(header) {
    return data[header] !== undefined
      ? data[header]
      : "";
  });

  sheet
    .getRange(rowNumber, 1, 1, headers.length)
    .setValues([row]);

  return true;
}


function deleteRecord_(sheetName, field, value) {
  const sheet = getSheet_(sheetName);
  const rowNumber = findRowById_(sheetName, field, value);

  if (rowNumber === -1) {
    throw new Error("Data tidak ditemukan.");
  }

  sheet.deleteRow(rowNumber);

  return true;
}


function generateId_(prefix) {
  return prefix +
    "-" +
    Utilities.formatDate(
      new Date(),
      TIMEZONE,
      "yyyyMMddHHmmss"
    ) +
    "-" +
    Math.floor(100 + Math.random() * 900);
}


function generateToken_() {
  return Utilities.getUuid().replace(/-/g, "").toUpperCase();
}


/* =====================================================
 * SEED DATABASE
 * ===================================================== */

function seedDatabase_() {
  const ss = getSpreadsheet_();

  /* USERS */

  const usersSheet = ss.getSheetByName(SHEETS.USERS);

  if (usersSheet.getLastRow() < 2) {

    const users = [
      [
        "USR-001",
        "salma",
        "salma218",
        "Administrator",
        "ADMIN",
        "",
        "AKTIF",
        new Date()
      ],
      [
        "USR-002",
        "operator",
        "operator123",
        "Operator Sekolah",
        "OPERATOR",
        "",
        "AKTIF",
        new Date()
      ],
      [
        "USR-003",
        "guru",
        "guru123",
        "Guru / Wali Kelas",
        "GURU",
        "G001",
        "AKTIF",
        new Date()
      ]
    ];

    usersSheet
      .getRange(
        2,
        1,
        users.length,
        users[0].length
      )
      .setValues(users);
  }


  /* KELAS */

  const kelasSheet = ss.getSheetByName(SHEETS.KELAS);

  if (kelasSheet.getLastRow() < 2) {

    const kelas = [
      [
        "KLS-001",
        "X IPA 1",
        "X",
        "IPA",
        "Bapak Ahmad",
        "R.101",
        "AKTIF",
        new Date()
      ],
      [
        "KLS-002",
        "X IPA 2",
        "X",
        "IPA",
        "Ibu Siti",
        "R.102",
        "AKTIF",
        new Date()
      ],
      [
        "KLS-003",
        "X IPS 1",
        "X",
        "IPS",
        "Bapak Rudi",
        "R.103",
        "AKTIF",
        new Date()
      ]
    ];

    kelasSheet
      .getRange(
        2,
        1,
        kelas.length,
        kelas[0].length
      )
      .setValues(kelas);
  }


  /* GURU */

  const guruSheet = ss.getSheetByName(SHEETS.GURU);

  if (guruSheet.getLastRow() < 2) {

    const guru = [
      [
        "G001",
        "198001012005011001",
        "Bapak Ahmad",
        "081234567890",
        "ahmad@example.com",
        "Guru / Wali Kelas",
        "KLS-001",
        "AKTIF",
        new Date()
      ],
      [
        "G002",
        "198202022006022002",
        "Ibu Siti",
        "081234567891",
        "siti@example.com",
        "Guru / Wali Kelas",
        "KLS-002",
        "AKTIF",
        new Date()
      ]
    ];

    guruSheet
      .getRange(
        2,
        1,
        guru.length,
        guru[0].length
      )
      .setValues(guru);
  }


  /* JADWAL */

  const jadwalSheet = ss.getSheetByName(SHEETS.JADWAL);

  if (jadwalSheet.getLastRow() < 2) {

    const jadwal = [
      [
        "JDW-001",
        "KLS-001",
        "X IPA 1",
        "07:00",
        "07:15",
        "14:00",
        "SENIN-JUMAT",
        "AKTIF"
      ],
      [
        "JDW-002",
        "KLS-002",
        "X IPA 2",
        "07:00",
        "07:15",
        "14:00",
        "SENIN-JUMAT",
        "AKTIF"
      ],
      [
        "JDW-003",
        "KLS-003",
        "X IPS 1",
        "07:00",
        "07:15",
        "14:00",
        "SENIN-JUMAT",
        "AKTIF"
      ]
    ];

    jadwalSheet
      .getRange(
        2,
        1,
        jadwal.length,
        jadwal[0].length
      )
      .setValues(jadwal);
  }


  /* SISWA */

  const siswaSheet = ss.getSheetByName(SHEETS.SISWA);

  if (siswaSheet.getLastRow() < 2) {

    const sample = [
      ["S001","24001","006700001","Ahmad Fauzan","L","Payakumbuh","2009-01-10","Jl. Merdeka","081234560001","Bapak Fauzan","KLS-001","X IPA 1","AKTIF","","", ""],
      ["S002","24002","006700002","Budi Pratama","L","Payakumbuh","2009-02-15","Jl. Sudirman","081234560002","Bapak Budi","KLS-001","X IPA 1","AKTIF","","",""],
      ["S003","24003","006700003","Citra Amelia","P","Bukittinggi","2009-03-20","Jl. Diponegoro","081234560003","Ibu Amelia","KLS-001","X IPA 1","AKTIF","","",""],
      ["S004","24004","006700004","Dinda Putri","P","Payakumbuh","2009-04-18","Jl. Soekarno","081234560004","Ibu Dinda","KLS-002","X IPA 2","AKTIF","","",""],
      ["S005","24005","006700005","Fajar Ramadhan","L","Payakumbuh","2009-05-12","Jl. Imam Bonjol","081234560005","Bapak Fajar","KLS-002","X IPA 2","AKTIF","","",""],
      ["S006","24006","006700006","Gilang Saputra","L","Payakumbuh","2009-06-25","Jl. Ahmad Yani","081234560006","Bapak Gilang","KLS-003","X IPS 1","AKTIF","","",""],
      ["S007","24007","006700007","Hana Safitri","P","Payakumbuh","2009-07-08","Jl. Veteran","081234560007","Ibu Hana","KLS-003","X IPS 1","AKTIF","","",""],
      ["S008","24008","006700008","Indra Maulana","L","Payakumbuh","2009-08-14","Jl. Khatib Sulaiman","081234560008","Bapak Indra","KLS-001","X IPA 1","AKTIF","","",""],
      ["S009","24009","006700009","Jihan Aulia","P","Payakumbuh","2009-09-05","Jl. Tan Malaka","081234560009","Ibu Jihan","KLS-002","X IPA 2","AKTIF","","",""],
      ["S010","24010","006700010","Kevin Akbar","L","Payakumbuh","2009-10-21","Jl. Cendana","081234560010","Bapak Kevin","KLS-003","X IPS 1","AKTIF","","",""]
    ];

    sample.forEach(function(row) {

      const token = generateToken_();

      row[13] = token;
      row[14] = new Date();
      row[15] = new Date();

      siswaSheet.appendRow(row);

      appendRecord_(SHEETS.QR_SISWA, {
        ID_QR: generateId_("QR"),
        ID_SISWA: row[0],
        NAMA_SISWA: row[3],
        QR_TOKEN: token,
        QR_DATA: "SKYATTEND|" + row[0] + "|" + token,
        STATUS: "AKTIF",
        CREATED_AT: new Date()
      });
    });
  }


  /* PENGATURAN */

  const pengaturan = ss.getSheetByName(SHEETS.PENGATURAN);

  if (pengaturan.getLastRow() < 2) {

    pengaturan.getRange(2,1,6,3).setValues([
      ["NAMA_SEKOLAH","Sekolah Demo SkyAttend","Nama sekolah"],
      ["ALAMAT_SEKOLAH","Indonesia","Alamat sekolah"],
      ["JAM_MASUK","07:00","Jam masuk"],
      ["BATAS_TERLAMBAT","07:15","Batas keterlambatan"],
      ["JAM_PULANG","14:00","Jam pulang"],
      ["EMAIL_ADMIN","","Email administrator"]
    ]);
  }
}


/* =====================================================
 * LOGIN
 * ===================================================== */

function login(username, password) {

  // Validasi akun utama dilakukan sebelum akses spreadsheet agar login
  // tidak tertahan oleh proses setup database.
  username = String(username || "").trim();
  password = String(password || "");

  if (!username || !password) {
    return {
      success: false,
      message: "Username dan password wajib diisi."
    };
  }

  // Login utama aplikasi.
  // Tetap berfungsi walaupun Sheet USERS masih menyimpan
  // akun lama dari versi sebelumnya.
  if (
    username.toLowerCase() === "salma" &&
    password === "salma218"
  ) {

    const token = Utilities.getUuid();

    CacheService
      .getScriptCache()
      .put(
        "SESSION_" + token,
        JSON.stringify({
          ID_USER: "USR-SALMA",
          USERNAME: "salma",
          NAMA: "Salma",
          ROLE: "ADMIN",
          ID_GURU: ""
        }),
        SESSION_SECONDS
      );

    return {
      success: true,
      token: token,
      user: {
        ID_USER: "USR-SALMA",
        USERNAME: "salma",
        NAMA: "Salma",
        ROLE: "ADMIN",
        ID_GURU: ""
      }
    };
  }

  // Akun lain dari Sheet USERS tetap dipertahankan.
  const users = getRecords_(SHEETS.USERS);

  const user = users.find(function(u) {
    return (
      String(u.USERNAME).toLowerCase() === username.toLowerCase() &&
      String(u.PASSWORD) === password &&
      String(u.STATUS).toUpperCase() === "AKTIF"
    );
  });

  if (!user) {
    return {
      success: false,
      message: "Username atau password salah."
    };
  }

  const token = Utilities.getUuid();

  CacheService
    .getScriptCache()
    .put(
      "SESSION_" + token,
      JSON.stringify({
        ID_USER: user.ID_USER,
        USERNAME: user.USERNAME,
        NAMA: user.NAMA,
        ROLE: user.ROLE,
        ID_GURU: user.ID_GURU
      }),
      SESSION_SECONDS
    );

  logActivity_(
    user,
    "LOGIN",
    "Login berhasil"
  );

  return {
    success: true,
    token: token,
    user: {
      ID_USER: user.ID_USER,
      USERNAME: user.USERNAME,
      NAMA: user.NAMA,
      ROLE: user.ROLE,
      ID_GURU: user.ID_GURU
    }
  };
}


function getSession(token) {
  // Fungsi publik untuk validasi sesi dari Index.html.
  // Jangan biarkan localStorage menentukan sesi tanpa verifikasi server.
  return getSession_(token);
}


function getSession_(token) {

  if (!token) {
    throw new Error("Sesi tidak ditemukan.");
  }

  const raw = CacheService
    .getScriptCache()
    .get("SESSION_" + token);

  if (!raw) {
    throw new Error("Sesi telah berakhir. Silakan login kembali.");
  }

  return JSON.parse(raw);
}


function logout(token) {

  const session = getSession_(token);

  CacheService
    .getScriptCache()
    .remove("SESSION_" + token);

  logActivity_(
    session,
    "LOGOUT",
    "Logout"
  );

  return true;
}


/* =====================================================
 * DASHBOARD
 * ===================================================== */

function getDashboard(token) {

  const session = getSession_(token);

  const siswa = getRecords_(SHEETS.SISWA);
  const absensi = getRecords_(SHEETS.ABSENSI);
  const kelas = getRecords_(SHEETS.KELAS);

  const today = Utilities.formatDate(
    new Date(),
    TIMEZONE,
    "yyyy-MM-dd"
  );

  let dataHariIni = absensi.filter(function(a) {
    return String(a.TANGGAL) === today;
  });

  if (session.ROLE === "GURU" && session.ID_GURU) {

    const guru = getRecords_(SHEETS.GURU)
      .find(function(g) {
        return String(g.ID_GURU) === String(session.ID_GURU);
      });

    if (guru && guru.ID_KELAS_WALI) {
      dataHariIni = dataHariIni.filter(function(a) {
        return String(a.ID_KELAS) === String(guru.ID_KELAS_WALI);
      });
    }
  }

  const count = function(status) {
    return dataHariIni.filter(function(a) {
      return String(a.STATUS).toUpperCase() === status;
    }).length;
  };

  return {
    success: true,

    totalSiswa: siswa.filter(function(s) {
      return String(s.STATUS).toUpperCase() === "AKTIF";
    }).length,

    totalKelas: kelas.filter(function(k) {
      return String(k.STATUS).toUpperCase() === "AKTIF";
    }).length,

    hariIni: {
      hadir: count("HADIR"),
      sakit: count("SAKIT"),
      izin: count("IZIN"),
      alpha: count("ALPHA"),
      terlambat: count("TERLAMBAT"),
      total: dataHariIni.length
    },

    recent: dataHariIni
      .sort(function(a,b) {
        return String(b.WAKTU).localeCompare(String(a.WAKTU));
      })
      .slice(0, 20),

    timestamp: Utilities.formatDate(
      new Date(),
      TIMEZONE,
      "yyyy-MM-dd HH:mm:ss"
    )
  };
}


/* =====================================================
 * SISWA
 * ===================================================== */

function getSiswa(token) {

  getSession_(token);

  return getRecords_(SHEETS.SISWA);
}


function saveSiswa(token, data) {

  const session = getSession_(token);

  if (!["ADMIN","OPERATOR"].includes(session.ROLE)) {
    throw new Error("Anda tidak memiliki akses.");
  }

  const now = new Date();

  if (!data.ID_SISWA) {

    data.ID_SISWA = generateId_("SIS");
    data.QR_TOKEN = generateToken_();
    data.CREATED_AT = now;
    data.UPDATED_AT = now;

    appendRecord_(SHEETS.SISWA, data);

    appendRecord_(SHEETS.QR_SISWA, {
      ID_QR: generateId_("QR"),
      ID_SISWA: data.ID_SISWA,
      NAMA_SISWA: data.NAMA_SISWA,
      QR_TOKEN: data.QR_TOKEN,
      QR_DATA:
        "SKYATTEND|" +
        data.ID_SISWA +
        "|" +
        data.QR_TOKEN,
      STATUS: "AKTIF",
      CREATED_AT: now
    });

    logActivity_(
      session,
      "CREATE SISWA",
      data.NAMA_SISWA
    );

  } else {

    const old = getRecords_(SHEETS.SISWA)
      .find(function(s) {
        return String(s.ID_SISWA) === String(data.ID_SISWA);
      });

    data.QR_TOKEN =
      data.QR_TOKEN ||
      (old ? old.QR_TOKEN : generateToken_());

    data.CREATED_AT =
      data.CREATED_AT ||
      (old ? old.CREATED_AT : now);

    data.UPDATED_AT = now;

    updateRecord_(
      SHEETS.SISWA,
      "ID_SISWA",
      data.ID_SISWA,
      data
    );

    const qrRow = findRowById_(
      SHEETS.QR_SISWA,
      "ID_SISWA",
      data.ID_SISWA
    );

    if (qrRow === -1) {

      appendRecord_(SHEETS.QR_SISWA, {
        ID_QR: generateId_("QR"),
        ID_SISWA: data.ID_SISWA,
        NAMA_SISWA: data.NAMA_SISWA,
        QR_TOKEN: data.QR_TOKEN,
        QR_DATA:
          "SKYATTEND|" +
          data.ID_SISWA +
          "|" +
          data.QR_TOKEN,
        STATUS: "AKTIF",
        CREATED_AT: now
      });

    } else {

      const qrSheet = getSheet_(SHEETS.QR_SISWA);
      const headers = HEADERS.QR_SISWA;

      const qrValues = qrSheet
        .getRange(qrRow,1,1,headers.length)
        .getValues()[0];

      const qrObj = {};

      headers.forEach(function(h,i) {
        qrObj[h] = qrValues[i];
      });

      qrObj.NAMA_SISWA = data.NAMA_SISWA;
      qrObj.QR_TOKEN = data.QR_TOKEN;
      qrObj.QR_DATA =
        "SKYATTEND|" +
        data.ID_SISWA +
        "|" +
        data.QR_TOKEN;

      updateRecord_(
        SHEETS.QR_SISWA,
        "ID_SISWA",
        data.ID_SISWA,
        qrObj
      );
    }

    logActivity_(
      session,
      "UPDATE SISWA",
      data.NAMA_SISWA
    );
  }

  return {
    success: true,
    message: "Data siswa berhasil disimpan."
  };
}


function deleteSiswa(token, id) {

  const session = getSession_(token);

  if (!["ADMIN","OPERATOR"].includes(session.ROLE)) {
    throw new Error("Anda tidak memiliki akses.");
  }

  deleteRecord_(
    SHEETS.SISWA,
    "ID_SISWA",
    id
  );

  const qrRow = findRowById_(
    SHEETS.QR_SISWA,
    "ID_SISWA",
    id
  );

  if (qrRow !== -1) {
    getSheet_(SHEETS.QR_SISWA).deleteRow(qrRow);
  }

  logActivity_(
    session,
    "DELETE SISWA",
    id
  );

  return true;
}


/* =====================================================
 * QR
 * ===================================================== */

function getQrSiswa(token) {

  getSession_(token);

  return getRecords_(SHEETS.QR_SISWA);
}


function regenerateQr(token, idSiswa) {

  const session = getSession_(token);

  if (!["ADMIN","OPERATOR"].includes(session.ROLE)) {
    throw new Error("Anda tidak memiliki akses.");
  }

  const siswa = getRecords_(SHEETS.SISWA)
    .find(function(s) {
      return String(s.ID_SISWA) === String(idSiswa);
    });

  if (!siswa) {
    throw new Error("Siswa tidak ditemukan.");
  }

  const newToken = generateToken_();

  siswa.QR_TOKEN = newToken;
  siswa.UPDATED_AT = new Date();

  updateRecord_(
    SHEETS.SISWA,
    "ID_SISWA",
    idSiswa,
    siswa
  );

  const qr = getRecords_(SHEETS.QR_SISWA)
    .find(function(q) {
      return String(q.ID_SISWA) === String(idSiswa);
    });

  if (qr) {
    qr.QR_TOKEN = newToken;
    qr.QR_DATA =
      "SKYATTEND|" +
      idSiswa +
      "|" +
      newToken;

    updateRecord_(
      SHEETS.QR_SISWA,
      "ID_SISWA",
      idSiswa,
      qr
    );
  }

  return {
    success: true,
    qrData:
      "SKYATTEND|" +
      idSiswa +
      "|" +
      newToken
  };
}


/* =====================================================
 * ABSENSI QR
 * ===================================================== */

function scanQr(token, qrData) {

  const session = getSession_(token);

  qrData = String(qrData || "").trim();

  if (!qrData) {
    throw new Error("QR tidak terbaca.");
  }

  const parts = qrData.split("|");

  if (parts.length !== 3 || parts[0] !== "SKYATTEND") {
    throw new Error("QR bukan QR siswa SkyAttend.");
  }

  const idSiswa = parts[1];
  const qrToken = parts[2];

  const siswa = getRecords_(SHEETS.SISWA)
    .find(function(s) {
      return (
        String(s.ID_SISWA) === String(idSiswa) &&
        String(s.QR_TOKEN) === String(qrToken) &&
        String(s.STATUS).toUpperCase() === "AKTIF"
      );
    });

  if (!siswa) {
    throw new Error("QR siswa tidak valid atau sudah tidak aktif.");
  }

  return recordAttendance_(
    session,
    siswa,
    "HADIR",
    "Scan QR"
  );
}


function manualAttendance(token, data) {

  const session = getSession_(token);

  const siswa = getRecords_(SHEETS.SISWA)
    .find(function(s) {
      return String(s.ID_SISWA) === String(data.ID_SISWA);
    });

  if (!siswa) {
    throw new Error("Siswa tidak ditemukan.");
  }

  return recordAttendance_(
    session,
    siswa,
    String(data.STATUS || "HADIR").toUpperCase(),
    data.KETERANGAN || "Input manual"
  );
}


function recordAttendance_(session, siswa, status, keterangan) {

  const allowed = [
    "HADIR",
    "SAKIT",
    "IZIN",
    "ALPHA",
    "TERLAMBAT"
  ];

  if (!allowed.includes(status)) {
    throw new Error("Status absensi tidak valid.");
  }

  const now = new Date();

  const tanggal = Utilities.formatDate(
    now,
    TIMEZONE,
    "yyyy-MM-dd"
  );

  const waktu = Utilities.formatDate(
    now,
    TIMEZONE,
    "HH:mm:ss"
  );

  const absensi = getRecords_(SHEETS.ABSENSI);

  const existing = absensi.find(function(a) {
    return (
      String(a.TANGGAL) === tanggal &&
      String(a.ID_SISWA) === String(siswa.ID_SISWA)
    );
  });

  /*
   * Jika sudah ada HADIR/TERLAMBAT,
   * scan berikutnya tidak membuat duplikat.
   */

  if (existing) {

    if (
      status === "HADIR" &&
      existing.STATUS === "HADIR"
    ) {
      return {
        success: true,
        duplicate: true,
        message:
          siswa.NAMA_SISWA +
          " sudah melakukan absensi hari ini.",
        siswa: siswa,
        attendance: existing
      };
    }

    const updated = {
      ...existing,
      WAKTU: waktu,
      STATUS: status,
      KETERANGAN: keterangan,
      METODE: "QR",
      ID_USER: session.ID_USER,
      NAMA_USER: session.NAMA,
      UPDATED_AT: now
    };

    updateRecord_(
      SHEETS.ABSENSI,
      "ID_ABSENSI",
      existing.ID_ABSENSI,
      updated
    );

    sendAttendanceNotification_(siswa, status, tanggal, waktu);

    return {
      success: true,
      duplicate: false,
      message:
        "Status " +
        status +
        " untuk " +
        siswa.NAMA_SISWA +
        " berhasil diperbarui.",
      siswa: siswa,
      attendance: updated
    };
  }


  /* Tentukan terlambat */

  if (status === "HADIR") {

    const jadwal = getRecords_(SHEETS.JADWAL)
      .find(function(j) {
        return (
          String(j.ID_KELAS) ===
          String(siswa.ID_KELAS) &&
          String(j.STATUS).toUpperCase() === "AKTIF"
        );
      });

    if (jadwal && jadwal.BATAS_TERLAMBAT) {

      if (waktu > String(jadwal.BATAS_TERLAMBAT) + ":00") {
        status = "TERLAMBAT";
      }
    }
  }


  const record = {
    ID_ABSENSI: generateId_("ABS"),
    TANGGAL: tanggal,
    WAKTU: waktu,
    ID_SISWA: siswa.ID_SISWA,
    NIS: siswa.NIS,
    NAMA_SISWA: siswa.NAMA_SISWA,
    ID_KELAS: siswa.ID_KELAS,
    NAMA_KELAS: siswa.NAMA_KELAS,
    STATUS: status,
    KETERANGAN: keterangan,
    METODE: "QR",
    ID_USER: session.ID_USER,
    NAMA_USER: session.NAMA,
    UPDATED_AT: now
  };

  appendRecord_(
    SHEETS.ABSENSI,
    record
  );

  sendAttendanceNotification_(
    siswa,
    status,
    tanggal,
    waktu
  );

  logActivity_(
    session,
    "ABSENSI",
    siswa.NAMA_SISWA +
    " - " +
    status
  );

  return {
    success: true,
    duplicate: false,
    message:
      siswa.NAMA_SISWA +
      " berhasil absen " +
      status +
      ".",
    siswa: siswa,
    attendance: record
  };
}


/* =====================================================
 * DATA ABSENSI
 * ===================================================== */

function getAbsensi(token, date) {

  const session = getSession_(token);

  const all = getRecords_(SHEETS.ABSENSI);

  let data = all;

  if (date) {
    data = data.filter(function(a) {
      return String(a.TANGGAL) === String(date);
    });
  }

  if (session.ROLE === "GURU" && session.ID_GURU) {

    const guru = getRecords_(SHEETS.GURU)
      .find(function(g) {
        return String(g.ID_GURU) === String(session.ID_GURU);
      });

    if (guru && guru.ID_KELAS_WALI) {
      data = data.filter(function(a) {
        return String(a.ID_KELAS) === String(guru.ID_KELAS_WALI);
      });
    }
  }

  return data.reverse();
}


/* =====================================================
 * GURU
 * ===================================================== */

function getGuru(token) {
  getSession_(token);
  return getRecords_(SHEETS.GURU);
}


function saveGuru(token, data) {

  const session = getSession_(token);

  if (!["ADMIN","OPERATOR"].includes(session.ROLE)) {
    throw new Error("Tidak memiliki akses.");
  }

  const now = new Date();

  if (!data.ID_GURU) {

    data.ID_GURU = generateId_("GURU");
    data.CREATED_AT = now;

    appendRecord_(
      SHEETS.GURU,
      data
    );

  } else {

    data.CREATED_AT =
      data.CREATED_AT || now;

    updateRecord_(
      SHEETS.GURU,
      "ID_GURU",
      data.ID_GURU,
      data
    );
  }

  return true;
}


function deleteGuru(token, id) {

  const session = getSession_(token);

  if (!["ADMIN","OPERATOR"].includes(session.ROLE)) {
    throw new Error("Tidak memiliki akses.");
  }

  deleteRecord_(
    SHEETS.GURU,
    "ID_GURU",
    id
  );

  return true;
}


/* =====================================================
 * KELAS
 * ===================================================== */

function getKelas(token) {
  getSession_(token);
  return getRecords_(SHEETS.KELAS);
}


function saveKelas(token, data) {

  const session = getSession_(token);

  if (!["ADMIN","OPERATOR"].includes(session.ROLE)) {
    throw new Error("Tidak memiliki akses.");
  }

  const now = new Date();

  if (!data.ID_KELAS) {

    data.ID_KELAS = generateId_("KLS");
    data.CREATED_AT = now;

    appendRecord_(
      SHEETS.KELAS,
      data
    );

  } else {

    data.CREATED_AT =
      data.CREATED_AT || now;

    updateRecord_(
      SHEETS.KELAS,
      "ID_KELAS",
      data.ID_KELAS,
      data
    );
  }

  return true;
}


function deleteKelas(token, id) {

  const session = getSession_(token);

  if (!["ADMIN","OPERATOR"].includes(session.ROLE)) {
    throw new Error("Tidak memiliki akses.");
  }

  deleteRecord_(
    SHEETS.KELAS,
    "ID_KELAS",
    id
  );

  return true;
}


/* =====================================================
 * JADWAL
 * ===================================================== */

function getJadwal(token) {
  getSession_(token);
  return getRecords_(SHEETS.JADWAL);
}


function saveJadwal(token, data) {

  const session = getSession_(token);

  if (!["ADMIN","OPERATOR"].includes(session.ROLE)) {
    throw new Error("Tidak memiliki akses.");
  }

  if (!data.ID_JADWAL) {

    data.ID_JADWAL = generateId_("JDW");

    appendRecord_(
      SHEETS.JADWAL,
      data
    );

  } else {

    updateRecord_(
      SHEETS.JADWAL,
      "ID_JADWAL",
      data.ID_JADWAL,
      data
    );
  }

  return true;
}


function deleteJadwal(token, id) {

  const session = getSession_(token);

  if (!["ADMIN","OPERATOR"].includes(session.ROLE)) {
    throw new Error("Tidak memiliki akses.");
  }

  deleteRecord_(
    SHEETS.JADWAL,
    "ID_JADWAL",
    id
  );

  return true;
}


/* =====================================================
 * USERS
 * ===================================================== */

function getUsers(token) {

  const session = getSession_(token);

  if (session.ROLE !== "ADMIN") {
    throw new Error("Hanya Admin yang dapat melihat pengguna.");
  }

  return getRecords_(SHEETS.USERS).map(function(u) {

    return {
      ID_USER: u.ID_USER,
      USERNAME: u.USERNAME,
      NAMA: u.NAMA,
      ROLE: u.ROLE,
      ID_GURU: u.ID_GURU,
      STATUS: u.STATUS,
      CREATED_AT: u.CREATED_AT
    };
  });
}


function saveUser(token, data) {

  const session = getSession_(token);

  if (session.ROLE !== "ADMIN") {
    throw new Error("Hanya Admin.");
  }

  if (!data.ID_USER) {

    data.ID_USER = generateId_("USR");
    data.CREATED_AT = new Date();

    appendRecord_(
      SHEETS.USERS,
      data
    );

  } else {

    const old = getRecords_(SHEETS.USERS)
      .find(function(u) {
        return String(u.ID_USER) === String(data.ID_USER);
      });

    data.CREATED_AT =
      data.CREATED_AT ||
      (old ? old.CREATED_AT : new Date());

    updateRecord_(
      SHEETS.USERS,
      "ID_USER",
      data.ID_USER,
      data
    );
  }

  return true;
}


function deleteUser(token, id) {

  const session = getSession_(token);

  if (session.ROLE !== "ADMIN") {
    throw new Error("Hanya Admin.");
  }

  if (String(id) === String(session.ID_USER)) {
    throw new Error("Akun yang sedang digunakan tidak dapat dihapus.");
  }

  deleteRecord_(
    SHEETS.USERS,
    "ID_USER",
    id
  );

  return true;
}


/* =====================================================
 * LAPORAN
 * ===================================================== */

function getReport(token, startDate, endDate) {

  const session = getSession_(token);

  let data = getRecords_(SHEETS.ABSENSI);

  if (startDate) {
    data = data.filter(function(a) {
      return String(a.TANGGAL) >= String(startDate);
    });
  }

  if (endDate) {
    data = data.filter(function(a) {
      return String(a.TANGGAL) <= String(endDate);
    });
  }

  if (session.ROLE === "GURU" && session.ID_GURU) {

    const guru = getRecords_(SHEETS.GURU)
      .find(function(g) {
        return String(g.ID_GURU) === String(session.ID_GURU);
      });

    if (guru && guru.ID_KELAS_WALI) {
      data = data.filter(function(a) {
        return String(a.ID_KELAS) === String(guru.ID_KELAS_WALI);
      });
    }
  }

  const summary = {
    HADIR: 0,
    TERLAMBAT: 0,
    SAKIT: 0,
    IZIN: 0,
    ALPHA: 0
  };

  data.forEach(function(a) {

    const status =
      String(a.STATUS).toUpperCase();

    if (summary[status] !== undefined) {
      summary[status]++;
    }
  });

  return {
    data: data.reverse(),
    summary: summary
  };
}


/* =====================================================
 * EXPORT EXCEL
 *
 * Google Apps Script tidak membuat XLSX native secara
 * langsung. Fungsi ini membuat Google Spreadsheet
 * sementara yang dapat didownload sebagai XLSX.
 * ===================================================== */

function exportExcel(token, startDate, endDate) {

  const session = getSession_(token);

  const report =
    getReport(
      token,
      startDate,
      endDate
    );

  const name =
    "Laporan_Absensi_" +
    (startDate || "") +
    "_" +
    (endDate || "");

  const ss =
    SpreadsheetApp.create(name);

  const sheet =
    ss.getSheets()[0];

  const headers = HEADERS.ABSENSI;

  sheet
    .getRange(1,1,1,headers.length)
    .setValues([headers]);

  if (report.data.length) {

    const values =
      report.data.map(function(row) {
        return headers.map(function(h) {
          return row[h] || "";
        });
      });

    sheet
      .getRange(
        2,
        1,
        values.length,
        headers.length
      )
      .setValues(values);
  }

  sheet.setFrozenRows(1);

  sheet
    .getRange(1,1,1,headers.length)
    .setFontWeight("bold");

  SpreadsheetApp.flush();

  logActivity_(
    session,
    "EXPORT EXCEL",
    startDate + " s/d " + endDate
  );

  return {
    success: true,
    url:
      "https://docs.google.com/spreadsheets/d/" +
      ss.getId() +
      "/export?format=xlsx",
    spreadsheetUrl: ss.getUrl()
  };
}


/* =====================================================
 * EXPORT PDF
 * ===================================================== */

function exportPdf(token, startDate, endDate) {

  const session = getSession_(token);

  const report =
    getReport(
      token,
      startDate,
      endDate
    );

  const temp =
    SpreadsheetApp.create(
      "PDF_ABSENSI_TEMP_" +
      new Date().getTime()
    );

  const sheet =
    temp.getSheets()[0];

  sheet.setName("Laporan Absensi");

  sheet
    .getRange("A1:H1")
    .merge();

  sheet
    .getRange("A1")
    .setValue(
      "LAPORAN ABSENSI SISWA"
    )
    .setFontWeight("bold")
    .setFontSize(16);

  sheet
    .getRange("A2")
    .setValue(
      "Periode: " +
      (startDate || "-") +
      " s/d " +
      (endDate || "-")
    );

  sheet
    .getRange("A4:E4")
    .setValues([[
      "Tanggal",
      "Waktu",
      "Nama Siswa",
      "Kelas",
      "Status"
    ]])
    .setFontWeight("bold");

  const rows =
    report.data.map(function(a) {
      return [
        a.TANGGAL,
        a.WAKTU,
        a.NAMA_SISWA,
        a.NAMA_KELAS,
        a.STATUS
      ];
    });

  if (rows.length) {

    sheet
      .getRange(
        5,
        1,
        rows.length,
        5
      )
      .setValues(rows);
  }

  sheet.autoResizeColumns(1,5);

  SpreadsheetApp.flush();

  const exportUrl =
    "https://docs.google.com/spreadsheets/d/" +
    temp.getId() +
    "/export?format=pdf" +
    "&portrait=true" +
    "&fitw=true" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=true" +
    "&gridlines=false" +
    "&fzr=false";

  const response =
    UrlFetchApp.fetch(
      exportUrl,
      {
        headers: {
          Authorization:
            "Bearer " +
            ScriptApp.getOAuthToken()
        }
      }
    );

  const blob =
    response.getBlob()
      .setName(
        "Laporan_Absensi_" +
        (startDate || "") +
        "_" +
        (endDate || "") +
        ".pdf"
      );

  const file =
    DriveApp.createFile(blob);

  /*
   * Berikan akses link agar dapat dibuka
   * oleh user yang memiliki link.
   */

  try {
    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );
  } catch (e) {}

  logActivity_(
    session,
    "EXPORT PDF",
    startDate + " s/d " + endDate
  );

  return {
    success: true,
    url: file.getUrl(),
    fileId: file.getId()
  };
}


/* =====================================================
 * NOTIFIKASI
 * ===================================================== */

function sendAttendanceNotification_(
  siswa,
  status,
  tanggal,
  waktu
) {

  const nomor =
    String(
      siswa.NO_HP_ORANG_TUA || ""
    ).trim();

  const pesan =
    "Informasi Absensi Sekolah\n\n" +
    "Nama: " +
    siswa.NAMA_SISWA +
    "\n" +
    "Kelas: " +
    siswa.NAMA_KELAS +
    "\n" +
    "Tanggal: " +
    tanggal +
    "\n" +
    "Waktu: " +
    waktu +
    "\n" +
    "Status: " +
    status +
    "\n\n" +
    "SkyAttend QR";

  /*
   * EMAIL
   */

  try {

    const guru =
      getRecords_(SHEETS.GURU)
        .find(function(g) {
          return String(g.ID_KELAS_WALI) ===
            String(siswa.ID_KELAS);
        });

    if (
      guru &&
      guru.EMAIL &&
      String(guru.EMAIL).includes("@")
    ) {

      MailApp.sendEmail({
        to: guru.EMAIL,
        subject:
          "Absensi " +
          siswa.NAMA_SISWA +
          " - " +
          status,
        body: pesan
      });
    }

  } catch (emailError) {

    appendNotificationLog_(
      siswa,
      status,
      pesan,
      "EMAIL",
      "ERROR: " + emailError.message
    );

    return;
  }


  /*
   * WHATSAPP
   *
   * Tidak mengirim otomatis jika API belum
   * dikonfigurasi.
   */

  if (nomor) {

    appendNotificationLog_(
      siswa,
      status,
      pesan,
      "WHATSAPP",
      "MENUNGGU_API"
    );
  }
}


function appendNotificationLog_(
  siswa,
  status,
  pesan,
  channel,
  response
) {

  appendRecord_(
    SHEETS.NOTIFIKASI,
    {
      ID_NOTIFIKASI: generateId_("NTF"),
      TANGGAL: Utilities.formatDate(
        new Date(),
        TIMEZONE,
        "yyyy-MM-dd"
      ),
      ID_SISWA: siswa.ID_SISWA,
      NAMA_SISWA: siswa.NAMA_SISWA,
      NO_HP_ORANG_TUA:
        siswa.NO_HP_ORANG_TUA,
      JENIS: status,
      PESAN: pesan,
      CHANNEL: channel,
      STATUS: "TERCATAT",
      RESPONSE: response,
      CREATED_AT: new Date()
    }
  );
}


/* =====================================================
 * LOG
 * ===================================================== */

function logActivity_(
  user,
  action,
  description
) {

  try {

    appendRecord_(
      SHEETS.LOG,
      {
        ID_LOG: generateId_("LOG"),
        WAKTU: new Date(),
        ID_USER: user.ID_USER || "",
        NAMA_USER: user.NAMA || "",
        ROLE: user.ROLE || "",
        AKSI: action,
        KETERANGAN: description
      }
    );

  } catch (e) {}
}


/* =====================================================
 * PENGATURAN
 * ===================================================== */

function getSettings(token) {

  getSession_(token);

  const data =
    getRecords_(SHEETS.PENGATURAN);

  const result = {};

  data.forEach(function(row) {
    result[row.KEY] = row.VALUE;
  });

  return result;
}


/* =====================================================
 * DATA AWAL LENGKAP
 * ===================================================== */



/* ============================================================
 * COMPATIBILITY LAYER FOR THE MOBILE INDEX
 * Menyesuaikan nama fungsi Index.html dengan backend lama.
 * TIDAK menghapus atau mengganti QR_TOKEN yang sudah ada.
 * ============================================================ */

function checkLogin(username, password) {
  const result = login(username, password);

  if (!result || !result.success) {
    return {
      ok: false,
      message: (result && result.message) || "Username atau password salah."
    };
  }

  return {
    ok: true,
    success: true,
    token: result.token,
    username: result.user.USERNAME,
    role: result.user.ROLE,
    name: result.user.NAMA,
    user: result.user
  };
}

function getAppData(token) {
  const session = getSession_(token);

  const siswaRaw = getRecords_(SHEETS.SISWA);
  const guruRaw = getRecords_(SHEETS.GURU);
  const kelasRaw = getRecords_(SHEETS.KELAS);
  const absensiRaw = getRecords_(SHEETS.ABSENSI);

  const students = siswaRaw.map(function(s) {
    return {
      "ID Siswa": s.ID_SISWA,
      "NIS": s.NIS,
      "NISN": s.NISN,
      "Nama Siswa": s.NAMA_SISWA,
      "JK": s.JK,
      "Tempat Lahir": s.TEMPAT_LAHIR,
      "Tanggal Lahir": s.TANGGAL_LAHIR,
      "Alamat": s.ALAMAT,
      "No. WhatsApp": s.NO_HP_ORANG_TUA,
      "Nama Orang Tua": s.NAMA_ORANG_TUA,
      "ID Kelas": s.ID_KELAS,
      "Nama Kelas": s.NAMA_KELAS,
      "Status": s.STATUS,
      "QR Token": s.QR_TOKEN
    };
  });

  const teachers = guruRaw.map(function(g) {
    return {
      "ID Guru": g.ID_GURU,
      "NIP": g.NIP,
      "Nama Guru": g.NAMA_GURU,
      "No. WhatsApp": g.NO_HP,
      "Email": g.EMAIL,
      "Mata Pelajaran": g.JABATAN,
      "ID Kelas Wali": g.ID_KELAS_WALI,
      "Status": g.STATUS
    };
  });

  const classes = kelasRaw.map(function(k) {
    return {
      "ID Kelas": k.ID_KELAS,
      "Nama Kelas": k.NAMA_KELAS,
      "Tingkat": k.TINGKAT,
      "Jurusan": k.JURUSAN,
      "Wali Kelas": k.WALI_KELAS,
      "Ruang": k.RUANG,
      "Status": k.STATUS
    };
  });

  const attendance = absensiRaw.map(function(a) {
    return {
      "ID Absensi": a.ID_ABSENSI,
      "Tanggal": a.TANGGAL,
      "Jam Scan": a.WAKTU,
      "ID Siswa": a.ID_SISWA,
      "NIS": a.NIS,
      "Nama Siswa": a.NAMA_SISWA,
      "ID Kelas": a.ID_KELAS,
      "Nama Kelas": a.NAMA_KELAS,
      "Status": a.STATUS,
      "Keterangan": a.KETERANGAN,
      "Metode": a.METODE,
      "ID User": a.ID_USER,
      "Nama User": a.NAMA_USER,
      "Updated At": a.UPDATED_AT
    };
  }).reverse();

  const today = Utilities.formatDate(
    new Date(),
    TIMEZONE,
    "yyyy-MM-dd"
  );

  const todayRows = absensiRaw.filter(function(a) {
    return String(a.TANGGAL) === today;
  });

  const counts = {
    HADIR: 0,
    IZIN: 0,
    SAKIT: 0,
    ALFA: 0,
    ALPHA: 0,
    TERLAMBAT: 0
  };

  todayRows.forEach(function(a) {
    const status =
      String(a.STATUS || "").toUpperCase();

    if (counts[status] !== undefined) {
      counts[status]++;
    }

    if (status === "ALPHA") {
      counts.ALFA++;
    }
  });

  return {
    students: students,
    teachers: teachers,
    classes: classes,
    attendance: attendance,
    dashboard: {
      date: today,
      totalStudents: students.filter(function(s) {
        return String(s.Status).toUpperCase() === "AKTIF";
      }).length,
      totalToday: todayRows.length,
      counts: counts
    }
  };
}


function processScan(token, decodedText, status, note) {
  const session = getSession_(token);

  decodedText =
    String(decodedText || "").trim();

  if (!decodedText) {
    return {
      ok: false,
      message: "QR kosong atau tidak terbaca."
    };
  }

  const siswa = getRecords_(SHEETS.SISWA)
    .find(function(s) {

      const tokenValue =
        String(s.QR_TOKEN || "").trim();

      const qrData =
        "SKYATTEND|" +
        String(s.ID_SISWA) +
        "|" +
        tokenValue;

      return (
        decodedText === tokenValue ||
        decodedText === qrData ||
        decodedText === String(s.ID_SISWA) ||
        decodedText.endsWith("|" + tokenValue)
      );
    });

  if (!siswa) {
    return {
      ok: false,
      message: "QR siswa tidak ditemukan di database."
    };
  }

  if (
    String(siswa.STATUS).toUpperCase() !==
    "AKTIF"
  ) {
    return {
      ok: false,
      message: "Siswa tidak aktif."
    };
  }

  const normalizedStatus =
    String(status || "HADIR")
      .toUpperCase()
      .replace("ALFA", "ALPHA");

  const result =
    recordAttendance_(
      session,
      siswa,
      normalizedStatus,
      note || "Scan QR"
    );

  return {
    ok: true,
    message: result.message,
    student: {
      id: siswa.ID_SISWA,
      name: siswa.NAMA_SISWA,
      nis: siswa.NIS,
      className: siswa.NAMA_KELAS
    },
    attendance: result.attendance
  };
}


function saveManualAttendance(
  token,
  studentId,
  status,
  note
) {
  const session = getSession_(token);

  const siswa =
    getRecords_(SHEETS.SISWA)
      .find(function(s) {
        return String(s.ID_SISWA) ===
          String(studentId);
      });

  if (!siswa) {
    throw new Error("Siswa tidak ditemukan.");
  }

  const normalized =
    String(status || "HADIR")
      .toUpperCase()
      .replace("ALFA", "ALPHA");

  const result =
    recordAttendance_(
      session,
      siswa,
      normalized,
      note || "Input manual"
    );

  return {
    ok: true,
    message: result.message
  };
}


function deleteAttendance(token, id) {
  const session = getSession_(token);

  if (!["ADMIN", "OPERATOR"].includes(session.ROLE)) {
    throw new Error("Anda tidak memiliki akses.");
  }

  deleteRecord_(
    SHEETS.ABSENSI,
    "ID_ABSENSI",
    id
  );

  return {
    ok: true,
    message: "Data absensi berhasil dihapus."
  };
}


function createStudent(token, data) {
  const session = getSession_(token);

  if (!["ADMIN", "OPERATOR"].includes(session.ROLE)) {
    throw new Error("Anda tidak memiliki akses.");
  }

  const now = new Date();
  const id = generateId_("SIS");
  const qrToken = generateToken_();

  appendRecord_(SHEETS.SISWA, {
    ID_SISWA: id,
    NIS: data.nis || "",
    NISN: "",
    NAMA_SISWA: data.name || "",
    JK: data.jk || "",
    TEMPAT_LAHIR: "",
    TANGGAL_LAHIR: "",
    ALAMAT: "",
    NO_HP_ORANG_TUA: data.whatsapp || "",
    NAMA_ORANG_TUA: data.parent || "",
    ID_KELAS: data.classId || "",
    NAMA_KELAS: data.className || "",
    STATUS: "AKTIF",
    QR_TOKEN: qrToken,
    CREATED_AT: now,
    UPDATED_AT: now
  });

  appendRecord_(SHEETS.QR_SISWA, {
    ID_QR: generateId_("QR"),
    ID_SISWA: id,
    NAMA_SISWA: data.name || "",
    QR_TOKEN: qrToken,
    QR_DATA: "SKYATTEND|" + id + "|" + qrToken,
    STATUS: "AKTIF",
    CREATED_AT: now
  });

  return {
    ok: true,
    message: "Siswa berhasil ditambahkan."
  };
}


function updateStudent(token, data) {
  const session = getSession_(token);

  if (!["ADMIN", "OPERATOR"].includes(session.ROLE)) {
    throw new Error("Anda tidak memiliki akses.");
  }

  const siswa =
    getRecords_(SHEETS.SISWA)
      .find(function(s) {
        return String(s.ID_SISWA) ===
          String(data.id);
      });

  if (!siswa) {
    throw new Error("Siswa tidak ditemukan.");
  }

  /*
   * QR_TOKEN SENGAJA DIAMBIL DARI DATA LAMA.
   * Edit siswa tidak boleh mengubah QR.
   */
  siswa.NIS = data.nis || siswa.NIS;
  siswa.NAMA_SISWA = data.name || siswa.NAMA_SISWA;
  siswa.JK = data.jk || siswa.JK;
  siswa.ID_KELAS = data.classId || siswa.ID_KELAS;
  siswa.NAMA_KELAS = data.className || siswa.NAMA_KELAS;
  siswa.NAMA_ORANG_TUA =
    data.parent || siswa.NAMA_ORANG_TUA;
  siswa.NO_HP_ORANG_TUA =
    data.whatsapp || siswa.NO_HP_ORANG_TUA;
  siswa.UPDATED_AT = new Date();

  updateRecord_(
    SHEETS.SISWA,
    "ID_SISWA",
    siswa.ID_SISWA,
    siswa
  );

  return {
    ok: true,
    message: "Data siswa berhasil diperbarui."
  };
}


function deleteStudent(token, id) {
  const session = getSession_(token);

  if (!["ADMIN", "OPERATOR"].includes(session.ROLE)) {
    throw new Error("Anda tidak memiliki akses.");
  }

  deleteRecord_(
    SHEETS.SISWA,
    "ID_SISWA",
    id
  );

  const qrRow =
    findRowById_(
      SHEETS.QR_SISWA,
      "ID_SISWA",
      id
    );

  if (qrRow !== -1) {
    getSheet_(SHEETS.QR_SISWA)
      .deleteRow(qrRow);
  }

  return {
    ok: true,
    message: "Siswa berhasil dihapus."
  };
}


function createTeacher(token, data) {
  const session = getSession_(token);

  if (!["ADMIN", "OPERATOR"].includes(session.ROLE)) {
    throw new Error("Anda tidak memiliki akses.");
  }

  const id = generateId_("GURU");

  appendRecord_(SHEETS.GURU, {
    ID_GURU: id,
    NIP: data.nip || "",
    NAMA_GURU: data.name || "",
    NO_HP: data.whatsapp || "",
    EMAIL: data.email || "",
    JABATAN: data.subject || "",
    ID_KELAS_WALI: "",
    STATUS: "AKTIF",
    CREATED_AT: new Date()
  });

  return {
    ok: true,
    message: "Data guru berhasil ditambahkan."
  };
}


function updateTeacher(token, data) {
  const session = getSession_(token);

  if (!["ADMIN", "OPERATOR"].includes(session.ROLE)) {
    throw new Error("Anda tidak memiliki akses.");
  }

  const guru =
    getRecords_(SHEETS.GURU)
      .find(function(g) {
        return String(g.ID_GURU) ===
          String(data.id);
      });

  if (!guru) {
    throw new Error("Guru tidak ditemukan.");
  }

  guru.NIP = data.nip || guru.NIP;
  guru.NAMA_GURU =
    data.name || guru.NAMA_GURU;
  guru.JABATAN =
    data.subject || guru.JABATAN;
  guru.EMAIL =
    data.email || guru.EMAIL;
  guru.NO_HP =
    data.whatsapp || guru.NO_HP;

  updateRecord_(
    SHEETS.GURU,
    "ID_GURU",
    guru.ID_GURU,
    guru
  );

  return {
    ok: true,
    message: "Data guru berhasil diperbarui."
  };
}


function deleteTeacher(token, id) {
  const session = getSession_(token);

  if (!["ADMIN", "OPERATOR"].includes(session.ROLE)) {
    throw new Error("Anda tidak memiliki akses.");
  }

  deleteRecord_(
    SHEETS.GURU,
    "ID_GURU",
    id
  );

  return {
    ok: true,
    message: "Data guru berhasil dihapus."
  };
}



function initializeApp() {

  const result =
    setupDatabase();

  return {
    success: true,
    message:
      result.message +
      " Aplikasi siap digunakan."
  };
}
