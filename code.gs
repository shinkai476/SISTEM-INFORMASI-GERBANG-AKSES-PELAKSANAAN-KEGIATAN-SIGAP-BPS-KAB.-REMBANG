/**
 * @OnlyCurrentDoc
 */



// ID Spreadsheet Database.
const SPREADSHEET_ID = "id spreadsheet anda ";

// Nama-nama Sheet Utama
const BUDGET_SHEET_NAME = "Data Anggaran";
const RISK_SHEET_NAME   = "Data Risiko";
const KINERJA_SHEET_NAME= "Data Kinerja";
const LEADER_SHEET_NAME = "Ref_KetuaTim"; 

// Durasi Cache (dalam detik). 3600 = 1 Jam.
const CACHE_EXPIRATION_SECONDS = 3600;

// Variabel global instance spreadsheet
let spreadsheet;

// ======================================================
// ================= HELPER FUNCTIONS ===================
// ======================================================

/**
 * Membuka Spreadsheet berdasarkan ID (Singleton).
 */
function getSpreadsheet() {
  if (spreadsheet) return spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    return spreadsheet;
  } catch (e) {
    console.error("KRITIS: Gagal membuka Spreadsheet: " + e.message);
    throw new Error("Gagal terhubung ke Database.");
  }
}

/**
 * Helper: Membersihkan format angka Indonesia
 */
const cleanNumber = v => {
  try {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'string' && v.trim().startsWith('#')) return 0; 
    
    // Hapus 'Rp', titik ribuan, spasi. Ganti koma desimal jadi titik.
    const cleanStr = String(v).replace(/Rp|\.| /g, '').replace(',', '.');
    return parseFloat(cleanStr) || 0;
  } catch (e) {
    return 0;
  }
};

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function clearCache(key) {
  const cache = CacheService.getScriptCache();
  if (key === BUDGET_SHEET_NAME) {
    cache.remove('budgetData_v2');
  } else if (key === LEADER_SHEET_NAME || key === 'leader_list') {
    cache.remove('leader_list');
  } else {
    cache.remove(key + '_data');
  }
}

function getAndPrepareSheet(sheetName, requiredHeaders) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet && sheetName !== BUDGET_SHEET_NAME) {
    if (!requiredHeaders) {
       requiredHeaders = ['ID', 'Judul Kegiatan', 'Ketua Tim', 'Deskripsi', 'Link Folder'];
    }
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(requiredHeaders);
    sheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight("bold").setBackground("#f1f5f9");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ======================================================
// ================= MAIN APP ROUTING ===================
// ======================================================

function doGet(e) {
  try { getSpreadsheet(); } catch(err) {} 
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('BPS Kabupaten Rembang - Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Router Backend: Perbaikan Utama ada di sini
 */
function getPageContent(pageName, requestToken) {
  try {
    let templateName;
    let pageData = null;
    
    // Routing Logic
    if (pageName === 'Dashboard' || pageName === 'dashboard') {
      templateName = 'dashboard';
      pageData = {
        budget: getBudgetData(),
        risk: getDataFromSheet(RISK_SHEET_NAME, ['ID', 'Risiko Aktual', 'Level', 'Keterangan', 'Link']),
        kinerja: getDataFromSheet(KINERJA_SHEET_NAME, ['ID', 'Indikator', 'Tw1', 'Tw2', 'Tw3', 'Tw4', 'TARGET', 'LINK'])
      };
      
    } else if (pageName && pageName.toLowerCase().startsWith('tim')) {
      templateName = 'tim';
      pageData = getDataFromSheet(pageName, ['ID', 'Judul Kegiatan', 'Ketua Tim', 'Deskripsi', 'Link Folder']);
      
    } else {
      templateName = 'home';
      pageData = getHomePageStats();
    }

    // --- [PERBAIKAN PENTING DI SINI] ---
    // Membuat template object terlebih dahulu
    const template = HtmlService.createTemplateFromFile(templateName);
    
    // Menyuntikkan variabel 'data' ke dalam template HTML
    // Ini agar kode <?!= JSON.stringify(data) ?> di HTML bisa berjalan
    template.data = pageData; 
    
    // Baru kemudian dievaluasi menjadi string HTML
    const html = template.evaluate().getContent();
    // -----------------------------------

    return { success: true, html: html, data: pageData, token: requestToken };
  } catch (e) {
    console.error(`Error in getPageContent for ${pageName}: ${e.stack}`);
    return { success: false, error: `Gagal memuat konten halaman: ${e.message}`, token: requestToken };
  }
}

// ======================================================
// ================ FUNGSI DATA ANGGARAN ================
// ======================================================

function getBudgetData() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'budgetData_v2'; 
  const cached = cache.get(cacheKey);

  if (cached != null) return JSON.parse(cached);

  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(BUDGET_SHEET_NAME);
    if (!sheet) return { success: false, error: `Sheet "${BUDGET_SHEET_NAME}" tidak ditemukan.` };

    const lastRow = sheet.getLastRow();
    if (lastRow < 5) return { success: true, allData: [] };

    const values = sheet.getRange(1, 1, lastRow, 20).getValues();
    
    let headerRowIndex = -1;
    let colIndices = { uraian: -1, pagu: -1, realisasi: -1, sisa: -1, persen: -1 };

    // 1. Logika Pencarian Header
    for (let i = 0; i < 20; i++) {
      const rowStr = values[i].map(String).join(" ").toLowerCase();
      if (rowStr.includes("pagu revisi") || (rowStr.includes("pagu") && rowStr.includes("revisi"))) {
        headerRowIndex = i;
        values[i].forEach((cell, idx) => {
          const c = String(cell).toLowerCase().trim();
          if (c.includes("uraian") || c.includes("program/kegiatan") || c.includes("kd.")) colIndices.uraian = idx;
          else if (c.includes("pagu revisi")) colIndices.pagu = idx;
          else if (c.includes("sisa anggaran") || c.includes("sisa dana")) colIndices.sisa = idx;
        });

        const checkRealisasi = (rIdx) => {
            if(rIdx >= values.length) return;
            values[rIdx].forEach((cell, idx) => {
                const c = String(cell).toLowerCase().trim();
                if (c.includes("s.d.") || c.includes("periode") || c.includes("realisasi")) colIndices.realisasi = idx;
                if (c === "%" || c.includes("persen")) colIndices.persen = idx;
            });
        };
        checkRealisasi(i);
        checkRealisasi(i + 1);
        
        if (colIndices.realisasi === -1 && colIndices.sisa !== -1) {
           colIndices.realisasi = colIndices.sisa - 2; 
        }
        break;
      }
    }

    // Fallback Header
    if (headerRowIndex === -1 || colIndices.pagu === -1) {
        colIndices = { uraian: 1, pagu: 4, realisasi: 5, sisa: 7 };
        headerRowIndex = 4; 
    }

    const allData = [];
    
    // 2. Loop Data
    for (let i = headerRowIndex + 1; i < values.length; i++) {
      const row = values[i];
      if (!row[colIndices.uraian] || String(row[colIndices.uraian]).includes("Hal 1 dari")) continue;

      const pagu = cleanNumber(row[colIndices.pagu]);
      const realisasi = cleanNumber(row[colIndices.realisasi]);
      const sisa = cleanNumber(row[colIndices.sisa]);
      let uraianLengkap = String(row[colIndices.uraian]).trim();
      
      if (colIndices.uraian > 0) {
          const kode = String(row[colIndices.uraian - 1]).trim();
          if (kode && kode.length < 15 && (kode.includes('.') || !isNaN(parseInt(kode)))) { 
              uraianLengkap = `[${kode}] ${uraianLengkap}`;
          }
      }

      if (pagu > 0 || realisasi > 0) {
        allData.push({
          rowIndex: i + 1,
          uraian: uraianLengkap,
          paguRevisi: pagu,
          sdPeriode: realisasi,
          sisaAnggaran: sisa,
          persenRealisasi: pagu > 0 ? (realisasi / pagu) : 0
        });
      }
    }

    const result = { success: true, allData };
    cache.put(cacheKey, JSON.stringify(result), CACHE_EXPIRATION_SECONDS);
    return result;

  } catch (e) {
    console.error(`Error in getBudgetData: ${e.stack}`);
    return { success: false, error: `Gagal membaca data anggaran: ${e.message}` };
  }
}

// ======================================================
// ================ FUNGSI DATA UMUM (CRUD) =============
// ======================================================

function getDataFromSheet(sheetName, requiredHeaders) {
  const cache = CacheService.getScriptCache();
  const cacheKey = sheetName + '_data';
  const cached = cache.get(cacheKey);

  if (cached != null) return JSON.parse(cached);
  
  try {
    const sheet = getAndPrepareSheet(sheetName, requiredHeaders);
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) return { success: true, data: [] };
    
    const numCols = requiredHeaders.length;
    const values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    
    const dataWithRowIndex = values.map((row, index) => {
      if (row && row[1] && String(row[1]).trim() !== '') {
        return [...row, index + 2]; 
      }
      return null;
    }).filter(Boolean);
    
    const result = { success: true, data: dataWithRowIndex };
    cache.put(cacheKey, JSON.stringify(result), CACHE_EXPIRATION_SECONDS);
    return result;
  } catch (e) {
    console.error(`Error in getDataFromSheet: ${e.stack}`);
    return { success: false, error: `Gagal mengambil data: ${e.message}` };
  }
}

function saveItem(itemData) {
  try {
    if (!itemData.judul || itemData.judul.trim() === '') throw new Error("Input utama wajib diisi.");

    let values;
    let requiredHeaders;
    const sheetNameLower = itemData.sheetName.toLowerCase();

    if (sheetNameLower.startsWith('tim')) {
      values = [itemData.judul, itemData.ketua, itemData.deskripsi, itemData.link];
      requiredHeaders = ['ID', 'Judul Kegiatan', 'Ketua Tim', 'Deskripsi', 'Link Folder'];
    } else if (itemData.sheetName === RISK_SHEET_NAME) {
      values = [itemData.judul, itemData.level, itemData.deskripsi, itemData.link];
      requiredHeaders = ['ID', 'Risiko Aktual', 'Level', 'Keterangan', 'Link'];
    } else if (itemData.sheetName === KINERJA_SHEET_NAME) {
      values = [itemData.judul, itemData.tw1, itemData.tw2, itemData.tw3, itemData.tw4, itemData.target, itemData.link];
      requiredHeaders = ['ID', 'Indikator', 'Tw1', 'Tw2', 'Tw3', 'Tw4', 'TARGET', 'LINK'];
    } else {
      throw new Error("Tipe sheet tidak dikenal.");
    }

    const sheet = getAndPrepareSheet(itemData.sheetName, requiredHeaders);
    
    if (itemData.rowIndex && itemData.rowIndex > 1) {
      sheet.getRange(itemData.rowIndex, 2, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow([Utilities.getUuid(), ...values]);
    }
    
    clearCache(itemData.sheetName);
    return { success: true, message: 'Data berhasil disimpan!' };
  } catch (e) {
    return { success: false, message: `Gagal menyimpan: ${e.message}` };
  }
}

function deleteItem(deleteData) {
  try {
    if (!deleteData.id || !deleteData.sheetName) throw new Error("Informasi hapus tidak lengkap.");

    const sheet = getAndPrepareSheet(deleteData.sheetName, null);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Item tidak ditemukan.' };
    
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const index = ids.indexOf(deleteData.id);
    
    if (index === -1) return { success: false, message: 'Item tidak ditemukan.' };

    sheet.deleteRow(index + 2);
    clearCache(deleteData.sheetName);
    return { success: true, message: 'Item berhasil dihapus.' };
  } catch (e) {
    return { success: false, message: `Gagal menghapus: ${e.message}` };
  }
}

// ======================================================
// ================== MANAJEMEN TIM =====================
// ======================================================

function createNewTeamSheet(teamName) {
  try {
    if (!teamName || teamName.trim() === '') throw new Error("Nama tim tidak boleh kosong.");
    const cleanName = "Tim " + teamName.trim(); 
    const ss = getSpreadsheet();
    if (ss.getSheetByName(cleanName)) return { success: false, message: `Tim sudah ada.` };

    const sheet = ss.insertSheet(cleanName);
    const headers = ['ID', 'Judul Kegiatan', 'Ketua Tim', 'Deskripsi', 'Link Folder'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f1f5f9");
    sheet.setFrozenRows(1);
    
    clearCache('stats_home');
    return { success: true, message: `Tim berhasil dibuat.`, sheetName: cleanName };
  } catch (e) {
    return { success: false, message: `Gagal membuat tim: ${e.message}` };
  }
}

function deleteTeamSheet(teamName) {
  try {
    if (!teamName) throw new Error("Nama tim kosong.");
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(teamName.trim());
    if (!sheet) return { success: false, message: `Tim tidak ditemukan.` };

    ss.deleteSheet(sheet);
    clearCache('stats_home');
    return { success: true, message: `Tim berhasil dihapus.` };
  } catch (e) {
    return { success: false, message: `Gagal menghapus tim: ${e.message}` };
  }
}

function getTeamList() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    const teamNames = sheets
      .map(sheet => sheet.getName())
      .filter(name => name.toLowerCase().startsWith("tim "))
      .sort();
    return { success: true, data: teamNames };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ======================================================
// ============ MANAJEMEN REFERENSI (KETUA) =============
// ======================================================

function getLeaderList() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('leader_list');
  if (cached != null) return JSON.parse(cached);

  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(LEADER_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(LEADER_SHEET_NAME);
    sheet.appendRow(['Nama Ketua']);
    const defaultNames = [
      ["Muncar Cahyono"], ["Herhardana"], ["Faisal Luthfi Arief"],
      ["Hadiyanto"], ["Mustaqhwiroh"], ["Winarso"],
      ["Wahyu Sri Lestari"], ["Miyan Andi Irawan"],
      ["M. Achiruzaman"], ["Khaerul Anwar"], ["Imam Mustofa"]
    ];
    sheet.getRange(2, 1, defaultNames.length, 1).setValues(defaultNames);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const sortedData = data.filter(n => n).sort();

  cache.put('leader_list', JSON.stringify(sortedData), 3600);
  return sortedData;
}

function addLeaderName(name) {
  if (!name) throw new Error("Nama tidak boleh kosong");
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(LEADER_SHEET_NAME);
  if (!sheet) { getLeaderList(); sheet = ss.getSheetByName(LEADER_SHEET_NAME); }
  
  sheet.appendRow([name]);
  clearCache('leader_list');
  return { success: true, message: "Nama berhasil ditambahkan." };
}

function deleteLeaderName(name) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(LEADER_SHEET_NAME);
  if (!sheet) return { success: false, message: "Database tidak ditemukan." };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const index = data.indexOf(name);
  if (index > -1) {
    sheet.deleteRow(index + 2); 
    clearCache('leader_list');
    return { success: true, message: "Nama berhasil dihapus." };
  }
  return { success: false, message: "Nama tidak ditemukan." };
}

// =========================================================================
// ================ SERVER STATS UTAMA (HOME PAGE) =========================
// =========================================================================

function getHomePageStats() {
  const cache = CacheService.getScriptCache();
  const cachedStats = cache.get('stats_home');
  if (cachedStats) return JSON.parse(cachedStats);

  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  
  let teamCount = 0;
  let totalActivities = 0;
  
  sheets.forEach(sheet => {
    if (sheet.getName().toLowerCase().startsWith("tim ")) {
      teamCount++;
      const count = Math.max(0, sheet.getLastRow() - 1);
      totalActivities += count;
    }
  });

  const riskSheet = ss.getSheetByName(RISK_SHEET_NAME);
  const riskCount = riskSheet ? Math.max(0, riskSheet.getLastRow() - 1) : 0;

  const result = { teamCount, activityCount: totalActivities, riskCount };

  cache.put('stats_home', JSON.stringify(result), 3600);
  return result;
}