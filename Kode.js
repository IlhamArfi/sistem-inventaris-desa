const SPREADSHEET_ID = '1XgN9R_2QAEdB14sJvaF1JP0vIhJ-tTDsg44Q5OygxH8'; 

function getSummaryData() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); 
  const kibACount = spreadsheet.getSheetByName('Tanah').getLastRow() - 1;

  const kibBCount = (spreadsheet.getSheetByName('Besar').getLastRow() - 1) + (spreadsheet.getSheetByName('Angkutan').getLastRow() - 1) + (spreadsheet.getSheetByName('Bengkel').getLastRow() - 1) + (spreadsheet.getSheetByName('Pertanian').getLastRow() - 1) + (spreadsheet.getSheetByName('RumahTangga').getLastRow() - 1) + (spreadsheet.getSheetByName('AlatStudio').getLastRow() - 1) + (spreadsheet.getSheetByName('AlatKedokteran').getLastRow() - 1) + (spreadsheet.getSheetByName('AlatLaboratorium').getLastRow() - 1) + (spreadsheet.getSheetByName('Persenjataan').getLastRow() - 1) + (spreadsheet.getSheetByName('Komputer').getLastRow() - 1) + (spreadsheet.getSheetByName('Eksplorasi').getLastRow() - 1) + (spreadsheet.getSheetByName('Pengeboran').getLastRow() - 1) + (spreadsheet.getSheetByName('Produksi').getLastRow() - 1) + (spreadsheet.getSheetByName('AlatBantuEksplorasi').getLastRow() - 1) + (spreadsheet.getSheetByName('AlatKeselamatanKerja').getLastRow() - 1) + (spreadsheet.getSheetByName('AlatPeraga').getLastRow() - 1) + (spreadsheet.getSheetByName('PeralatanProsesProduksi').getLastRow() - 1) + (spreadsheet.getSheetByName('RambuRambu').getLastRow() - 1) + (spreadsheet.getSheetByName('PeralatanOlahraga').getLastRow() - 1);

  const kibCCount = (spreadsheet.getSheetByName('BangunanGedung').getLastRow() - 1) + (spreadsheet.getSheetByName('Monumen').getLastRow() - 1) + (spreadsheet.getSheetByName('BangunanMenara').getLastRow() - 1) + (spreadsheet.getSheetByName('TuguTitikKontrol').getLastRow() - 1);
  
  const kibDCount = (spreadsheet.getSheetByName('Jalan').getLastRow() - 1) + (spreadsheet.getSheetByName('BangunanAir').getLastRow() - 1) + (spreadsheet.getSheetByName('Instalasi').getLastRow() - 1) + (spreadsheet.getSheetByName('Jaringan').getLastRow() - 1);
            
  const kibECount = (spreadsheet.getSheetByName('BahanPerpustakaan').getLastRow() - 1) + (spreadsheet.getSheetByName('BarangBercorakKesenian').getLastRow() - 1) + (spreadsheet.getSheetByName('Hewan').getLastRow() - 1) + (spreadsheet.getSheetByName('BiotaPerairan').getLastRow() - 1) + (spreadsheet.getSheetByName('Tanaman').getLastRow() - 1) + (spreadsheet.getSheetByName('BarangKoleksiNonbudaya').getLastRow() - 1) + (spreadsheet.getSheetByName('AsetTetapDalamRenovasi').getLastRow() - 1);

  const kibFCount = spreadsheet.getSheetByName('KontruksiDalamPengerjaan').getLastRow() - 1;

  data = {
    kibACount: kibACount,
    kibBCount: kibBCount,
    kibCCount: kibCCount,
    kibDCount: kibDCount,
    kibECount: kibECount,
    kibFCount: kibFCount,
  };
  console.log(data)
  return data;
}

function getPageTemplate(pageName) {
  return HtmlService.createHtmlOutputFromFile(pageName).getContent();
}

function myURL() {
   return ScriptApp.getService().getUrl();
}

function getNewHtml(e) {
  var html = HtmlService.createTemplateFromFile(e.parameter.page).evaluate().getContent();
  return html;
}

function include(filename) {
   return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

// Fungsi untuk mencoba berbagai format tanggal
function parseDate(dateString) {
    let date = new Date(dateString);
    if (!isNaN(date)) return setStartOfDay(date);

    const parts = dateString.split('/');
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      const year = parseInt(parts[2], 10);
      date = new Date(year, month, day);
      if (!isNaN(date)) return setStartOfDay(date);
    }

    return null;
}

function setStartOfDay(date) {
  if (!date || isNaN(date)) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function generateReport(data){
  const kiba = ["Tanah"];
  const kibb = ["PeralatanOlahraga", "Besar", "Angkutan", "Bengkel", "Pertanian", "RumahTangga", "AlatStudio", "AlatKedokteran", "AlatLaboratorium", "Persenjataan", "Komputer", "Eksplorasi", "Pengeboran", "Produksi", "AlatBantuEksplorasi", "AlatKeselamatanKerja", "AlatPeraga", "PeralatanProsesProduksi", "RambuRambu"];
  const kibc = ["BangunanGedung", "Monumen", "BangunanMenara", "TuguTitikKontrol"];  
  const kibd = ["Jalan", "BangunanAir", "Instalasi", "Jaringan"];
  const kibe = ["BahanPerpustakaan", "BarangBercorakKesenian", "Hewan", "BiotaPerairan", "Tanaman", "BarangKoleksiNonbudaya", "AsetTetapDalamRenovasi"];
  const kibf = ["KontruksiDalamPengerjaan"];

  const start = new Date(data.start);
  const end = new Date(data.end);
  const jenis = data.jenis;
  
  let selectedSheets = [];
  
  // Tentukan sheet berdasarkan jenis
  switch (jenis) {
    case 'a':
      selectedSheets = kiba;
      break;
    case 'b':
      selectedSheets = kibb;
      break;
    case 'c':
      selectedSheets = kibc;
      break;
    case 'd':
      selectedSheets = kibd;
      break;
    case 'e':
      selectedSheets = kibe;
      break;
    case 'f':
      selectedSheets = kibf;
      break;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let reportData = [];

  // Ambil data dari sheet yang dipilih
  selectedSheets.forEach(sheetName => {
    Logger.log("sheetName " + sheetName);
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      const headers = data[0]; // Asumsi baris pertama adalah header
      const dateColIndex = headers.indexOf("TANGGAL PENERIMAAN");
      
      if (dateColIndex !== -1) {
        const filteredData = data.slice(1).filter(row => {
          const date = new Date(row[dateColIndex]);
          return date >= start && date <= end;
        });
        Logger.log("filteredData " + filteredData);
        if (filteredData.length > 0) {
          reportData.push({
            sheetName,
            headers: headers.slice(0, 19),
            data: filteredData.map(row => row.slice(0, 19))
          });
        }
      }
    }
  });
  Logger.log("reportData" + reportData);
  Logger.log(reportData);
  
  return JSON.stringify(reportData);
}

function KibPageService(kibType, entity) {
  // Pemetaan data handler berdasarkan kibType
  const dataHandlers = {
    KIBA: this.getAllDataKIBA,
    KIBB: this.getAllDataKIBB,
    KIBC: this.getAllDataKIBC,
    KIBD: this.getAllDataKIBD,
    KIBE: this.getAllDataKIBE,
    KIBF: this.getAllDataKIBF,
  };

  // Validasi kibType
  const handler = dataHandlers[kibType];
  if (!handler) {
    throw new Error(`Jenis KIB "${kibType}" tidak dikenali.`);
  }

  // Ambil data menggunakan handler
  const kibData = JSON.stringify(handler.call(this, entity));

  // Load template
  let template;
  try {
    template = HtmlService.createTemplateFromFile(kibType);
  } catch (error) {
    throw new Error(`Template untuk kibType "${kibType}" tidak ditemukan.`);
  }

  // Set data ke template
  template.data = kibData;
  template.kibType = kibType;
  template.barangKIB = JSON.stringify(this.getBarangKIB(kibType));
  template.asalUsulData = JSON.stringify(this.getAsalUsul());
  template.modelType = entity;

  // Return template yang sudah di-render
  return template
    .evaluate()
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doGet(e) {
  let page = e.parameter.page || "Index";
  let entity = e.parameter.entity;
  if (page === "KIBA") {
    const entityHandlers = {
      Tanah: () => {
        // Logika untuk "Tanah"
        return this.KibPageService("KIBA", "Tanah");
      },
    }
    const handler = entityHandlers[entity];
    if (handler) {
      return handler();
    } else {
      Logger.log(`Entity "${entity}" tidak dikenali.`);
    }
  } else if (page === 'KIBB') {
    const entityHandlers = {
      Besar: () => {
        // Logika untuk "Besar"
        return this.KibPageService("KIBB", "Besar");
      },
      Angkutan: () => {
        // Logika untuk "Angkutan"
        return this.KibPageService("KIBB", "Angkutan");
      },
      Bengkel: () => {
        // Logika untuk "Bengkel"
        return this.KibPageService("KIBB", "Bengkel");
      },
      Pertanian: () => {
        // Logika untuk "Pertanian"
        return this.KibPageService("KIBB", "Pertanian");
      },
      RumahTangga: () => {
        // Logika untuk "Rumah Tangga"
        return this.KibPageService("KIBB", "RumahTangga");
      },
      AlatStudio: () => {
        // Logika untuk "Alat Studio"
        return this.KibPageService("KIBB", "AlatStudio");
      },
      AlatKedokteran: () => {
        // Logika untuk "Alat Kedokteran"
        return this.KibPageService("KIBB", "AlatKedokteran");
      },
      AlatLaboratorium: () => {
        // Logika untuk "Alat Laboratorium"
        return this.KibPageService("KIBB", "AlatLaboratorium");
      },
      Persenjataan: () => {
        // Logika untuk "Persenjataan"
        return this.KibPageService("KIBB", "Persenjataan");
      },
      Komputer: () => {
        // Logika untuk "Komputer"
        return this.KibPageService("KIBB", "Komputer");
      },
      Eksplorasi: () => {
        // Logika untuk "Eksplorasi"
        return this.KibPageService("KIBB", "Eksplorasi");
      },
      Pengeboran: () => {
        // Logika untuk "Pengeboran"
        return this.KibPageService("KIBB", "Pengeboran");
      },
      Produksi: () => {
        // Logika untuk "Produksi"
        return this.KibPageService("KIBB", "Produksi");
      },
      AlatBantuEksplorasi: () => {
        // Logika untuk "Alat Bantu Eksplorasi"
        return this.KibPageService("KIBB", "AlatBantuEksplorasi");
      },
      AlatKeselamatanKerja: () => {
        // Logika untuk "Alat Keselamatan Kerja"
        return this.KibPageService("KIBB", "AlatKeselamatanKerja");
      },
      AlatPeraga: () => {
        // Logika untuk "Alat Peraga"
        return this.KibPageService("KIBB", "AlatPeraga");
      },
      PeralatanProsesProduksi: () => {
        // Logika untuk "Peralatan Proses Produksi"
        return this.KibPageService("KIBB", "PeralatanProsesProduksi");
      },
      RambuRambu: () => {
        // Logika untuk "Rambu-Rambu"
        return this.KibPageService("KIBB", "RambuRambu");
      },
      PeralatanOlahraga: () => {
        // Logika untuk "Peralatan Olahraga"
        return this.KibPageService("KIBB", "PeralatanOlahraga");
      },
    };

    const handler = entityHandlers[entity];
    if (handler) {
      return handler();
    } else {
      Logger.log(`Entity "${entity}" tidak dikenali.`);
    }
  } else if (page === 'KIBC') {
    const entityHandlers = {
      BangunanGedung: () => {
        // Logika untuk "BangunanGedung"
        return this.KibPageService("KIBC", "BangunanGedung");
      },
      Monumen: () => {
        // Logika untuk "Monumen"
        return this.KibPageService("KIBC", "Monumen");
      },
      BangunanMenara: () => {
        // Logika untuk "BangunanMenara"
        return this.KibPageService("KIBC", "BangunanMenara");
      },
      TuguTitikKontrol: () => {
        // Logika untuk "TuguTitikKontrol"
        return this.KibPageService("KIBC", "TuguTitikKontrol");
      },
    };

    const handler = entityHandlers[entity];
    if (handler) {
      return handler();
    } else {
      Logger.log(`Entity "${entity}" tidak dikenali.`);
    }
  } else if (page === 'KIBD') {
    const entityHandlers = {
      Jalan: () => {
        // Logika untuk "Jalan"
        return this.KibPageService("KIBD", "Jalan");
      },
      BangunanAir: () => {
        // Logika untuk "BangunanAir"
        return this.KibPageService("KIBD", "BangunanAir");
      },
      Instalasi: () => {
        // Logika untuk "Instalasi"
        return this.KibPageService("KIBD", "Instalasi");
      },
      Jaringan: () => {
        // Logika untuk "Jaringan"
        return this.KibPageService("KIBD", "Jaringan");
      },
    };

    const handler = entityHandlers[entity];
    if (handler) {
      return handler();
    } else {
      Logger.log(`Entity "${entity}" tidak dikenali.`);
    }
  } else if (page === 'KIBE') {
    const entityHandlers = {
      BahanPerpustakaan: () => {
        // Logika untuk "BahanPerpustakaan"
        return this.KibPageService("KIBE", "BahanPerpustakaan");
      },
      BarangBercorakKesenian: () => {
        // Logika untuk "BarangBercorakKesenian"
        return this.KibPageService("KIBE", "BarangBercorakKesenian");
      },
      Hewan: () => {
        // Logika untuk "Hewan"
        return this.KibPageService("KIBE", "Hewan");
      },
      BiotaPerairan: () => {
        // Logika untuk "BiotaPerairan"
        return this.KibPageService("KIBE", "BiotaPerairan");
      },
      Tanaman: () => {
        // Logika untuk "Tanaman"
        return this.KibPageService("KIBE", "Tanaman");
      },
      BarangKoleksiNonbudaya: () => {
        // Logika untuk "BarangKoleksiNonbudaya"
        return this.KibPageService("KIBE", "BarangKoleksiNonbudaya");
      },
      AsetTetapDalamRenovasi: () => {
        // Logika untuk "AsetTetapDalamRenovasi"
        return this.KibPageService("KIBE", "AsetTetapDalamRenovasi");
      },
    };

    const handler = entityHandlers[entity];
    if (handler) {
      return handler();
    } else {
      Logger.log(`Entity "${entity}" tidak dikenali.`);
    }
  } else if (page === 'KIBF') {
    const entityHandlers = {
      KontruksiDalamPengerjaan: () => {
        // Logika untuk "KontruksiDalamPengerjaan"
        return this.KibPageService("KIBF", "KontruksiDalamPengerjaan");
      },
    };

    const handler = entityHandlers[entity];
    if (handler) {
      return handler();
    } else {
      Logger.log(`Entity "${entity}" tidak dikenali.`);
    }
  } else {
    return HtmlService.createTemplateFromFile(page)
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function getAsalUsul(){
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName('Asal Usul');
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (sheet) {
    const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    return data;
  }
}

function getBarangKIB(kibType) {
  const kibClasses = {
    KIBA: KibA,
    KIBB: KibB,
    KIBC: KibC,
    KIBD: KibD,
    KIBE: KibE,
    KIBF: KibF,
  };

  try {
    const KibClass = kibClasses[kibType];
    if (!KibClass) {
      throw new Error(`Barang KIB "${kibType}" tidak dikenali.`);
    }
    return new KibClass(SPREADSHEET_ID).getBarang();
  } catch (error) {
    Logger.log(error.message);
    return null;
  }
}

function getKelompokByBarangKIB(kibType, barang) {
    Logger.log(kibType, barang);
    Logger.log(barang);
  const kibClasses = {
    KIBA: KibA,
    KIBB: KibB,
    KIBC: KibC,
    KIBD: KibD,
    KIBE: KibE,
    KIBF: KibF,
  };

  try {
    const KibClass = kibClasses[kibType];
    if (!KibClass) {
      throw new Error(`Barang KIB "${kibType}" tidak dikenali.`);
    }
    return new KibClass(SPREADSHEET_ID).getKelompokByBarang(barang);
  } catch (error) {
    Logger.log(error.message);
    return null;
  }
}

function getSubKelompokByKelompokKIB(kibType, kelompok){
  const kibClasses = {
    KIBA: KibA,
    KIBB: KibB,
    KIBC: KibC,
    KIBD: KibD,
    KIBE: KibE,
    KIBF: KibF,
  };

  try {
    const KibClass = kibClasses[kibType];
    if (!KibClass) {
      throw new Error(`Sub Kelompok KIB "${kibType}" tidak dikenali.`);
    }
    return new KibClass(SPREADSHEET_ID).getSubKelompokByKelompok(kelompok);
  } catch (error) {
    Logger.log(error.message);
    return null;
  }
}

function getJenisBySubKelompokKIB(kibType, subkelompok){
  const kibClasses = {
    KIBA: KibA,
    KIBB: KibB,
    KIBC: KibC,
    KIBD: KibD,
    KIBE: KibE,
    KIBF: KibF,
  };

  try {
    const KibClass = kibClasses[kibType];
    if (!KibClass) {
      throw new Error(`Jenis KIB "${kibType}" tidak dikenali.`);
    }
    return new KibClass(SPREADSHEET_ID).getJenisBySubKelompok(subkelompok);
  } catch (error) {
    Logger.log(error.message);
    return null;
  }
}

function getJenisKIB(kibType) {
  const kibClasses = {
    KIBA: KibA,
    KIBB: KibB,
    KIBC: KibC,
    KIBD: KibD,
    KIBE: KibE,
    KIBF: KibF,
  };

  try {
    const KibClass = kibClasses[kibType];
    if (!KibClass) {
      throw new Error(`Jenis KIB "${kibType}" tidak dikenali.`);
    }
    return new KibClass(SPREADSHEET_ID).getJenis();
  } catch (error) {
    Logger.log(error.message);
    return null;
  }
}

// ------------------------------------------------------------
// KIBA Model Operation
function getAllDataKIBA(entity) {
  const kibA = new KibA(SPREADSHEET_ID);
  
  try {
    const data = kibA.getAll(entity); 
    return data.map(row => ({
      no: row[0],
      jenis_barang: row[1],
      jenis_barang: row[1],
      kode_barang: row[2],
      register: row[3],
      tanggal_pengadaan: row[4],
      tanggal_penerimaan: row[5],
      asal_usul: row[6],
      letak: row[7],
      penggunaan: row[8],
      kondisi: row[9],
      no_sertifikat: row[10],
      luas_m2: row[11],
      harga_per_meter: row[12],
      jumlah_harga: row[13],
      keterangan: row[14],
      last_update: row[15]
    }));
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    return [];
  }
}

function addKIBA(data, entity){
  let jumlah = data.jumlah_barang;
  for(let i=0; i <jumlah; i++){
    let kibA = new KibA(SPREADSHEET_ID);
    kibA.create(data, entity);
  }
  return JSON.stringify(this.getAllDataKIBA(entity));
}

function getKIBAbyNo(no, entity){
  const kibA = new KibA(SPREADSHEET_ID);
  const data = kibA.getKIBAByNo(no, entity); 
  return JSON.stringify(data);
}

function updateKIBA(data, entity){
  let kibA = new KibA(SPREADSHEET_ID);
  kibA.update(data, entity);
  return JSON.stringify(this.getAllDataKIBA(entity));
}

function deleteKIBA(no, entity){  
  let kibA = new KibA(SPREADSHEET_ID);
  kibA.delete(no, entity);
  return JSON.stringify(this.getAllDataKIBA(entity));
}

// ------------------------------------------------------------
// KIBB Model Operation
function getAllDataKIBB(entity) {
  const kibB = new KibB(SPREADSHEET_ID);
  try {
    const data = kibB.getAll(entity); 
    return data.map(row => ({
      no: row[0],
      kode_barang: row[1],
      jenis_barang: row[2],
      register: row[3],
      tanggal_pengadaan: row[4],
      tanggal_penerimaan: row[5],
      merk: row[6],
      type: row[7],
      ukuran: row[8],
      bahan: row[9],
      no_pabrik: row[10],
      no_rangka: row[11],
      no_mesin: row[12],
      no_polisi: row[13],
      no_bpkb: row[14],
      kondisi: row[15],
      asal_usul: row[16],
      harga: row[17],
      keterangan: row[18],
      last_update: row[19]
    }));
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    return [];
  }
}

function addKIBB(data, entity){
  let jumlah = data.jumlah_barang;
  for(let i=0; i <jumlah; i++){
    let kibB = new KibB(SPREADSHEET_ID);
    kibB.create(data, entity);
  }
  return JSON.stringify(this.getAllDataKIBB(entity));
}

function getKIBBbyNo(no, entity){
  const kibB = new KibB(SPREADSHEET_ID);
  const data = kibB.getKIBBByNo(no, entity); 
  return JSON.stringify(data);
}

function updateKIBB(data, entity){
  const kibB = new KibB(SPREADSHEET_ID);
  kibB.update(data, entity);
  return JSON.stringify(this.getAllDataKIBB(entity));
}

function deleteKIBB(no, entity){  
  const kibB = new KibB(SPREADSHEET_ID);
  kibB.delete(no, entity);
  return JSON.stringify(this.getAllDataKIBB(entity));
}

// ------------------------------------------------------------
// KIBC Model Operation
function getAllDataKIBC(entity) {
  const kibC = new KibC(SPREADSHEET_ID);
  try {
    const data = kibC.getAll(entity); 
    return data.map(row => ({
      no: row[0],
      kode_barang: row[1],
      jenis_barang: row[2],
      register: row[3],
      tanggal_perolehan: row[4],
      tanggal_penerimaan: row[5],
      bertingkat: row[6],
      beton: row[7],
      luas_lantai_m2: row[8],
      letak: row[9],
      status_tanah: row[10],
      kondisi: row[11],
      asal_usul: row[12],
      harga: row[13],
      keterangan: row[14],
      last_update: row[15]
    }));
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    return [];
  }
}

function addKIBC(data, entity){
  let jumlah = data.jumlah_barang;
  for(let i=0; i <jumlah; i++){
    let kibC = new KibC(SPREADSHEET_ID);
    kibC.create(data, entity);
  }
  return JSON.stringify(this.getAllDataKIBC(entity));
}

function getKIBCbyNo(no, entity){
  const kibC = new KibC(SPREADSHEET_ID);
  const data = kibC.getKIBCByNo(no, entity); 
  return JSON.stringify(data);
}

function updateKIBC(data, entity){
  const kibC = new KibC(SPREADSHEET_ID);
  kibC.update(data, entity);
  return JSON.stringify(this.getAllDataKIBC(entity));
}

function deleteKIBC(no, entity){  
  const kibC = new KibC(SPREADSHEET_ID);
  kibC.delete(no, entity);
  return JSON.stringify(this.getAllDataKIBC(entity));
}

// ------------------------------------------------------------
// KIBD Model Operation
function getAllDataKIBD(entity) {
  const kibD = new KibD(SPREADSHEET_ID);
  try {
    const data = kibD.getAll(entity); 
    return data.map(row => ({
      no: row[0],
      kode_barang: row[1],
      jenis_barang: row[2],
      register: row[3],
      tanggal_perolehan: row[4],
      tanggal_penerimaan: row[5],
      kontruksi: row[6],
      panjang_km: row[7],
      lebar_m: row[8],
      luas_m2: row[9],
      letak: row[10],
      status_tanah: row[11],
      kondisi: row[12],
      asal_usul: row[13],
      harga: row[14],
      keterangan: row[15],
      last_update: row[16]
    }));
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    return [];
  }
}

function addKIBD(data, entity){
  let jumlah = data.jumlah_barang;
  for(let i=0; i <jumlah; i++){
    let kibD = new KibD(SPREADSHEET_ID);
    kibD.create(data, entity);
  }
  return JSON.stringify(this.getAllDataKIBD(entity));
}

function getKIBDbyNo(no, entity){
  const kibD = new KibD(SPREADSHEET_ID);
  const data = kibD.getKIBDByNo(no, entity); 
  return JSON.stringify(data);
}

function updateKIBD(data, entity){
  const kibD = new KibD(SPREADSHEET_ID);
  kibD.update(data, entity);
  return JSON.stringify(this.getAllDataKIBD(entity));
}

function deleteKIBD(no, entity){  
  const kibD = new KibD(SPREADSHEET_ID);
  kibD.delete(no, entity);
  return JSON.stringify(this.getAllDataKIBD(entity));
}

// KIBE Model Operation
function getAllDataKIBE(entity) {
  const kibE = new KibE(SPREADSHEET_ID);
  try {
    const data = kibE.getAll(entity); 
    return data.map(row => ({
      no: row[0],
      kode_barang: row[1],
      jenis_barang: row[2],
      register: row[3],
      tanggal_perolehan: row[4],
      tanggal_penerimaan: row[5],
      judul: row[6],
      pencipta: row[7],
      bahan: row[8],
      ukuran: row[9],
      asal_usul_daerah: row[10],
      kondisi: row[11],
      jumlah: row[12],
      harga_satuan: row[13],
      harga: row[14],
      asal_usul: row[15],
      keterangan: row[16],
      last_update: row[17]
    }));
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    return [];
  }
}

function addKIBE(data, entity){
  let jumlah = data.jumlah_barang;
  for(let i=0; i <jumlah; i++){
    let kibE = new KibE(SPREADSHEET_ID);
    kibE.create(data, entity);
  }
  return JSON.stringify(this.getAllDataKIBE(entity));
}

function getKIBEbyNo(no, entity){
  const kibE = new KibE(SPREADSHEET_ID);
  const data = kibE.getKIBEByNo(no, entity); 
  return JSON.stringify(data);
}

function updateKIBE(data, entity){
  const kibE = new KibE(SPREADSHEET_ID);
  kibE.update(data, entity);
  return JSON.stringify(this.getAllDataKIBE(entity));
}

function deleteKIBE(no, entity){  
  const kibE = new KibE(SPREADSHEET_ID);
  kibE.delete(no, entity);
  return JSON.stringify(this.getAllDataKIBE(entity));
}


// ------------------------------------------------------------
// KIBF Model Operation
function getAllDataKIBF(entity) {
  const kibD = new KibD(SPREADSHEET_ID);
  try {
    const data = kibD.getAll(entity); 
    return data.map(row => ({
      no: row[0],
      kode_barang: row[1],
      jenis_barang: row[2],
      register: row[3],
      tanggal_perolehan: row[4],
      tanggal_penerimaan: row[5],
      bertingkat: row[6],
      beton: row[7],
      panjang_km: row[7],
      lebar_m: row[9],
      luas_m2: row[10],
      letak: row[11],
      status_tanah: row[12],
      kondisi: row[13],
      asal_usul: row[14],
      harga: row[15],
      keterangan: row[16],
      last_update: row[17]
    }));
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    return [];
  }
}

function addKIBF(data, entity){
  let jumlah = data.jumlah_barang;
  for(let i=0; i <jumlah; i++){
    let kibF = new KibF(SPREADSHEET_ID);
    kibF.create(data, entity);
  }
  return JSON.stringify(this.getAllDataKIBF(entity));
}

function getKIBFbyNo(no, entity){
  const kibF = new KibF(SPREADSHEET_ID);
  const data = kibF.getKIBFByNo(no, entity); 
  return JSON.stringify(data);
}

function updateKIBF(data, entity){
  const kibF = new KibF(SPREADSHEET_ID);
  kibF.update(data, entity);
  return JSON.stringify(this.getAllDataKIBF(entity));
}

function deleteKIBF(no, entity){  
  const kibF = new KibF(SPREADSHEET_ID);
  kibF.delete(no, entity);
  return JSON.stringify(this.getAllDataKIBF(entity));
}

// ------------------------------------------------------------
// Barang Model Operation
function getAllBarang() {
  const barangModel = new Barang(SPREADSHEET_ID);
  return barangModel.getAll();
}

function addBarang(barang) {
  const barangModel = new Barang(SPREADSHEET_ID);
  barangModel.create(barang.nama);
}

function getBarangByNo(no) {
  const barangModel = new Barang(SPREADSHEET_ID);
  return barangModel.getBarangByNo(no);
}

function updateBarang(barang) {
  const barangModel = new Barang(SPREADSHEET_ID);
  try {
    return barangModel.update(barang.no, barang.nama);
  } catch (error) {
    throw new Error(error.message);
  }
}

function deleteBarang(no) {
  const barangModel = new Barang(SPREADSHEET_ID);
  try {
    return barangModel.delete(no);
  } catch (error) {
    throw new Error(error.message);
  }
}

// Kelompok Model Operation
function getAllKelompok() {
  const kelompokModel = new Kelompok(SPREADSHEET_ID);
  return kelompokModel.getAll();
}

function addKelompok(kelompok) {
  const kelompokModel = new Kelompok(SPREADSHEET_ID);
  kelompokModel.create(kelompok.barang, kelompok.nama);
}

function getKelompokByNo(no) {
  const kelompokModel = new Kelompok(SPREADSHEET_ID);
  return kelompokModel.getKelompokByNo(no);
}

function updateKelompok(kelompok) {
  const kelompokModel = new Kelompok(SPREADSHEET_ID);
  try {
    return kelompokModel.update(kelompok.no, kelompok.barang, kelompok.nama);
  } catch (error) {
    throw new Error(error.message);
  }
}

function deleteKelompok(no) {
  const kelompokModel = new Kelompok(SPREADSHEET_ID);
  try {
    return kelompokModel.delete(no);
  } catch (error) {
    throw new Error(error.message);
  }
}

// Sub Kelompok Model Operation
function getAllSubKelompok() {
  const subKelompokModel = new SubKelompok(SPREADSHEET_ID);
  return subKelompokModel.getAll();
}

function addSubKelompok(subKelompok) {
  const subKelompokModel = new SubKelompok(SPREADSHEET_ID);
  subKelompokModel.create(subKelompok.kelompok, subKelompok.nama);
}

function getSubKelompokByNo(no) {
  const subKelompokModel = new SubKelompok(SPREADSHEET_ID);
  return subKelompokModel.getSubKelompokByNo(no);
}

function updateSubKelompok(subKelompok) {
  const subKelompokModel = new SubKelompok(SPREADSHEET_ID);
  try {
    return subKelompokModel.update(subKelompok.no, subKelompok.kelompok, subKelompok.nama);
  } catch (error) {
    throw new Error(error.message);
  }
}

function deleteSubKelompok(no) {
  const subKelompokModel = new SubKelompok(SPREADSHEET_ID);
  try {
    return subKelompokModel.delete(no);
  } catch (error) {
    throw new Error(error.message);
  }
}

// Jenis Model Operation
function getAllJenis() {
  const jenisModel = new Jenis(SPREADSHEET_ID);
  return jenisModel.getAll();
}

function addJenis(jenis) {
  const jenisModel = new Jenis(SPREADSHEET_ID);
  jenisModel.create(jenis.subKelompok, jenis.nama, jenis.kode);
}

function getKodeBySubKelompok(subKelompok){
  const jenisModel = new Jenis(SPREADSHEET_ID);
  return jenisModel.generateKode(subKelompok);
}

function getJenisByNo(no) {
  const jenisModel = new Jenis(SPREADSHEET_ID);
  return jenisModel.getJenisByNo(no);
}

function updateJenis(jenis) {
  const jenisModel = new Jenis(SPREADSHEET_ID);
  try {
    return jenisModel.update(jenis.no, jenis.subKelompok, jenis.nama, jenis.kode);
  } catch (error) {
    throw new Error(error.message);
  }
}

function deleteJenis(no) {
  const jenisModel = new Jenis(SPREADSHEET_ID);
  try {
    return jenisModel.delete(no);
  } catch (error) {
    throw new Error(error.message);
  }
}