class KibE {
  constructor(sheetId) {
    this.sheet = SpreadsheetApp.openById(sheetId);
  }

  setChild(name){
    this.sheet.getSheetByName(name);
  }

  getKIBEByNo(no, entity) {
    this.sheet = this.sheet.getSheetByName(entity);
    const data = this.sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == no) { 
        return data[i]; 
      }
    }
    throw new Error(`KIB E dengan No ${no} tidak ditemukan`);
  }

  // Baca semua data barang
  getAll(entity) {
    const sheet = this.sheet.getSheetByName(entity);
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (sheet) {
      const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
      return data;
    }
  }
  
  getBarang(){
    const sheet = this.sheet.getSheetByName('Barang-E');
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (sheet) {
      const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
      return data;
    }
  }

  getKelompokByBarang(barang){
    // Mendapatkan sheet aktif (ganti dengan nama sheet yang sesuai jika perlu)
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kelompok-E");

    // Mengambil semua data dari sheet
    const data = sheet.getDataRange().getValues();

    // Ambil header untuk referensi
    const headers = data.shift(); // Baris pertama sebagai header

    // Menentukan indeks kolom berdasarkan header
    const barangIndex = headers.indexOf("Barang");
    const namaIndex = headers.indexOf("Nama");

    if (barangIndex === -1 || namaIndex === -1) {
      throw new Error("Kolom 'Barang' atau 'Nama' tidak ditemukan pada header.");
    }

    // Filter data berdasarkan barang
    const hasil = data
      .filter(row => row[barangIndex] === barang) // Hanya baris dengan barang sesuai
      .map(row => row[namaIndex]); // Ambil kolom Nama saja

    // Mengembalikan array Nama
    return hasil; 
  }

  getSubKelompokByKelompok(kelompok){
    // Mendapatkan sheet aktif (ganti dengan nama sheet yang sesuai jika perlu)
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sub Kelompok-E");

    // Mengambil semua data dari sheet
    const data = sheet.getDataRange().getValues();

    // Ambil header untuk referensi
    const headers = data.shift(); // Baris pertama sebagai header

    // Menentukan indeks kolom berdasarkan header
    const kelompokIndex = headers.indexOf("Kelompok");
    const namaIndex = headers.indexOf("Nama");

    if (kelompokIndex === -1 || namaIndex === -1) {
      throw new Error("Kolom 'Kelompok' atau 'Nama' tidak ditemukan pada header.");
    }

    // Filter data berdasarkan barang
    const hasil = data
      .filter(row => row[kelompokIndex] === kelompok) // Hanya baris dengan barang sesuai
      .map(row => row[namaIndex]); // Ambil kolom Nama saja

    // Mengembalikan array Nama
    return hasil; 
  }

  getJenisBySubKelompok(subkelompok){
    // Mendapatkan sheet aktif (ganti dengan nama sheet yang sesuai jika perlu)
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jenis-E");

    // Mengambil semua data dari sheet
    const data = sheet.getDataRange().getValues();

    // Ambil header untuk referensi
    const headers = data.shift(); // Baris pertama sebagai header

    // Menentukan indeks kolom berdasarkan header
    const subKelompokIndex = headers.indexOf("Sub Kelompok");
    const namaIndex = headers.indexOf("Nama");
    const kodeIndex = headers.indexOf("Kode");

    if (subKelompokIndex === -1 || namaIndex === -1) {
      throw new Error("Kolom 'Sub Kelompok' atau 'Nama' tidak ditemukan pada header.");
    }

    // Filter data berdasarkan barang
    const hasil = data
      .filter(row => row[subKelompokIndex] === subkelompok) // Hanya baris dengan barang sesuai
      .map(row => ({
        nama: row[namaIndex],      // Kolom Nama
        kode: row[kodeIndex]   // Kolom Kode
      })); 

    // Mengembalikan array Nama
    return hasil; 
  }
  
  capitalizeWords(str) {
    const exceptions = ["II", "III", "IV"]; // Daftar kata yang tidak boleh dikapitalisasi ulang

    return str.replace(/\w\S*/g, word => {
      // Jika kata ada dalam daftar pengecualian, kembalikan apa adanya
      if (exceptions.includes(word.toUpperCase())) {
        return word.toUpperCase(); // Pastikan kata tetap dalam format uppercase
      }
      // Kapitalisasi jika bukan termasuk pengecualian
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    });
  }

  getJenis(){
    const sheet = this.sheet.getSheetByName('Jenis-E');
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (sheet) {
      const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
      return data;
    }
  }

  create(barang, entity) {
    // Ambil semua data dari sheet
    this.sheet = this.sheet.getSheetByName(entity);
    Logger.log(`entity: ${entity}`);
    const data = this.sheet.getDataRange().getValues();
    const headers = data.shift(); 
    let register = null;
    let insertionIndex = data.length; 

    // Cari posisi penyisipan berdasarkan Sub Kelompok
    for (let i = 0; i < data.length; i++) {
      if (data[i][2] === barang.jenis_barang) {
        register = parseInt(data[i][3]) + 1; 
        insertionIndex = i + 1; // Sisipkan setelah baris terakhir dari Sub Kelompok yang sama
      }
    }
    if (register === null) {
      register = 1;
    }

    const formattedPerolehan = barang.tanggal_perolehan
      ? Utilities.formatDate(new Date(barang.tanggal_perolehan), Session.getScriptTimeZone(), "yyyy-MM-dd")
      : "";
    const formattedPenerimaan = barang.tanggal_penerimaan
      ? Utilities.formatDate(new Date(barang.tanggal_penerimaan), Session.getScriptTimeZone(), "yyyy-MM-dd")
      : "";
    const currentTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    // Tambahkan data baru
    const newNo = insertionIndex + 1; // No untuk data baru							
    const newData = [
      newNo, 
      barang.kode_barang, 
      barang.jenis_barang, 
      register,
      formattedPerolehan,
      formattedPenerimaan,
      barang.judul,
      barang.pencipta,
      barang.bahan,
      barang.ukuran,
      barang.asal_usul_daerah,
      barang.kondisi,
      barang.jumlah,
      barang.harga_satuan,
      barang.harga,
      barang.asal_usul,
      barang.keterangan,
      currentTime,
    ];
    data.splice(insertionIndex, 0, newData); // Sisipkan data baru
    data.sort((a, b) => {
      // Urutkan berdasarkan KODE BARANG
      if (a[1] < b[1]) return -1;
      if (a[1] > b[1]) return 1;
      // Jika KODE BARANG sama, urutkan berdasarkan REGISTER
      return parseInt(a[3]) - parseInt(b[3]);
    });

    // Perbarui nomor urut untuk semua baris setelah data baru
    for (let i = 0; i < data.length; i++) {
      data[i][0] = i + 1; // Perbarui kolom "No"
    }

    // Tulis ulang data ke sheet
    this.sheet.clear(); // Hapus data lama
    this.sheet.appendRow(headers); // Tulis ulang header
    this.sheet.getRange(2, 1, data.length, data[0].length).setValues(data); // Tulis ulang data

    return newData;
  }

  // Update barang berdasarkan No
  update(barang, entity) {
    this.sheet = this.sheet.getSheetByName(entity);
    const data = this.sheet.getDataRange().getValues();
    const currentTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == barang.no) {
        if(data[i][2] !== barang.jenis_barang){
          const headers = data.shift(); 
          let register = null;
          let insertionIndex = data.length; 

          // Cari posisi penyisipan berdasarkan Sub Kelompok
          for (let i = 0; i < data.length; i++) {
            if (data[i][2] === barang.jenis_barang) {
              register = parseInt(data[i][3]) + 1; 
              insertionIndex = i + 1; // Sisipkan setelah baris terakhir dari Sub Kelompok yang sama
            }
          }
          if (register === null) {
            register = 1;
          }

          const formattedPerolehan = barang.tanggal_perolehan
            ? Utilities.formatDate(new Date(barang.tanggal_perolehan), Session.getScriptTimeZone(), "yyyy-MM-dd")
            : "";
          const formattedPenerimaan = barang.tanggal_penerimaan
            ? Utilities.formatDate(new Date(barang.tanggal_penerimaan), Session.getScriptTimeZone(), "yyyy-MM-dd")
            : "";
          // Tambahkan data baru
          const newNo = insertionIndex + 1; // No untuk data baru
          const newData = [
              newNo, 
              barang.kode_barang, 
              barang.jenis_barang, 
              register,
              formattedPerolehan,
              formattedPenerimaan,
              barang.judul,
              barang.pencipta,
              barang.bahan,
              barang.ukuran,
              barang.asal_usul_daerah,
              barang.kondisi,
              barang.jumlah,
              barang.harga_satuan,
              barang.harga,
              barang.asal_usul,
              barang.keterangan,
              currentTime,
          ];
          data.splice(insertionIndex, 0, newData); // Sisipkan data baru
          data.sort((a, b) => {
            // Urutkan berdasarkan KODE BARANG
            if (a[1] < b[1]) return -1;
            if (a[1] > b[1]) return 1;
            // Jika KODE BARANG sama, urutkan berdasarkan REGISTER
            return parseInt(a[3]) - parseInt(b[3]);
          });

          // Perbarui nomor urut untuk semua baris setelah data baru
          for (let i = 0; i < data.length; i++) {
            data[i][0] = i + 1; // Perbarui kolom "No"
          }

          // Tulis ulang data ke sheet
          this.sheet.clear(); // Hapus data lama
          this.sheet.appendRow(headers); // Tulis ulang header
          this.sheet.getRange(2, 1, data.length, data[0].length).setValues(data); // Tulis ulang data

          return newData;

        }else{
          this.sheet.getRange(i + 1, 3).setValue(barang.jenis_barang);
          this.sheet.getRange(i + 1, 2).setValue(barang.kode_barang);
          // this.sheet.getRange(i + 1, 4).setValue(barang.register);
          this.sheet.getRange(i + 1, 5).setValue(barang.tanggal_perolehan);
          this.sheet.getRange(i + 1, 6).setValue(barang.tanggal_penerimaan);
          this.sheet.getRange(i + 1, 7).setValue(barang.judul);
          this.sheet.getRange(i + 1, 8).setValue(barang.pencipta);
          this.sheet.getRange(i + 1, 9).setValue(barang.bahan);
          this.sheet.getRange(i + 1, 10).setValue(barang.ukuran);
          this.sheet.getRange(i + 1, 11).setValue(barang.asal_usul_daerah);
          this.sheet.getRange(i + 1, 12).setValue(barang.kondisi);
          this.sheet.getRange(i + 1, 13).setValue(barang.jumlah);
          this.sheet.getRange(i + 1, 14).setValue(barang.harga_satuan);
          this.sheet.getRange(i + 1, 15).setValue(barang.harga);
          this.sheet.getRange(i + 1, 16).setValue(barang.asal_usul);
          this.sheet.getRange(i + 1, 17).setValue(barang.keterangan);
          this.sheet.getRange(i + 1, 18).setValue(currentTime);
          return barang;
        }
      }
    }
    throw new Error(`Barang dengan No ${no} tidak ditemukan`);
  }

  // Hapus barang berdasarkan No
  delete(no, entity) {
    this.sheet = this.sheet.getSheetByName(entity);
    const data = this.sheet.getDataRange().getValues();
    const headers = data.shift(); // Ambil header

    // Hapus baris secara terbalik
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] == no) {
        this.sheet.deleteRow(i + 2); // +2 karena header di baris pertama
        data.splice(i, 1); // Hapus juga dari array
      }
    }

    // Perbarui kolom "No"
    for (let i = 0; i < data.length; i++) {
      data[i][0] = i + 1; // Perbarui nomor urut
    }

    // Tulis ulang data ke sheet
    this.sheet.clear(); // Hapus data lama
    this.sheet.appendRow(headers); // Tulis ulang header
    if (data.length > 0) {
      this.sheet.getRange(2, 1, data.length, data[0].length).setValues(data); // Tulis ulang data
    }

    return true;
  }
}