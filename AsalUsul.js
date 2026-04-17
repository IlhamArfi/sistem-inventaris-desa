class AsalUsul {
  constructor(sheetId) {
    this.sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Asal Usul');
  }

  // Ambil semua data Asal Usul
  getAll() {
    const data = this.sheet.getDataRange().getValues();
    return data.slice(1).map(row => ({
      no: row[0],
      nama: row[1],
    }));
  }

  // Tambah data Asal Usul baru
  create(nama) {
    const no = this.generateNo();
    this.sheet.appendRow([no, nama]);
    return { no, nama };
  }

  // Update data Asal Usul berdasarkan No
  update(no, nama) {
    const data = this.sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == no) {
        this.sheet.getRange(i + 1, 2).setValue(nama);
        return { no, nama };
      }
    }
    throw new Error(`Asal Usul dengan No ${no} tidak ditemukan`);
  }

  // Hapus data Asal Usul berdasarkan No
  delete(no) {
    const data = this.sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == no) {
        this.sheet.deleteRow(i + 1);
        return true;
      }
    }
    throw new Error(`Asal Usul dengan No ${no} tidak ditemukan`);
  }

  // Generate No unik untuk Asal Usul
  generateNo() {
    const data = this.sheet.getDataRange().getValues();
    return data.length; // Menggunakan jumlah baris sebagai No baru
  }
}
