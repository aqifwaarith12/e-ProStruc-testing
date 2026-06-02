function doGet() {
  // Jalankan semakan dan bina tab secara automatik sebelum paparkan halaman
  binaTabSecaraAutomatik();

  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('PRO STRUC Builder')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// FUNGSI AUTOMASI: Membina tab dan header sekiranya tiada dalam Google Sheets
function binaTabSecaraAutomatik() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Semak & Bina Tab 'Pengguna'
  var sheetPengguna = ss.getSheetByName('Pengguna');
  if (!sheetPengguna) {
    sheetPengguna = ss.insertSheet('Pengguna');
    // Bina header: Nama, Email, Username, Password
    sheetPengguna.getRange("A1:D1").setValues([["Nama", "Email", "Username", "Password"]]);
    sheetPengguna.getRange("A1:D1").setFontWeight("bold").setBackground("#e0e7ff"); 
    // Tambah satu akaun contoh supaya tak kosong semasa test
    sheetPengguna.getRange("A2:D2").setValues([["Aqif Waarith", "aqif@example.com", "aqif123", "123456"]]);
  }
  
  // 2. Semak & Bina Tab 'Soalan'
  var sheetSoalan = ss.getSheetByName('Soalan');
  if (!sheetSoalan) {
    sheetSoalan = ss.insertSheet('Soalan');
    // Bina header rujukan soalan
    sheetSoalan.getRange("A1:E1").setValues([["Tingkatan", "Stimulus", "Soalan A", "Soalan B", "Soalan C"]]);
    sheetSoalan.getRange("A1:E1").setFontWeight("bold").setBackground("#fdf4ff");
    // Tambah satu soalan contoh
    sheetSoalan.getRange("A2:E2").setValues([[
      "Tingkatan 4", 
      "Kerajaan Agraria mempunyai sistem pengairan yang maju.", 
      "Jelaskan kegiatan ekonomi masyarakat Agraria.", 
      "Terangkan faktor-faktor perkembangan sistem pengairan.", 
      "Pada pendapat anda, apakah usaha-usaha mengekalkan warisan sejarah?"
    ]]);
  }
}

// Fungsi mendapatkan koleksi soalan (Membaca data)
function dapatkanKoleksiSoalan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Soalan'); 
  
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === "") continue;
    hasil.push({
      id: "Q" + i,
      tingkatan: data[i][0], 
      stimulus: data[i][1],  
      soalanA: data[i][2],   
      soalanB: data[i][3],   
      soalanC: data[i][4]    
    });
  }
  return hasil;
}

// Fungsi semak log masuk pelajar
function semakLogMasuk(usernameInput, passwordInput) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Pengguna');
  
  if (!sheet) {
    return { sah: false, mesej: "Ralat: Tab 'Pengguna' tiada!" };
  }
  
  var data = sheet.getDataRange().getValues();
  var userCari = usernameInput.trim().toLowerCase();
  var passCari = passwordInput.toString().trim();
  
  for (var i = 1; i < data.length; i++) {
    var namaSheet = data[i][0];
    var emailSheet = data[i][1];
    var userSheet = data[i][2].toString().trim().toLowerCase();
    var passSheet = data[i][3].toString().trim();
    
    if (userSheet === userCari && passSheet === passCari) {
      return { 
        sah: true, 
        nama: namaSheet,
        email: emailSheet
      };
    }
  }
  
  return { sah: false, mesej: "Kombinasi nama pengguna atau kata laluan adalah salah!" };
}