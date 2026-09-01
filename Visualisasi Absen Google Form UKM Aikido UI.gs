// maap klo berantakan, masih pemula ~baron 7 Agustus 2026
// sistemnya, dia bakal rewrite/replace semua data pada bulan ini. bulan sebelum-sebelumnya ga disentuh.
// usahakan tabs "Form Responses 1" steril terutama nama dan tanggal, kecuali kebutuhan update sistem 
function buatSheetBulanIniManual() {
  var today = new Date();
  generateSheetAbsensi(today.getMonth(), today.getFullYear());
  buatSheetRangkumanTahun();
}

function generateSheetAbsensi(targetBulan, targetTahun) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName("JAWABAN FORM"); // nama tabs response gform
  
  if (!formSheet) {
    Browser.msgBox("Sheet 'JAWABAN FORM' tidak ditemukan! cek nama tabnya.");
    return;
  }

  var namaBulanIndo = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];

  // pemetaan inisial hari Eng
  var dayInitials = {
    2: "Tu",  // Tuesday
    4: "Th",  // Thursday
    6: "Sat"  // Saturday
  };

  var namaSheetBaru = namaBulanIndo[targetBulan] + " " + targetTahun;

  // --- 1. GENERATE DATA TANGGAL & HEADER BARIS 1 & 2 ---
  var datesFormattedCompare = []; 
  var headerDays = ["Nama Responden"];
  var headerDates = [""]; 
  var datePointer = new Date(targetTahun, targetBulan, 1);
  var dayTypes = []; 

  while (datePointer.getMonth() === targetBulan) {
    var day = datePointer.getDay();
    // 2 = Selasa, 4 = Kamis, 6 = Sabtu
    if (day === 2 || day === 4 || day === 6) {
      var fullFormattedDate = Utilities.formatDate(datePointer, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      datesFormattedCompare.push(fullFormattedDate);

      headerDays.push(dayInitials[day]);
      var dateNumber = Utilities.formatDate(datePointer, ss.getSpreadsheetTimeZone(), "dd");
      headerDates.push(dateNumber);
      dayTypes.push(day);
    }
    datePointer.setDate(datePointer.getDate() + 1);
  }

  // --- 2. BACA & PROSES DATA DARI FORM RESPONSES ---
  var data = formSheet.getDataRange().getValues();
  var frequencyMap = {};
  var attendanceRecords = {};
  var dataDitemukan = 0;

  for (var i = 1; i < data.length; i++) {
    var rawTimestamp = data[i][0]; // Kolom A: Timestamp
    
    // Evaluasi nama langsung di memori: Kolom D (indeks 3) atau Kolom E (indeks 4)
    var namaD = data[i][3] ? data[i][3].toString().trim() : "";
    var namaE = data[i][4] ? data[i][4].toString().trim() : "";
    var nama = namaD || namaE; // Mengambil nama yang terisi

    if (!rawTimestamp || !nama) continue;

    var timestamp = new Date(rawTimestamp);

    if (!isNaN(timestamp.getTime())) {
      if (timestamp.getMonth() === targetBulan && timestamp.getFullYear() === targetTahun) {
        dataDitemukan++;
        
        frequencyMap[nama] = (frequencyMap[nama] || 0) + 1;
        
        if (!attendanceRecords[nama]) {
          attendanceRecords[nama] = [];
        }
        
        var tglRespon = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
        attendanceRecords[nama].push(tglRespon);
      }
    }
  }

  // JIKA TIDAK ADA DATA DITEMUKAN
  if (dataDitemukan === 0) {
    Browser.msgBox("Tidak ada data respon yang cocok untuk bulan " + namaBulanIndo[targetBulan] + " " + targetTahun + ".");
    return;
  }

  // --- 3. BUAT/REPLACE TAB BARU ---
  var sheetBaru = ss.getSheetByName(namaSheetBaru);
  if (sheetBaru) {
    ss.deleteSheet(sheetBaru);
  }
  sheetBaru = ss.insertSheet(namaSheetBaru);

  // Susun Header
  headerDays.push("Total Kehadiran");
  headerDates.push("");

  var totalCols = headerDays.length;

  // Tulis Baris 1 (Inisial Hari) & Baris 2 (Tanggal)
  var rangeHeaderBaris1 = sheetBaru.getRange(1, 1, 1, totalCols);
  var rangeHeaderBaris2 = sheetBaru.getRange(2, 1, 1, totalCols);

  rangeHeaderBaris1.setValues([headerDays]).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center").setVerticalAlignment("middle");
  rangeHeaderBaris2.setValues([headerDates]).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center").setVerticalAlignment("middle");

  // --- MERGE KOLOM A1:A2 & TOTAL KEHADIRAN ---
  var rangeNamaResponden = sheetBaru.getRange("A1:A2");
  rangeNamaResponden.merge().setValue("Nama Responden").setHorizontalAlignment("left").setVerticalAlignment("middle");

  var rangeTotalKehadiran = sheetBaru.getRange(1, totalCols, 2, 1);
  rangeTotalKehadiran.merge().setValue("Total Kehadiran").setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Urutkan nama berdasarkan kehadiran terbanyak
  var sortedNama = Object.keys(frequencyMap).sort(function(a, b) {
    return frequencyMap[b] - frequencyMap[a];
  });

  // --- 4. ISI MATRIX KEHADIRAN (✓), TOTAL, DAN BREAKDOWN HARI ---
  var matrixOutput = [];
  var breakdownHariMap = {}; 

  for (var r = 0; r < sortedNama.length; r++) {
    var namaSiswa = sortedNama[r];
    var row = [namaSiswa];
    var userDates = attendanceRecords[namaSiswa] || [];
    var totalHadir = 0;

    var tuCount = 0, thCount = 0, satCount = 0;

    for (var c = 0; c < datesFormattedCompare.length; c++) {
      var targetTgl = datesFormattedCompare[c];
      var dType = dayTypes[c];

      if (userDates.indexOf(targetTgl) !== -1) {
        row.push("✓");
        totalHadir++;
        if (dType === 2) tuCount++;
        if (dType === 4) thCount++;
        if (dType === 6) satCount++;
      } else {
        row.push("");
      }
    }
    
    row.push(totalHadir);
    matrixOutput.push(row);

    breakdownHariMap[namaSiswa] = {
      "Selasa": tuCount,
      "Kamis": thCount,
      "Sabtu": satCount,
      "Total": totalHadir
    };
  }

  // Masukkan Data Utama (Mulai Baris 3)
  var startRow = 3;
  var totalRows = matrixOutput.length;

  var dataRange = sheetBaru.getRange(startRow, 1, totalRows, totalCols);
  dataRange.setValues(matrixOutput);

  // --- 5. FORMATTING, TEXT WRAPPING, & DESAIN TABEL ---
  // A. Alignment & Vertical Center
  sheetBaru.getRange(startRow, 2, totalRows, totalCols - 1).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheetBaru.getRange(startRow, 1, totalRows, 1).setHorizontalAlignment("left").setVerticalAlignment("middle");

  // B. Text Wrapping ke Wrap untuk seluruh tabel
  var fullHeaderAndDataRange = sheetBaru.getRange(1, 1, totalRows + 2, totalCols);
  fullHeaderAndDataRange.setWrap(true);

  // C. Warna Fill Kolom "Total Kehadiran" (#a4c2f4)
  sheetBaru.getRange(1, totalCols, totalRows + 2, 1).setBackground("#a4c2f4");

  // D. Border untuk Tabel Utama
  fullHeaderAndDataRange.setBorder(true, true, true, true, true, true, "#b7b7b7", SpreadsheetApp.BorderStyle.SOLID);

  // E. Bekukan Kolom A & Baris 2 (Freeze Column A & Row 2)
  sheetBaru.setFrozenColumns(1);
  sheetBaru.setFrozenRows(2);

  // F. Pengaturan Lebar Kolom (Fix 175 untuk Kolom A)
  sheetBaru.setColumnWidth(1, 175); // FIX 175 PX
  for (var k = 2; k < totalCols; k++) {
    sheetBaru.setColumnWidth(k, 42); // Lebar ringkas untuk tanggal ("Tu", "Th", "Sat")
  }
  sheetBaru.setColumnWidth(totalCols, 90); // Lebar pas untuk "Total Kehadiran"

  // --- 6. BUAT TABEL RANGKUMAN (HELPER) & GRAFIK BATANG ---
  buatTabelRangkumanDanGrafik(sheetBaru, sortedNama, breakdownHariMap, totalCols);

  Browser.msgBox("Gacor king yeyy bisa gess letsgoo.");
}

// FUNGSI BUAT TABEL RANGKUMAN & GRAFIK BATANG BERKELOMPOK
function buatTabelRangkumanDanGrafik(sheet, sortedNama, breakdownHariMap, totalCols) {
  var helperStartCol = totalCols + 2; // Berjarak 1 kolom setelah tabel utama
  var helperColsCount = 5; // Nama, Selasa, Kamis, Sabtu, Total Kehadiran
  
  // A. Header Tabel Rangkuman (Baris 1 & 2 di-merge)
  var headersRangkuman = ["Nama", "Selasa", "Kamis", "Sabtu", "Total Kehadiran"];
  
  for (var h = 0; h < helperColsCount; h++) {
    var colPos = helperStartCol + h;
    var rangeHeader = sheet.getRange(1, colPos, 2, 1);
    rangeHeader.merge()
               .setValue(headersRangkuman[h])
               .setFontWeight("bold")
               .setBackground("#d9ead3")
               .setHorizontalAlignment("center")
               .setVerticalAlignment("middle")
               .setWrap(true);
  }

  // B. Data Rangkuman Responden (Mulai dari Baris ke-3)
  var helperData = [];
  for (var i = 0; i < sortedNama.length; i++) {
    var n = sortedNama[i];
    helperData.push([
      n,
      breakdownHariMap[n]["Selasa"],
      breakdownHariMap[n]["Kamis"],
      breakdownHariMap[n]["Sabtu"],
      breakdownHariMap[n]["Total"]
    ]);
  }

  var startRowData = 3;
  var helperDataRange = sheet.getRange(startRowData, helperStartCol, helperData.length, helperColsCount);
  helperDataRange.setValues(helperData);

  // Formatting Tabel Rangkuman
  sheet.getRange(startRowData, helperStartCol, helperData.length, 1).setHorizontalAlignment("left").setVerticalAlignment("middle");
  sheet.getRange(startRowData, helperStartCol + 1, helperData.length, 4).setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  var fullHelperRange = sheet.getRange(1, helperStartCol, helperData.length + 2, helperColsCount);
  fullHelperRange.setBorder(true, true, true, true, true, true, "#b7b7b7", SpreadsheetApp.BorderStyle.SOLID);
  
  // Warna background untuk kolom "Total Kehadiran" di tabel rangkuman
  sheet.getRange(1, helperStartCol + 4, helperData.length + 2, 1).setBackground("#a4c2f4");

  // Rapikan lebar kolom tabel rangkuman
  sheet.setColumnWidth(helperStartCol, 150); // Lebar kolom Nama Rangkuman
  sheet.setColumnWidth(helperStartCol + 1, 65); // Selasa
  sheet.setColumnWidth(helperStartCol + 2, 65); // Kamis
  sheet.setColumnWidth(helperStartCol + 3, 65); // Sabtu
  sheet.setColumnWidth(helperStartCol + 4, 90); // Total Kehadiran

  // C. Range Acuan untuk Grafik (Hanya Kolom Nama, Selasa, Kamis, Sabtu - tanpa Total Kehadiran)
  var rangeGrafikNama = sheet.getRange(1, helperStartCol, helperData.length + 2, 1);
  var rangeGrafikHari = sheet.getRange(1, helperStartCol + 1, helperData.length + 2, 3);

  // D. Buat Grafik Batang Vertikal Berkelompok (Grouped Column Chart)
  var chartRowPosition = helperData.length + 5;

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(rangeGrafikNama)
    .addRange(rangeGrafikHari)
    .setPosition(chartRowPosition, helperStartCol, 0, 0)
    .setOption('title', 'Rincian Kehadiran Per Hari (Selasa, Kamis, Sabtu)')
    .setOption('hAxis', {title: 'Nama Responden'})
    .setOption('vAxis', {
      title: 'Jumlah Kehadiran',
      minValue: 0,
      format: '0',
      viewWindow: {min: 0}
    })
    .setOption('legend', {position: 'top'})
    .setOption('width', 750)
    .setOption('height', 420)
    .build();

  sheet.insertChart(chart);
}

// update 1 september
/**
 * Fungsi untuk membuat/mereplace tab "Rangkuman Kehadiran"
 * berdasarkan data dari tab-tab bulanan yang sudah dibuat.
 */
function buatSheetRangkumanTahun() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();

  var monthNames = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];

  // 1. CARI TAB BULANAN YANG SESUAI FORMAT ("Bulan Tahun")
  var monthlySheets = [];
  
  for (var i = 0; i < allSheets.length; i++) {
    var sheetName = allSheets[i].getName();
    var parts = sheetName.split(" ");
    
    // Cek apakah nama sheet terdiri dari 2 kata (Bulan dan Tahun)
    if (parts.length === 2) {
      var monthIdx = monthNames.indexOf(parts[0]);
      var yearNum = parseInt(parts[1], 10);

      if (monthIdx !== -1 && !isNaN(yearNum)) {
        monthlySheets.push({
          sheet: allSheets[i],
          name: sheetName,
          monthName: parts[0],
          monthIdx: monthIdx,
          year: yearNum
        });
      }
    }
  }

  if (monthlySheets.length === 0) {
    Browser.msgBox("Tidak ditemukan tab bulanan (contoh: 'Agustus 2026') untuk dibuatkan rangkuman.");
    return;
  }

  // Urutkan tab berdasarkan Tahun lalu Bulan secara kronologis
  monthlySheets.sort(function(a, b) {
    if (a.year !== b.year) return a.year - b.year;
    return a.monthIdx - b.monthIdx;
  });

  // 2. EKSTRAKSI DATA KEHADIRAN & DAFTAR RESPONDEN
  var respondentsSet = {};
  var attendanceMatrix = {}; // Structure: { nama: { sheetName: totalKehadiran } }

  for (var m = 0; m < monthlySheets.length; m++) {
    var mInfo = monthlySheets[m];
    var mSheet = mInfo.sheet;
    var data = mSheet.getDataRange().getValues();

    if (data.length < 3) continue; // Skip jika tidak ada data responden

    // Kolom Total Kehadiran di tab bulanan selalu berada sebelum tabel helper rangkuman
    // Kita cari indeks kolom "Total Kehadiran" dari baris 1 atau 2
    var totalColIdx = -1;
    for (var c = 0; c < data[0].length; c++) {
      var cellValHeader1 = data[0][c].toString().trim();
      var cellValHeader2 = data[1][c].toString().trim();
      if (cellValHeader1 === "Total Kehadiran" || cellValHeader2 === "Total Kehadiran") {
        totalColIdx = c;
        break;
      }
    }

    if (totalColIdx === -1) continue;

    // Baca data responden mulai dari baris ke-3
    for (var r = 2; r < data.length; r++) {
      var nama = data[r][0] ? data[r][0].toString().trim() : "";
      var totalHadir = data[r][totalColIdx];

      if (nama !== "") {
        respondentsSet[nama] = true;
        if (!attendanceMatrix[nama]) {
          attendanceMatrix[nama] = {};
        }
        attendanceMatrix[nama][mInfo.name] = totalHadir !== "" ? totalHadir : 0;
      }
    }
  }

  var sortedRespondents = Object.keys(respondentsSet).sort();

  if (sortedRespondents.length === 0) {
    Browser.msgBox("Tidak ada data responden yang valid pada tab bulanan.");
    return;
  }

  // 3. BUAT / REPLACE TAB "Rangkuman Kehadiran"
  var targetSheetName = "Rangkuman Kehadiran";
  var rangkumanSheet = ss.getSheetByName(targetSheetName);
  if (rangkumanSheet) {
    ss.deleteSheet(rangkumanSheet);
  }
  rangkumanSheet = ss.insertSheet(targetSheetName);

  // 4. SUSUN HEADER BARIS 1 & BARIS 2
  var headerRow1 = ["Nama Responden"];
  var headerRow2 = [""];

  for (var k = 0; k < monthlySheets.length; k++) {
    headerRow1.push(monthlySheets[k].year.toString());
    headerRow2.push(monthlySheets[k].monthName);
  }

  var totalCols = headerRow1.length;

  // Tulis Header
  rangkumanSheet.getRange(1, 1, 1, totalCols).setValues([headerRow1]);
  rangkumanSheet.getRange(2, 1, 1, totalCols).setValues([headerRow2]);

  // Merge Kolom A1:A2 (Nama Responden)
  var rangeNama = rangkumanSheet.getRange("A1:A2");
  rangeNama.merge()
           .setValue("Nama Responden")
           .setFontWeight("bold")
           .setBackground("#d9ead3")
           .setHorizontalAlignment("left")
           .setVerticalAlignment("middle");

  // Format Header Bulan & Tahun (Kolom B dan seterusnya)
  var rangeHeaderMonths = rangkumanSheet.getRange(2, 2, 1, totalCols - 1);
  rangeHeaderMonths.setFontWeight("bold")
                   .setBackground("#d9ead3")
                   .setHorizontalAlignment("center")
                   .setVerticalAlignment("middle");

  // Merge Baris 1 untuk Tahun yang Sama
  var startCol = 2;
  while (startCol <= totalCols) {
    var currentYear = rangkumanSheet.getRange(1, startCol).getValue();
    var endCol = startCol;

    while (endCol + 1 <= totalCols && rangkumanSheet.getRange(1, endCol + 1).getValue() == currentYear) {
      endCol++;
    }

    var rangeYear = rangkumanSheet.getRange(1, startCol, 1, endCol - startCol + 1);
    if (endCol > startCol) {
      rangeYear.merge();
    }
    rangeYear.setFontWeight("bold")
             .setBackground("#a4c2f4")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");

    startCol = endCol + 1;
  }

  // 5. ISI DATA UTAMA (Mulai Baris 3)
  var matrixOutput = [];
  for (var i = 0; i < sortedRespondents.length; i++) {
    var resNama = sortedRespondents[i];
    var row = [resNama];

    for (var j = 0; j < monthlySheets.length; j++) {
      var sName = monthlySheets[j].name;
      var count = attendanceMatrix[resNama][sName];
      row.push(count !== undefined ? count : 0);
    }
    matrixOutput.push(row);
  }

  var startRow = 3;
  var totalRows = matrixOutput.length;
  var dataRange = rangkumanSheet.getRange(startRow, 1, totalRows, totalCols);
  dataRange.setValues(matrixOutput);

  // 6. FORMATTING & LAYOUT TABEL
  // Format rata tengah untuk angka kehadiran
  rangkumanSheet.getRange(startRow, 2, totalRows, totalCols - 1)
                .setHorizontalAlignment("center")
                .setVerticalAlignment("middle");

  // Format rata kiri untuk nama responden
  rangkumanSheet.getRange(startRow, 1, totalRows, 1)
                .setHorizontalAlignment("left")
                .setVerticalAlignment("middle");

  // Border seluruh tabel
  var fullTableRange = rangkumanSheet.getRange(1, 1, totalRows + 2, totalCols);
  fullTableRange.setBorder(true, true, true, true, true, true, "#b7b7b7", SpreadsheetApp.BorderStyle.SOLID);
  fullTableRange.setWrap(true);

  // Freeze Baris 1-2 & Kolom A
  rangkumanSheet.setFrozenRows(2);
  rangkumanSheet.setFrozenColumns(1);

  // Lebar Kolom
  rangkumanSheet.setColumnWidth(1, 175); // Kolom Nama Responden
  for (var cIdx = 2; cIdx <= totalCols; cIdx++) {
    rangkumanSheet.setColumnWidth(cIdx, 100); // Lebar Kolom Bulan
  }

  Browser.msgBox("Tab 'Rangkuman Kehadiran' berhasil diperbarui!");
}
