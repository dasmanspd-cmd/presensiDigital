/**
 * APLIKASI PRESENSI DIGITAL BERBASIS WEB (GOOGLE APPS SCRIPT)
 * Guru: DASMAN, S.Pd
 * Satuan Pendidikan: SMAN 1 X KOTO DIATAS
 */

const CONFIG = {
  APP_PASSWORD: "Garefa6957Ds#",
  TEACHER_NAME: "DASMAN, S.Pd",
  ROLE: "GURU MAPEL / WALI KELAS",
  SUBJECT: "EKONOMI",
  SCHOOL: "SMAN 1 X KOTO DIATAS",
  ACADEMIC_YEAR: "2026/2027",
  
  // DAFTAR KELAS
  CLASSES: [
    "Kelas X E1", 
    "Kelas X E2", 
    "Kelas X E3", 
    "Kelas XI P2", 
    "Kelas XI P3", 
    "Kelas XII P2", 
    "Kelas XII P3"
  ],
  
  TOTAL_MEETINGS: 45, // P1 s/d P45
  
  // DAFTAR SISWA DEFAULT
  STUDENTS: {
    "Kelas X E1": [
      { nis: "1001", name: "Ahmad Fauzi", gender: "L" },
      { nis: "1002", name: "Annisa Rahmawati", gender: "P" },
      { nis: "1003", name: "Budi Santoso", gender: "L" },
      { nis: "1004", name: "Citra Dewi", gender: "P" },
      { nis: "1005", name: "Dedi Kurniawan", gender: "L" }
    ],
    "Kelas X E2": [
      { nis: "1051", name: "Eka Nurpratiwi", gender: "P" },
      { nis: "1052", name: "Fajar Ramadan", gender: "L" },
      { nis: "1053", name: "Gita Savitri", gender: "P" },
      { nis: "1054", name: "Hendra Gunawan", gender: "L" },
      { nis: "1055", name: "Intan Permatasari", gender: "P" }
    ],
    "Kelas X E3": [
      { nis: "1101", name: "Joko Susilo", gender: "L" },
      { nis: "1102", name: "Kartika Sari", gender: "P" },
      { nis: "1103", name: "Lukman Hakim", gender: "L" },
      { nis: "1104", name: "Maya Anggraini", gender: "P" },
      { nis: "1105", name: "Naufal Rizky", gender: "L" }
    ],
    "Kelas XI P2": [
      { nis: "2001", name: "Aditya Pratama", gender: "L" },
      { nis: "2002", name: "Bella Safira", gender: "P" },
      { nis: "2003", name: "Candra Wijaya", gender: "L" },
      { nis: "2004", name: "Dina Mariana", gender: "P" },
      { nis: "2005", name: "Eko Prasetyo", gender: "L" }
    ],
    "Kelas XI P3": [
      { nis: "2051", name: "Fitri Handayani", gender: "P" },
      { nis: "2052", name: "Gilang Perdana", gender: "L" },
      { nis: "2053", name: "Hany Lestari", gender: "P" },
      { nis: "2054", name: "Irfan Hakim", gender: "L" },
      { nis: "2055", name: "Jasmine Putri", gender: "P" }
    ],
    "Kelas XII P2": [
      { nis: "3001", name: "Kurnia Sandi", gender: "L" },
      { nis: "3002", name: "Larasati Dewi", gender: "P" },
      { nis: "3003", name: "Muhammad Rizky", gender: "L" },
      { nis: "3004", name: "Nabila Syakieb", gender: "P" },
      { nis: "3005", name: "Oky Setiawan", gender: "L" }
    ],
    "Kelas XII P3": [
      { nis: "3051", name: "Putri Ayu", gender: "P" },
      { nis: "3052", name: "Qori Agung", gender: "L" },
      { nis: "3053", name: "Rina Nose", gender: "P" },
      { nis: "3054", name: "Surya Saputra", gender: "L" },
      { nis: "3055", name: "Tania Putri", gender: "P" }
    ]
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Presensi Digital - ' + CONFIG.SCHOOL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * OTOMATISASI STRUKTUR DATABASE GOOGLE SHEETS
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let rekapSheet = ss.getSheetByName('Rekap Presensi');
  if (!rekapSheet) {
    rekapSheet = ss.insertSheet('Rekap Presensi', 0);
  } else {
    rekapSheet.clear();
  }
  
  const rangeA1 = rekapSheet.getRange("A1");
  rangeA1.setValue("REKAPITULASI PRESENSI KESELURUHAN KELAS");
  rangeA1.setFontWeight("bold");
  rangeA1.setFontSize(14);
  rangeA1.setFontColor("#1F4E79");

  const rangeA2 = rekapSheet.getRange("A2");
  rangeA2.setValue(CONFIG.SCHOOL + " - TAHUN AJARAN " + CONFIG.ACADEMIC_YEAR);
  rangeA2.setFontStyle("italic");
  rangeA2.setFontSize(10);
  
  const rekapHeaders = [["NO", "NAMA KELAS", "TOTAL SISWA", "TOTAL HADIR (H)", "TOTAL SAKIT (S)", "TOTAL IZIN (I)", "TOTAL ALFA (A)", "TOTAL CABUT (C)", "RATA-RATA % KEHADIRAN"]];
  const rangeHeaders = rekapSheet.getRange(4, 1, 1, 9);
  rangeHeaders.setValues(rekapHeaders);
  rangeHeaders.setBackground("#1F4E79");
  rangeHeaders.setFontColor("#FFFFFF");
  rangeHeaders.setFontWeight("bold");
  rangeHeaders.setHorizontalAlignment("center");
  rangeHeaders.setVerticalAlignment("middle");
    
  rekapSheet.setRowHeight(4, 30);

  const requiredColumns = 4 + CONFIG.TOTAL_MEETINGS + 7;

  CONFIG.CLASSES.forEach((className, index) => {
    let sheet = ss.getSheetByName(className);
    if (!sheet) {
      sheet = ss.insertSheet(className);
    } else {
      sheet.clear();
    }
    
    const maxCols = sheet.getMaxColumns();
    if (maxCols < requiredColumns) {
      sheet.insertColumnsAfter(maxCols, requiredColumns - maxCols);
    }
    
    sheet.getRange("A1").setValue("NAMA GURU").setFontWeight("bold");
    sheet.getRange("B1").setValue(CONFIG.TEACHER_NAME);
    sheet.getRange("A2").setValue("JABATAN / MAPEL").setFontWeight("bold");
    sheet.getRange("B2").setValue(CONFIG.ROLE + " - " + CONFIG.SUBJECT);
    sheet.getRange("A3").setValue("SEKOLAH / TA").setFontWeight("bold");
    sheet.getRange("B3").setValue(CONFIG.SCHOOL + " / " + CONFIG.ACADEMIC_YEAR);
    
    const rangeLegend = sheet.getRange("E1");
    rangeLegend.setValue("KETERANGAN KODE: H = Hadir | S = Sakit | I = Izin | A = Alfa | C = Cabut");
    rangeLegend.setFontStyle("italic");
    rangeLegend.setFontSize(9);
    
    sheet.getRange("A7").setValue("NO");
    sheet.getRange("B7").setValue("NIS / NISN");
    sheet.getRange("C7").setValue("NAMA SISWA");
    sheet.getRange("D7").setValue("L/P");
    
    const rangeHeaderSiswa = sheet.getRange("A7:D8");
    rangeHeaderSiswa.merge();
    rangeHeaderSiswa.setBackground("#1F4E79");
    rangeHeaderSiswa.setFontColor("#FFFFFF");
    rangeHeaderSiswa.setFontWeight("bold");
    rangeHeaderSiswa.setHorizontalAlignment("center");
    rangeHeaderSiswa.setVerticalAlignment("middle");
    
    const meetingHeaders = [];
    for (let p = 1; p <= CONFIG.TOTAL_MEETINGS; p++) {
      meetingHeaders.push("P" + p);
    }
    
    const rangePTitle = sheet.getRange(7, 5, 1, CONFIG.TOTAL_MEETINGS);
    rangePTitle.merge();
    rangePTitle.setValue("PERTEMUAN / TANGGAL PRESENSI");
    rangePTitle.setBackground("#1F4E79");
    rangePTitle.setFontColor("#FFFFFF");
    rangePTitle.setFontWeight("bold");
    rangePTitle.setHorizontalAlignment("center");

    const rangePSub = sheet.getRange(8, 5, 1, CONFIG.TOTAL_MEETINGS);
    rangePSub.setValues([meetingHeaders]);
    rangePSub.setBackground("#D9E1F2");
    rangePSub.setFontColor("#1F4E79");
    rangePSub.setFontWeight("bold");
    rangePSub.setHorizontalAlignment("center");
    
    const summaryColStart = 5 + CONFIG.TOTAL_MEETINGS;
    const rangeRekapTitle = sheet.getRange(7, summaryColStart, 1, 7);
    rangeRekapTitle.merge();
    rangeRekapTitle.setValue("REKAPITULASI PRESENSI");
    rangeRekapTitle.setBackground("#1F4E79");
    rangeRekapTitle.setFontColor("#FFFFFF");
    rangeRekapTitle.setFontWeight("bold");
    rangeRekapTitle.setHorizontalAlignment("center");

    const rangeRekapSub = sheet.getRange(8, summaryColStart, 1, 7);
    rangeRekapSub.setValues([["H", "S", "I", "A", "C", "TOTAL", "% HADIR"]]);
    rangeRekapSub.setBackground("#D9E1F2");
    rangeRekapSub.setFontColor("#1F4E79");
    rangeRekapSub.setFontWeight("bold");
    rangeRekapSub.setHorizontalAlignment("center");
    
    const students = CONFIG.STUDENTS[className] || [];
    const studentRows = [];
    
    students.forEach((s, idx) => {
      studentRows.push([idx + 1, s.nis, s.name, s.gender]);
    });
    
    if (studentRows.length > 0) {
      sheet.getRange(9, 1, studentRows.length, 4).setValues(studentRows);
      
      for (let i = 0; i < studentRows.length; i++) {
        const row = 9 + i;
        applyStudentFormulas(sheet, row);
      }
      
      const dataRange = sheet.getRange(9, 1, studentRows.length, requiredColumns);
      dataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(9, 1, studentRows.length, 2).setHorizontalAlignment("center");
      sheet.getRange(9, 4, studentRows.length, requiredColumns - 3).setHorizontalAlignment("center");
    }
    
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 110);
    sheet.setColumnWidth(3, 230);
    sheet.setColumnWidth(4, 50);
    for (let c = 5; c <= 4 + CONFIG.TOTAL_MEETINGS; c++) {
      sheet.setColumnWidth(c, 40);
    }
    
    updateRekapSheetRow(rekapSheet, index, className, studentRows.length);
  });
  
  const rekapDataRange = rekapSheet.getRange(5, 1, CONFIG.CLASSES.length, 9);
  rekapDataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
  rekapSheet.getRange(5, 1, CONFIG.CLASSES.length, 1).setHorizontalAlignment("center");
  rekapSheet.getRange(5, 3, CONFIG.CLASSES.length, 7).setHorizontalAlignment("center");
  
  for (let col = 1; col <= 9; col++) {
    rekapSheet.autoResizeColumn(col);
  }
  
  return "Database Berhasil Dikonfigurasi!";
}

function applyStudentFormulas(sheet, row) {
  const summaryColStart = 5 + CONFIG.TOTAL_MEETINGS;
  const lastPCol = columnToLetter(4 + CONFIG.TOTAL_MEETINGS);
  
  sheet.getRange(row, summaryColStart).setFormula(`=COUNTIF(E${row}:${lastPCol}${row}, "H")`);
  sheet.getRange(row, summaryColStart + 1).setFormula(`=COUNTIF(E${row}:${lastPCol}${row}, "S")`);
  sheet.getRange(row, summaryColStart + 2).setFormula(`=COUNTIF(E${row}:${lastPCol}${row}, "I")`);
  sheet.getRange(row, summaryColStart + 3).setFormula(`=COUNTIF(E${row}:${lastPCol}${row}, "A")`);
  sheet.getRange(row, summaryColStart + 4).setFormula(`=COUNTIF(E${row}:${lastPCol}${row}, "C")`);
  sheet.getRange(row, summaryColStart + 5).setFormula(`=COUNTA(E${row}:${lastPCol}${row})`);
  
  const totCell = columnToLetter(summaryColStart + 5) + row;
  const hCell = columnToLetter(summaryColStart) + row;
  const pctRange = sheet.getRange(row, summaryColStart + 6);
  pctRange.setFormula(`=IF(${totCell}>0, ${hCell}/${totCell}, 0)`);
  pctRange.setNumberFormat("0.0%");
}

function updateRekapSheetRow(rekapSheet, classIndex, className, totalStudents) {
  const summaryColStart = 5 + CONFIG.TOTAL_MEETINGS;
  const rekapRow = 4 + classIndex + 1;
  const endRow = 8 + (totalStudents > 0 ? totalStudents : 1);
  const colH = columnToLetter(summaryColStart);
  const colS = columnToLetter(summaryColStart + 1);
  const colI = columnToLetter(summaryColStart + 2);
  const colA = columnToLetter(summaryColStart + 3);
  const colC = columnToLetter(summaryColStart + 4);
  const colPct = columnToLetter(summaryColStart + 6);
  
  rekapSheet.getRange(rekapRow, 1, 1, 9).setValues([[
    classIndex + 1,
    className,
    `=COUNTA('${className}'!C9:C${endRow})`,
    `=SUM('${className}'!${colH}9:${colH}${endRow})`,
    `=SUM('${className}'!${colS}9:${colS}${endRow})`,
    `=SUM('${className}'!${colI}9:${colI}${endRow})`,
    `=SUM('${className}'!${colA}9:${colA}${endRow})`,
    `=SUM('${className}'!${colC}9:${colC}${endRow})`,
    `=AVERAGE('${className}'!${colPct}9:${colPct}${endRow})`
  ]]);
  rekapSheet.getRange(rekapRow, 9).setNumberFormat("0.0%");
}

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * SERVIS / INTERAKSI CLIENT WEB
 */
function verifyPassword(pwd) {
  if (pwd === CONFIG.APP_PASSWORD) {
    return {
      success: true,
      teacher: CONFIG.TEACHER_NAME,
      school: CONFIG.SCHOOL,
      classes: CONFIG.CLASSES,
      totalMeetings: CONFIG.TOTAL_MEETINGS
    };
  }
  return { success: false, message: "Password salah! Silakan coba lagi." };
}

function getClassData(className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(className);
  
  if (!sheet) {
    throw new Error("Sheet kelas '" + className + "' tidak ditemukan. Jalankan setupDatabase() terlebih dahulu.");
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 9) return [];
  
  const summaryColStart = 5 + CONFIG.TOTAL_MEETINGS;
  
  const studentInfo = sheet.getRange(9, 1, lastRow - 8, 4).getValues();
  const summaryInfo = sheet.getRange(9, summaryColStart + 3, lastRow - 8, 4).getValues();
  
  const students = [];
  for (let i = 0; i < studentInfo.length; i++) {
    if (studentInfo[i][2]) {
      students.push({
        row: 9 + i,
        no: studentInfo[i][0],
        nis: studentInfo[i][1],
        name: studentInfo[i][2],
        gender: studentInfo[i][3],
        alfa: summaryInfo[i][0] || 0,
        cabut: summaryInfo[i][1] || 0,
        pct: summaryInfo[i][3] ? (summaryInfo[i][3] * 100).toFixed(1) + "%" : "0.0%"
      });
    }
  }
  
  return students;
}

function addStudent(payload) {
  const { className, nis, name, gender } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(className);
  
  if (!sheet) throw new Error("Sheet kelas tidak ditemukan.");
  
  const lastRow = sheet.getLastRow();
  const newRow = lastRow < 9 ? 9 : lastRow + 1;
  const newNo = newRow - 8;
  
  sheet.getRange(newRow, 1, 1, 4).setValues([[newNo, nis, name, gender]]);
  applyStudentFormulas(sheet, newRow);
  
  const requiredColumns = 4 + CONFIG.TOTAL_MEETINGS + 7;
  const rowRange = sheet.getRange(newRow, 1, 1, requiredColumns);
  rowRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(newRow, 1, 1, 2).setHorizontalAlignment("center");
  sheet.getRange(newRow, 4, 1, requiredColumns - 3).setHorizontalAlignment("center");
  
  const classIndex = CONFIG.CLASSES.indexOf(className);
  const rekapSheet = ss.getSheetByName('Rekap Presensi');
  if (rekapSheet && classIndex !== -1) {
    updateRekapSheetRow(rekapSheet, classIndex, className, newNo);
  }
  
  return { success: true, message: `Siswa "${name}" berhasil ditambahkan ke ${className}!` };
}

function importStudentsBulk(payload) {
  const { className, studentList } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(className);
  
  if (!sheet) throw new Error("Sheet kelas tidak ditemukan.");
  if (!studentList || studentList.length === 0) throw new Error("Data siswa kosong.");
  
  const currentLastRow = sheet.getLastRow();
  const startRow = currentLastRow < 9 ? 9 : currentLastRow + 1;
  
  const studentRows = [];
  studentList.forEach((s, idx) => {
    const no = (startRow - 8) + idx;
    const nis = String(s.nis || '').trim();
    const name = String(s.name || '').trim();
    let gender = String(s.gender || 'L').trim().toUpperCase();
    if (gender.startsWith('P')) gender = 'P';
    else gender = 'L';
    
    if (name) {
      studentRows.push([no, nis, name, gender]);
    }
  });
  
  if (studentRows.length === 0) throw new Error("Tidak ada data siswa valid untuk diimpor.");
  
  sheet.getRange(startRow, 1, studentRows.length, 4).setValues(studentRows);
  
  for (let i = 0; i < studentRows.length; i++) {
    const row = startRow + i;
    applyStudentFormulas(sheet, row);
  }
  
  const requiredColumns = 4 + CONFIG.TOTAL_MEETINGS + 7;
  const dataRange = sheet.getRange(startRow, 1, studentRows.length, requiredColumns);
  dataRange.setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(startRow, 1, studentRows.length, 2).setHorizontalAlignment("center");
  sheet.getRange(startRow, 4, studentRows.length, requiredColumns - 3).setHorizontalAlignment("center");
  
  const totalStudents = (startRow - 8) + studentRows.length;
  const classIndex = CONFIG.CLASSES.indexOf(className);
  const rekapSheet = ss.getSheetByName('Rekap Presensi');
  if (rekapSheet && classIndex !== -1) {
    updateRekapSheetRow(rekapSheet, classIndex, className, totalStudents);
  }
  
  return { 
    success: true, 
    message: `Berhasil mengimpor ${studentRows.length} siswa ke ${className}!` 
  };
}

function deleteStudent(payload) {
  const { className, row, name } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(className);
  
  if (!sheet) throw new Error("Sheet kelas tidak ditemukan.");
  
  sheet.deleteRow(row);
  
  const newLastRow = sheet.getLastRow();
  const totalStudents = newLastRow >= 9 ? newLastRow - 8 : 0;
  
  if (totalStudents > 0) {
    for (let r = 9; r <= newLastRow; r++) {
      sheet.getRange(r, 1).setValue(r - 8);
      applyStudentFormulas(sheet, r);
    }
  }
  
  const classIndex = CONFIG.CLASSES.indexOf(className);
  const rekapSheet = ss.getSheetByName('Rekap Presensi');
  if (rekapSheet && classIndex !== -1) {
    updateRekapSheetRow(rekapSheet, classIndex, className, totalStudents);
  }
  
  return { success: true, message: `Data siswa "${name}" berhasil dihapus dari ${className}.` };
}

function saveAttendance(payload) {
  const { className, meetingIndex, dateStr, attendanceData } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(className);
  
  if (!sheet) throw new Error("Sheet kelas tidak ditemukan.");
  
  const targetCol = 4 + parseInt(meetingIndex);
  
  sheet.getRange(8, targetCol).setValue("P" + meetingIndex + "\n(" + dateStr + ")");
  
  attendanceData.forEach(item => {
    sheet.getRange(item.row, targetCol).setValue(item.status);
  });
  
  return { success: true, message: `Presensi Pertemuan ${meetingIndex} (${className}) Berhasil Disimpan!` };
}