const SPREADSHEET_ID = '1tFd7rHbOPbq9iCssAXNKPzUuk9CTebtNrq8FVZ0fPTM';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Accounting Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// 1. DATA DASHBOARD
function getDashboardSummary() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  function parseNominal(val) {
    if (!val) return 0;
    if (typeof val === 'string') val = val.replace(/Rp/gi, '').replace(/\./g, '').replace(/,/g, '').trim();
    let num = parseFloat(val);
    return isNaN(num) ? 0 : num;
  }

  const now = new Date(new Date().toLocaleString("en-US", {timeZone: "Asia/Jakarta"}));
  const currentY = now.getFullYear();
  const currentM = now.getMonth(); 

  function getMonthYear(row) {
    let dateVal = row[2]; 
    let bulanVal = row[3]; 
    
    let mm = -1;
    let yyyy = currentY;
    
    if (Object.prototype.toString.call(dateVal) === '[object Date]') {
       yyyy = dateVal.getFullYear();
    } else if (typeof dateVal === 'string') {
       let match = dateVal.match(/\d{4}/);
       if(match) yyyy = parseInt(match[0]);
    }
    
    if (bulanVal && typeof bulanVal === 'string') {
       const mEn = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
       let idx = mEn.indexOf(bulanVal.trim());
       if (idx !== -1) mm = idx;
    }
    
    if (mm === -1) {
       if (Object.prototype.toString.call(dateVal) === '[object Date]') mm = dateVal.getMonth();
       else if (typeof dateVal === 'string') {
          let parts = dateVal.split('/');
          if(parts.length === 3) mm = parseInt(parts[1]) - 1; 
       }
    }
    
    if (mm === -1 || isNaN(yyyy)) return null;
    return { mm: mm, yyyy: yyyy };
  }

  const debitData = ss.getSheetByName('Debit') ? ss.getSheetByName('Debit').getDataRange().getValues() : [];
  const kreditData = ss.getSheetByName('Kredit') ? ss.getSheetByName('Kredit').getDataRange().getValues() : [];
  const pinjamanData = ss.getSheetByName('Pinjaman') ? ss.getSheetByName('Pinjaman').getDataRange().getValues() : [];
  
  let debitTotal = 0, pinjamanTotal = 0, kreditOperasional = 0, kreditBayarHutang = 0;
  let currentMonthIncome = 0, currentMonthExpense = 0; 
  let topPicsMap = {}, topPicsMonthlyMap = {}, hutangMap = {}; 
  
  let monthlyMap = {}; 
  let monthNames = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Ags", "Sep", "Okt", "Nov", "Des"];
  for(let m=0; m<12; m++) { monthlyMap[m] = { label: monthNames[m], income: 0, expense: 0, ts: m }; }
  let yearlyMap = {};  

  // --- PROSES DEBIT ---
  for(let i=1; i<debitData.length; i++) {
    if(debitData[i][0] === "") continue;
    let nom = parseNominal(debitData[i][6]);
    debitTotal += nom;
    
    let mData = getMonthYear(debitData[i]);
    if(mData) {
      if(mData.yyyy === currentY) monthlyMap[mData.mm].income += nom;
      
      let yKey = mData.yyyy.toString();
      if(!yearlyMap[yKey]) yearlyMap[yKey] = { label: yKey, ts: mData.yyyy, income: 0, expense: 0 };
      yearlyMap[yKey].income += nom;

      if(mData.mm === currentM && mData.yyyy === currentY) currentMonthIncome += nom;
    }

    let pic = debitData[i][7] ? String(debitData[i][7]).trim() : '-';
    if(pic !== "" && pic !== "-") {
      topPicsMap[pic] = (topPicsMap[pic] || 0) + nom; 
      if(mData && mData.mm === currentM && mData.yyyy === currentY) {
        topPicsMonthlyMap[pic] = (topPicsMonthlyMap[pic] || 0) + nom;
      }
    }
  }
  
  // --- PROSES KREDIT ---
  for(let i=1; i<kreditData.length; i++) { 
    if(kreditData[i][0] === "") continue;
    let nom = parseNominal(kreditData[i][6]);
    let pin = parseInt(kreditData[i][4]);
    let pic = kreditData[i][7] ? String(kreditData[i][7]).trim() : '-';
    let mData = getMonthYear(kreditData[i]);

    if(pin === 10) { 
      kreditBayarHutang += nom;
      if(pic !== "" && pic !== "-") {
        let pKey = pic.toUpperCase(); 
        if(!hutangMap[pKey]) hutangMap[pKey] = { name: pic, pinjaman: 0, bayar: 0 };
        hutangMap[pKey].bayar += nom;
      }
    } else { 
      kreditOperasional += nom; 
      if(mData) {
        if(mData.yyyy === currentY) monthlyMap[mData.mm].expense += nom;
        
        let yKey = mData.yyyy.toString();
        if(!yearlyMap[yKey]) yearlyMap[yKey] = { label: yKey, ts: mData.yyyy, income: 0, expense: 0 };
        yearlyMap[yKey].expense += nom;
      }
    }

    if(mData && mData.mm === currentM && mData.yyyy === currentY) {
      currentMonthExpense += nom;
    }
  }

  // --- PROSES PINJAMAN ---
  for(let i=1; i<pinjamanData.length; i++) { 
    if(pinjamanData[i][0] === "") continue;
    let nom = parseNominal(pinjamanData[i][6]);
    pinjamanTotal += nom;

    let pic = pinjamanData[i][7] ? String(pinjamanData[i][7]).trim() : '-';
    if(pic !== "" && pic !== "-") {
      let pKey = pic.toUpperCase();
      if(!hutangMap[pKey]) hutangMap[pKey] = { name: pic, pinjaman: 0, bayar: 0 };
      hutangMap[pKey].pinjaman += nom;
    }
  }

  // --- PERBAIKAN LOGIKA SALDO KAS ---
  // Murni Cash Flow: Hanya Uang yang riil masuk (Debit) dikurangi Uang Keluar (Kredit)
  let totalKreditAllTime = kreditOperasional + kreditBayarHutang;
  let saldoKasSaatIni = debitTotal - totalKreditAllTime;

  let topPics = Object.keys(topPicsMap).map(k => ({pic: k, total: topPicsMap[k]})).sort((a, b) => b.total - a.total);
  let topPicsMonthly = Object.keys(topPicsMonthlyMap).map(k => ({pic: k, total: topPicsMonthlyMap[k]})).sort((a, b) => b.total - a.total);

  let debtTracker = [];
  for(let k in hutangMap) {
    let d = hutangMap[k];
    let sisa = d.pinjaman - d.bayar;
    if(d.pinjaman > 0 || d.bayar > 0) debtTracker.push({ pic: d.name, pinjaman: d.pinjaman, bayar: d.bayar, sisa: sisa });
  }
  debtTracker.sort((a,b) => b.sisa - a.sisa); 

  let monthlyChart = Object.values(monthlyMap).sort((a, b) => a.ts - b.ts);
  let yearlyChart = Object.values(yearlyMap).sort((a, b) => a.ts - b.ts);
  
  return {
    saldoKas: saldoKasSaatIni,
    currentMonthIncome: currentMonthIncome,
    currentMonthExpense: currentMonthExpense,
    labaKotor: debitTotal,
    totalPengeluaran: totalKreditAllTime, // Di-passing ke Frontend
    labaBersih: debitTotal - kreditOperasional,
    sisaHutang: pinjamanTotal - kreditBayarHutang,
    topPics: topPics.slice(0, 10),
    topPicsMonthly: topPicsMonthly.slice(0, 10),
    debtTracker: debtTracker,
    monthlyChart: monthlyChart,
    yearlyChart: yearlyChart,
    currentYear: currentY
  };
}

// 2. MENGAMBIL DATA LOGS
function getRecentLogs() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName('Logs');
  if(!logSheet) return [];
  
  const lastRow = logSheet.getLastRow();
  if(lastRow <= 1) return []; 
  
  const numRowsToFetch = Math.min(15, lastRow - 1);
  const startRow = lastRow - numRowsToFetch + 1;
  const data = logSheet.getRange(startRow, 1, numRowsToFetch, 9).getValues(); 
  
  let logs = [];
  for(let i = data.length - 1; i >= 0; i--) {
    let row = data[i];
    if(row[0] === "") continue;
    
    let pin = parseInt(row[4]); 
    let typeName = "Lainnya";
    
    if ((pin >= 1 && pin <= 4) || pin === 10) typeName = "Kredit";
    else if (pin >= 5 && pin <= 8) typeName = "Debit";
    else if (pin === 9) typeName = "Pinjaman";

    let dateVal = row[2];
    let dStr = "-";
    if (Object.prototype.toString.call(dateVal) === '[object Date]') {
      dStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } else if (dateVal) { dStr = String(dateVal); }

    let rawNominal = typeof row[6] === 'string' ? row[6].replace(/Rp/gi, '').replace(/\./g, '').replace(/,/g, '').trim() : row[6];
    
    logs.push({
      type: typeName,
      kodeAkun: String(row[5] || "-"),
      tanggal: dStr,
      nominal: parseFloat(rawNominal) || 0,
      pic: row[7] || '-',
    });
    if(logs.length === 10) break;
  }
  return logs;
}

// 3. MENYIMPAN DATA DARI FORM
function saveTransaction(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.tipe);
  const logSheet = ss.getSheetByName('Logs'); 
  
  if (!sheet) throw new Error("Sheet " + data.tipe + " tidak ditemukan!");
  if (!logSheet) throw new Error("Sheet 'Logs' tidak ditemukan!");

  let isNewPic = false;

  if (data.tipe === 'Debit' && data.pic && data.pic.trim() !== "") {
    const existingData = sheet.getDataRange().getValues();
    let picExists = false;
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][7] && existingData[i][7].toString().trim().toLowerCase() === data.pic.toLowerCase()) {
        picExists = true; break;
      }
    }
    if (!picExists) isNewPic = true;
  }
  
  const nextNo = getNextNo(data.tipe);
  let formattedDate = data.tanggal; 
  if(data.tanggal && data.tanggal.includes('-')) {
     let p = data.tanggal.split('-'); 
     formattedDate = p[2] + "/" + p[1] + "/" + p[0]; 
  }
  
  const rowData = [
    nextNo, data.ketUtama, formattedDate, data.bulan, data.pin,
    data.kodeAkun, data.nominal, data.pic, data.ketTambahan
  ];

  sheet.appendRow(rowData);
  logSheet.appendRow(rowData);
  
  return { success: true, message: "Data tersimpan di " + data.tipe + " & Logs!", isNewPic: isNewPic, picName: data.pic };
}

function getNextNo(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 1;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 1;
  let lastNo = parseInt(data[data.length - 1][0]);
  return isNaN(lastNo) ? 1 : lastNo + 1;
}