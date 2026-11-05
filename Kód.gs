const SHEET_NAME = "Platby";
const HISTORY_SHEET_NAME = "Historia_Zmien";
const ARCHIVE_SHEET_NAME = "Archiv_Platieb";

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["id", "nazov", "kategoria", "suma", "frekvencia", "najblizsia_platba", "poznamka", "koniec_viazanosti", "email_filter"]);
  }
  return sheet;
}

function getHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET_NAME);
    sheet.appendRow(["id", "payment_id", "datum_zmeny", "nazov", "stara_suma", "nova_suma", "rozdiel", "typ_akcie"]);
  }
  return sheet;
}

function getArchiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ARCHIVE_SHEET_NAME);
    sheet.appendRow(["log_id", "payment_id", "datum_uhrady", "nazov", "kategoria", "suma", "povodne_splatne"]);
  }
  return sheet;
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Paušály & Správa zmlúv')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPayments() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return { success: true, data: [] };

  const tz = Session.getScriptTimeZone();
  let items = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    let platbaDate = row[5] instanceof Date ? Utilities.formatDate(row[5], tz, 'yyyy-MM-dd') : String(row[5] || '').trim();
    let viazanostDate = row[7] instanceof Date ? Utilities.formatDate(row[7], tz, 'yyyy-MM-dd') : String(row[7] || '').trim();

    items.push({
      id: String(row[0]), nazov: row[1] || '', kategoria: row[2] || 'Iné',
      suma: parseFloat(row[3]) || 0, frekvencia: row[4] || 'Mesačne',
      najblizsia_platba: platbaDate, poznamka: row[6] || '',
      koniec_viazanosti: viazanostDate, email_filter: row[8] || ''
    });
  }
  return { success: true, data: items };
}

function getPaymentHistory(paymentId) {
  const sheet = getHistorySheet();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return { success: true, data: [] };

  const tz = Session.getScriptTimeZone();
  let history = [];

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === String(paymentId)) {
      let dStr = values[i][2] instanceof Date ? Utilities.formatDate(values[i][2], tz, 'yyyy-MM-dd HH:mm') : values[i][2];
      history.push({
        id: values[i][0], payment_id: values[i][1], datum: dStr, nazov: values[i][3],
        stara_suma: parseFloat(values[i][4]) || 0, nova_suma: parseFloat(values[i][5]) || 0,
        rozdiel: parseFloat(values[i][6]) || 0, typ_akcie: values[i][7] || 'Zmena ceny'
      });
    }
  }
  return { success: true, data: history };
}

function addPayment(item) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();
    const id = Utilities.getUuid();
    const initialSuma = parseFloat(item.suma) || 0;
    
    sheet.appendRow([
      id, item.nazov || '', item.kategoria || 'Iné', initialSuma, item.frekvencia || 'Mesačne',
      item.najblizsia_platba || '', item.poznamka || '', item.koniec_viazanosti || '', item.email_filter || ''
    ]);

    const histSheet = getHistorySheet();
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    histSheet.appendRow([Utilities.getUuid(), id, todayStr, item.nazov || '', 0, initialSuma, initialSuma, 'Vytvorenie paušálu']);
    
    return { success: true, id: id };
  } catch(e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updatePayment(item) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(item.id)) {
        const rowIndex = i + 1;
        const staraSuma = parseFloat(values[i][3]) || 0;
        const novaSuma = parseFloat(item.suma) || 0;

        if (staraSuma !== novaSuma) {
          const histSheet = getHistorySheet();
          const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
          histSheet.appendRow([Utilities.getUuid(), item.id, todayStr, item.nazov || values[i][1], staraSuma, novaSuma, novaSuma - staraSuma, 'Manuálna zmena ceny']);
        }

        sheet.getRange(rowIndex, 2).setValue(item.nazov || '');
        sheet.getRange(rowIndex, 3).setValue(item.kategoria || 'Iné');
        sheet.getRange(rowIndex, 4).setValue(novaSuma);
        sheet.getRange(rowIndex, 5).setValue(item.frekvencia || 'Mesačne');
        sheet.getRange(rowIndex, 6).setValue(item.najblizsia_platba || '');
        sheet.getRange(rowIndex, 7).setValue(item.poznamka || '');
        sheet.getRange(rowIndex, 8).setValue(item.koniec_viazanosti || '');
        sheet.getRange(rowIndex, 9).setValue(item.email_filter || '');

        return { success: true };
      }
    }
    return { success: false, error: 'Položka sa nenašla' };
  } catch(e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

function deletePayment(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Položka sa nenašla' };
  } catch(e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

function markAsPaid(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();
    const values = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        const rowIndex = i + 1;
        const nazov = values[i][1];
        const kategoria = values[i][2];
        const suma = parseFloat(values[i][3]) || 0;
        const frekvencia = values[i][4] || 'Mesačne';
        
        let currentDate = values[i][5] instanceof Date ? new Date(values[i][5]) : new Date(values[i][5]);
        if (isNaN(currentDate.getTime())) currentDate = new Date();
        const povodneSplatne = Utilities.formatDate(currentDate, tz, 'yyyy-MM-dd');

        // 1. Zápis do Archívu (Infinite History)
        const archiveSheet = getArchiveSheet();
        archiveSheet.appendRow([Utilities.getUuid(), id, todayStr, nazov, kategoria, suma, povodneSplatne]);

        // 2. Posun dátumu na ďalší cyklus
        if (frekvencia === 'Mesačne') currentDate.setMonth(currentDate.getMonth() + 1);
        else if (frekvencia === 'Kvartálne') currentDate.setMonth(currentDate.getMonth() + 3);
        else if (frekvencia === 'Polročne') currentDate.setMonth(currentDate.getMonth() + 6);
        else if (frekvencia === 'Ročne') currentDate.setFullYear(currentDate.getFullYear() + 1);

        const newDateStr = Utilities.formatDate(currentDate, tz, 'yyyy-MM-dd');
        sheet.getRange(rowIndex, 6).setValue(newDateStr);

        return { success: true, nextDate: newDateStr };
      }
    }
    return { success: false, error: 'Položka sa nenašla' };
  } catch(e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

function autoUpdateFromEmails() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet();
    const values = sheet.getDataRange().getValues();
    const histSheet = getHistorySheet();
    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

    for (let i = 1; i < values.length; i++) {
      const id = values[i][0];
      const nazov = values[i][1];
      const staraSuma = parseFloat(values[i][3]) || 0;
      const emailFilter = String(values[i][8] || '').trim();

      if (emailFilter.length > 5) {
        const threads = GmailApp.search(emailFilter + ' newer_than:3d', 0, 1);
        if (threads.length > 0) {
          const msg = threads[0].getMessages()[0];
          const body = msg.getPlainBody();
          const regex = /(?:suma|úhrada|k úhrade|spolu)[:\s]*(\d+[.,]\d{2})\s*(?:€|EUR)/i;
          const match = body.match(regex) || msg.getSubject().match(regex);
          
          if (match && match[1]) {
            const novaSuma = parseFloat(match[1].replace(',', '.'));
            if (novaSuma > 0 && novaSuma !== staraSuma) {
              sheet.getRange(i + 1, 4).setValue(novaSuma);
              histSheet.appendRow([Utilities.getUuid(), id, todayStr, nazov, staraSuma, novaSuma, novaSuma - staraSuma, 'Automatická aktualizácia z e-mailu']);
            }
          }
        }
      }
    }
  } catch(e) { console.error(e); }
  finally { lock.releaseLock(); }
}
