function onEdit(e) {
  const ss          = e.source;
  const sh          = e.range.getSheet();
  const sheetName   = 'Kitchen Production';
  const invName     = 'Current Inventory';
  const parName     = 'Par Level';
  const ordersSheet = 'Orders & Retail';
  const distShName  = 'DC-Cannabinoid';

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row <= 1) return;  // skip header row

  // Kitchen Production columns
  const dateCol     = 1; // A
  const productCol  = 2; // B
  const quantityCol = 3; // C
  const dcCol       = 4; // D
  const batchCol    = 5; // E

  // Orders & Retail columns (with Order Date in C)
  const orProductCol = 2; // B: Product
  const orDateCol    = 3; // C: Order Date
  const orQtySoldCol = 4; // D: Quantity Sold
  const orBatchCol   = 5; // E: Batch ID dropdown
  const orStatusCol  = 8; // H: Status

  //
  // Part 1: Kitchen Production
  //
  if (sh.getName() === sheetName) {

    // A) PRODUCT pick → stamp date + rebuild DC dropdown
    if (col === productCol && e.value) {
      // stamp today’s date
      sh.getRange(row, dateCol).setValue(new Date());
      sh.getRange(2, dcCol, sh.getLastRow() - 1).clearDataValidations();

      const rawProduct = e.value.toString();
      // includes ND for NonDosed
      const doseCode   = (rawProduct.match(/(D8|D9|THCO|FS|CAF|ND)/) || [])[1];
      if (doseCode) {
        const distSh = ss.getSheetByName(distShName);
        if (distSh) {
          const last = distSh.getLastRow();
          const distData = distSh.getRange(2, 2, last - 1, 2).getDisplayValues();
          const opts = distData
            .filter(r => r[1] === doseCode)
            .map(r => `${r[0]}-${r[1]}`);

          const dcCell = sh.getRange(row, dcCol);
          dcCell.clearContent().setNumberFormat('@').clearDataValidations();
          if (opts.length) {
            const rule = SpreadsheetApp.newDataValidation()
              .requireValueInList(opts, true)
              .setAllowInvalid(false)
              .build();
            dcCell.setDataValidation(rule);
          }
        }
      }
    }

    // B) DC pick → generate batch code & bump inventory
    if (col === dcCol && e.value) {
      const dateVal = sh.getRange(row, dateCol).getValue();
      if (!(dateVal instanceof Date)) return;

      const rawProduct  = sh.getRange(row, productCol).getValue().toString();
      const productCode = rawProduct.replace(/\s+/g, '-');
      const dcDisplay   = e.range.getDisplayValue();
      const dc = dcDisplay.includes('-')
               ? dcDisplay.slice(0, dcDisplay.lastIndexOf('-'))
               : dcDisplay;

      const qtyMade = Number(sh.getRange(row, quantityCol).getValue());
      if (isNaN(qtyMade)) return;

      // count existing today’s batches
      const lastKP  = sh.getLastRow();
      const prods   = sh.getRange(2, productCol, lastKP - 1, 1).getValues().flat();
      const dates   = sh.getRange(2, dateCol,    lastKP - 1, 1).getValues().flat();
      const dcsDisp = sh.getRange(2, dcCol,      lastKP - 1, 1).getDisplayValues().flat();
      const dcsArr  = dcsDisp.map(v => v.includes('-') ? v.slice(0, v.lastIndexOf('-')) : v);

      let count = 0;
      prods.forEach((p, i) => {
        const d = dates[i];
        if (
          p === rawProduct &&
          dcsArr[i] === dc &&
          d instanceof Date &&
          d.getFullYear() === dateVal.getFullYear() &&
          d.getMonth()    === dateVal.getMonth() &&
          d.getDate()     === dateVal.getDate()
        ) count++;
      });

      const fmtDate = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'MM-dd-yy');
      const code = `${productCode}-${fmtDate}-DC-${dc}.${count}`;
      sh.getRange(row, batchCol).setValue(code);

      // bump Current Inventory & recolor
      const invSh    = ss.getSheetByName(invName);
      if (!invSh) return;
      const lastInv  = invSh.getLastRow();
      const invProds = invSh.getRange(2, 1, lastInv - 1, 1).getValues().flat();
      const invIdx   = invProds.findIndex(p => p === rawProduct);

      if (invIdx !== -1) {
        const qtyCell    = invSh.getRange(invIdx + 2, 2);
        const currentQty = Number(qtyCell.getValue()) || 0;
        const newQty     = currentQty + qtyMade;
        qtyCell.setValue(newQty);
        highlightInventoryCell(ss.getSheetByName(parName), rawProduct, newQty, qtyCell);
      } else {
        invSh.appendRow([rawProduct, qtyMade]);
        const newRow     = invSh.getLastRow();
        const newQtyCell = invSh.getRange(newRow, 2);
        highlightInventoryCell(ss.getSheetByName(parName), rawProduct, qtyMade, newQtyCell);
      }
    }
  }

  //
  // Part 2: Manual edits in Current Inventory → recolor
  //
  if (sh.getName() === invName && col === 2) {
    const rawProduct = sh.getRange(row, 1).getValue().toString();
    const newQty     = Number(e.range.getValue());
    if (!isNaN(newQty)) {
      highlightInventoryCell(ss.getSheetByName(parName), rawProduct, newQty, e.range);
    }
  }

  //
  // Part 3: Orders & Retail
  //
  if (sh.getName() === ordersSheet) {
    // a) PRODUCT pick → stamp Order Date, clear stale, then populate Batch ID
    if (col === orProductCol && e.value) {
      // stamp today’s date in C
      sh.getRange(row, orDateCol).setValue(new Date());

      // clear out old dropdowns in E
      sh.getRange(2, orBatchCol, sh.getLastRow() - 1).clearDataValidations();

      // build & apply the 3 most recent batch codes
      const rawProduct = e.value.toString();
      const kpSh       = ss.getSheetByName(sheetName);
      const lastKP     = kpSh.getLastRow();
      const prodArr    = kpSh.getRange(2, productCol, lastKP - 1, 1).getValues().flat();
      const dateArr    = kpSh.getRange(2, dateCol,    lastKP - 1, 1).getValues().flat();
      const codeArr    = kpSh.getRange(2, batchCol,   lastKP - 1, 1).getValues().flat().map(String);

      const rows = [];
      prodArr.forEach((p, i) => {
        if (p === rawProduct && codeArr[i]) rows.push({ date: dateArr[i], code: codeArr[i] });
      });
      rows.sort((a, b) => b.date - a.date);
      const recentCodes = rows.slice(0, 3).map(r => r.code);

      const batchCell = sh.getRange(row, orBatchCol);
      batchCell.clearContent();
      if (recentCodes.length) {
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(recentCodes, true)
          .setAllowInvalid(false)
          .build();
        batchCell.setDataValidation(rule);
      }
    }

    // b) STATUS = Complete → subtract Quantity Sold from inventory
    if (col === orStatusCol && e.value === 'Complete') {
      const rawProduct = sh.getRange(row, orProductCol).getValue().toString();
      const soldQty    = Number(sh.getRange(row, orQtySoldCol).getValue());
      if (!isNaN(soldQty)) {
        const invSh  = ss.getSheetByName(invName);
        const lastInv= invSh.getLastRow();
        const invArr = invSh.getRange(2, 1, lastInv - 1, 1).getValues().flat();
        const idx    = invArr.findIndex(p => p === rawProduct);
        if (idx !== -1) {
          const cell    = invSh.getRange(idx + 2, 2);
          const current = Number(cell.getValue()) || 0;
          const updated = Math.max(0, current - soldQty);
          cell.setValue(updated);
          highlightInventoryCell(ss.getSheetByName(parName), rawProduct, updated, cell);
        }
      }
    }
  }
}

// helper: color + font based on Par Level thresholds
function highlightInventoryCell(parSh, productName, qty, cell) {
  const parLast = parSh.getLastRow();
  const parData = parSh.getRange(2, 1, parLast - 1, 2).getValues();
  const match   = parData.find(r => r[0] === productName);
  if (!match) return;
  const parValue = Number(match[1]);
  if (isNaN(parValue) || parValue <= 0) return;

  let bg;
  if      (qty >= parValue * 1.75) bg = '#006400';
  else if (qty >= parValue * 1.50) bg = '#228B22';
  else if (qty >= parValue * 1.25) bg = '#90EE90';
  else if (qty >= parValue * 1.00) bg = '#00FF00';
  else if (qty >= parValue * 0.75) bg = '#FFC0CB';
  else if (qty >= parValue * 0.50) bg = '#FF6666';
  else if (qty >= parValue * 0.25) bg = '#FF0000';
  else                              bg = '#990000';

  cell.setBackground(bg);
  cell.setFontColor(
    (qty >= parValue * 1.50 || qty <= parValue * 0.25)
      ? '#FFFFFF'
      : '#000000'
  );
}

// onOpen: color all existing inventory rows on load
function onOpen() {
  const ss    = SpreadsheetApp.getActive();
  const invSh = ss.getSheetByName('Current Inventory');
  const parSh = ss.getSheetByName('Par Level');
  const last  = invSh.getLastRow();
  const data  = invSh.getRange(2, 1, last - 1, 2).getValues();

  data.forEach((r, i) => {
    const prod = r[0].toString();
    const qty  = Number(r[1]);
    if (!isNaN(qty)) {
      const cell = invSh.getRange(i + 2, 2);
      highlightInventoryCell(parSh, prod, qty, cell);
    }
  });
}
