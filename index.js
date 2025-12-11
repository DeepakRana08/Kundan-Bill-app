const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const PORT = 3000;

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// MAIN PAGE
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// SAVE INVOICE TO EXCEL
app.post('/save', async (req, res) => {
  try {
    console.log('=== SAVING INVOICE ===');
    const raw = req.body.invoiceData || '{}';
    const inv = JSON.parse(raw);
    console.log('Invoice Data:', inv);

    const workbook = new ExcelJS.Workbook();
    const filePath = path.join(__dirname, 'bills.xlsx');

    // Load existing file or create new
    try {
      await workbook.xlsx.readFile(filePath);
      console.log('✅ Existing file loaded');
    } catch (e) {
      console.log('📄 New file will be created');
    }

    // INVOICE SUMMARY SHEET
    let summarySheet = workbook.getWorksheet('Invoice Summary');
    if (!summarySheet) {
      summarySheet = workbook.addWorksheet('Invoice Summary');
      summarySheet.addRow([
        'BillRSID', 'Customer Name', 'Phone', 'GSTIN', 'Place of Supply',
        'Invoice No', 'Invoice Date', 'Total Qty', 'Discount', 'Total Tax',
        'Grand Total', 'Received Amount', 'Due Balance', 'Created Date'
      ]);
    }

    const rsid = 'RSID-' + Date.now();
    const todayIso = new Date().toISOString().slice(0, 10);

    summarySheet.addRow([
      rsid,
      inv.customerName || 'N/A',
      inv.phone || 'N/A',
      inv.gst || 'N/A',
      inv.place || 'N/A',
      rsid,                               // Invoice No me bhi RSID store
      inv.invDate || todayIso,
      parseFloat(inv.totalQty) || 0,
      parseFloat(inv.discount) || 0,
      parseFloat(inv.totalTax) || 0,
      parseFloat(inv.grandTotal) || 0,
      parseFloat(inv.received) || 0,
      parseFloat(inv.dueBalance) || 0,
      todayIso
    ]);

    // ITEMS DETAILS SHEET
    let itemsSheet = workbook.getWorksheet('Items Details');
    if (!itemsSheet) {
      itemsSheet = workbook.addWorksheet('Items Details');
      itemsSheet.addRow([
        'BillRSID', 'Sr No', 'Item Name', 'Item Size', 'Quantity (KG)',
        'Price/Unit', 'Tax/Unit', 'Amount'
      ]);
    }

    const itemsArray = inv.items || [];
    itemsArray.forEach((item, index) => {
      itemsSheet.addRow([
        rsid,
        item.sr || (index + 1),
        item.name || 'N/A',
        item.size || 'N/A',
        parseFloat(item.qty) || 0,
        parseFloat(item.price) || 0,
        parseFloat(item.tax) || 0,
        parseFloat(item.amount) || 0
      ]);
    });

    await workbook.xlsx.writeFile(filePath);
    console.log('✅ SAVED! RSID:', rsid);

    // RSID front‑end ko bhejo
    res.json({ success: true, rsid });
  } catch (err) {
    console.error('❌ ERROR:', err);
    res.json({ success: false, error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`✅ Server: http://localhost:${PORT}`);
});
