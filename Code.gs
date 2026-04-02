function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const text = e.postData.contents;

  // пробуем распарсить JSON
  let data;
  try {
    data = JSON.parse(text);
  } catch (err) {
    data = [];
  }

  if (!Array.isArray(data)) {
    data = [data];
  }

  data.forEach(row => {
    sheet.appendRow([
      row.period || '',
      row.order || '',
      row.shipment || '',
      row.writeoff || '',
      row.profit || ''
    ]);
  });

  return ContentService.createTextOutput("OK");
}
