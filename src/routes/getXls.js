const ExcelJS = require('exceljs');

module.exports = async (req, res) => {
    const xls = await getQuoteAsXls(req);
    res.send(xls);
};

const itemsJSONtoItemsTable = (itemsJSON) => {
    const parsed = JSON.parse(itemsJSON);
    const res = parsed.map((item) => ({
      ...item,
      basePrice: parseFloat(item.basePrice),
      discPrice: parseFloat(item.discPrice),
      productVariation: parseInt(item.productVariation),
      count: parseInt(item.count),
      value: parseFloat(item.value),
    }));
    return res;
  };
  

const populateQuoteHeader = (ws, q) => {
  let rowIndex = 3;
  let row = ws.getRow(rowIndex);

  row.values = ['#', 'Назва', 'Cтворено', '', 'Оновлено', ''];
  row.font = { bold: true };

  const columnWidths = [10, 45, 7, 8, 9, 10];

  row.eachCell((cell, colNumber) => {
    const columnIndex = colNumber - 1;
    const columnWidth = columnWidths[columnIndex];
    ws.getColumn(colNumber).width = columnWidth;
  });

  rowIndex += 1;
  row = ws.getRow(rowIndex);
  row.values = [q.quoteId, q.title, q.createdAt, '', q.updatedAt, ''];

  rowIndex += 1;
  row = ws.getRow(rowIndex);
  row.values = ['', '', q.createdBy, '', q.updatedBy, ''];

  ws.mergeCells(`C3:D3`);
  ws.mergeCells(`C4:D4`);
  ws.mergeCells(`E3:F3`);
  ws.mergeCells(`E4:F4`);

  rowIndex += 2;
  row = ws.getRow(rowIndex);
  row.values = ['Для', q.volunteerId];
};

const populateItems = (ws, q) => {
  // Initialize the row index
  let rowIndex = 10;

  const row = ws.getRow(rowIndex);
  row.values = ['Product Id', 'Title', 'Count', 'Price', 'Sum', 'Verification'];
  row.font = { bold: true };

  // Loop over the items
  const items = itemsJSONtoItemsTable(q.itemsTable);
  items.forEach((item, index) => {
    const valueRowIndex = rowIndex + index + 1;
    const row = ws.getRow(valueRowIndex);
    row.getCell('A').value = item.productId;
    row.getCell('B').value = item.title;
    row.getCell('C').value = item.count;
    row.getCell('D').value = item.discPrice;
    row.getCell('E').value = item.value;
    row.getCell('F').value = {
      formula: `C${valueRowIndex}*D${valueRowIndex}`,
      result: item.count * item.discPrice,
    };

    row.getCell('B').alignment = { wrapText: true };
  });
  rowIndex += items.length;
};

const drawAttributes = (ws, q) => {
  let rowIndex =
    itemsJSONtoItemsTable(q.itemsTable).length + 12;
  let row = ws.getRow(rowIndex);
  row.values = ['Вартість засобів', '', '', '', q.itemsCost];
  row.font = { bold: true };
  row = ws.getRow(++rowIndex);
  row.values = ['Доставка', '', '', '', q.deliveryCost];
  row.font = { bold: true };
  row = ws.getRow(++rowIndex);
  row.values = ['До сплати', '', '', '', q.totalValue];
  row.font = { bold: true };

  rowIndex++;

  row = ws.getRow(++rowIndex);
  row.values = ['', 'Адреса доставки', 'НП №', 'Телефон', 'ПІБ'];
  row.font = { bold: true };

  row = ws.getRow(++rowIndex);
  row.values = [
    '',
    q.deliveryAddress,
    q.deliveryServiceDepartmentId,
    q.deliveryPhoneNumber,
    q.deliveryName,
  ];

  // TODO: Інвойси
  // TODO: Трекінг
  // TODO: Офіційний запит
  // TODO: АКТ П/П
};

const drawBorders = (ws) => {
  // Define the border style
  const borderStyle = {
    style: 'thin', // 'thin', 'medium', 'thick', etc
    color: { argb: '00000000' },
  };

  // Loop through all cells and apply the border style
  ws.eachRow((row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.border = {
        top: borderStyle,
        bottom: borderStyle,
      };
    });
  });
};

// XLS Export
const getQuoteAsXls = (q) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(`Quote-${q.quoteId}`, {
    pageSetup: { paperSize: 9, orientation: 'landscape' },
  });
  try {
    populateItems(worksheet, q);
    populateQuoteHeader(worksheet, q);
    drawBorders(worksheet);
    drawAttributes(worksheet, q);

    return workbook.xlsx.writeBuffer();
  } catch (err) {
    console.log(err);
    return workbook.xlsx.writeBuffer();
  }
};
