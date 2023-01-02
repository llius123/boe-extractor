const fs = require("fs");
const { XMLParser } = require("fast-xml-parser");

const allTables = getDataFromBoe();

const { excel, excelPage } = createAndConfigureExcel();

let row = 1;

// Add headers to excel page
createHeader(excelPage);

// Move to next row after adding headers
row++;

// Write data from xml to the excel page
writeDataOnSpecificExcelPage();

// GRoup all excelpages on excel
excel.write("boe.xlsx");

function createHeader(excelPage) {
  const headers = [
    null,
    "Marca",
    "Modelo",
    "Periodo comercial",
    "CC",
    "N.º Cilidros",
    "G/D",
    "P kW",
    "cvf",
    "cv",
    "valor",
  ];
  for (let index = 1; index < 11; index++) {
    excelPage.cell(row, index).string(headers[index]);
  }
}

function createBody(marca, tr, currentWs) {
  if (!Array.isArray(tr)) {
    addBody(currentWs, tr.td, marca);
    return;
  }
  for (const { td } of tr) {
    addBody(currentWs, td, marca);
  }
}

function addBody(cell, data, titulo) {
  cell.cell(row, 1).string(titulo);

  const modelo = `${data[0]["#text"]}`;
  cell.cell(row, 2).string(modelo.replace(/\s+/g, " "));

  let data1 = data[1]["#text"];
  let data2 = data[2]["#text"];
  if (data1 === undefined) {
    data1 = "";
  }
  if (data2 === undefined) {
    data2 = "";
  }

  cell.cell(row, 3).string(data1.toString() + "-" + data2.toString());

  let data3 = `${data[3]["#text"]}`;
  if (data3 === "undefined") {
    data3 = "";
  }
  cell.cell(row, 4).string(data3);

  let data4 = `${data[4]["#text"]}`;
  if (data4 === "undefined") {
    data4 = "";
  }
  cell.cell(row, 5).string(data4);

  let data5 = `${data[5]["#text"]}`;
  if (data5 === "undefined") {
    data5 = "";
  }
  cell.cell(row, 6).string(data5);

  let data6 = `${data[6]["#text"]}`;
  if (data6 === "undefined") {
    data6 = "";
  }
  cell.cell(row, 7).string(data6);

  let data7 = `${data[7]["#text"]}`;
  if (data7 === "undefined") {
    data7 = "";
  }
  cell.cell(row, 8).string(data7);

  let data8 = `${data[8]["#text"]}`;
  if (data8 === "undefined") {
    data8 = "";
  }
  cell.cell(row, 9).string(data8);

  let data9 = `${data[9]["#text"]}`;
  if (data9 === "undefined") {
    data9 = "";
  }
  cell.cell(row, 10).string(data9);
  row++;
}

function getDataFromBoe() {
  const dataXML = readDataFromXML();
  const parser = new XMLParser({ ignoreAttributes: false });
  let jsonObj = parser.parse(dataXML);
  return jsonObj.documento.texto.table;

  function readDataFromXML() {
    let data;
    try {
      data = fs.readFileSync("boe.xml", "utf8");
    } catch (err) {
      console.error(err);
    }
    return data;
  }
}

function createAndConfigureExcel() {
  //Create excel
  var xl = require("excel4node");

  // Create a new instance of a Workbook class
  var excel = new xl.Workbook();

  // Add Worksheets to the workbook
  var excelPage = excel.addWorksheet("Coches");
  return { excel, excelPage };
}

function writeDataOnSpecificExcelPage() {
  for (const table of allTables) {
    if (table.thead.tr[0] === undefined) {
      break;
    }
    const marca = table.thead.tr[0].th["#text"].replace(/Marca: /g, "");

    createBody(marca, table.tbody.tr, excelPage);
  }
}
