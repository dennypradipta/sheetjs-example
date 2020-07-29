const express = require("express");
const router = express.Router();
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

/* GET Excel files from AOA data Files */
router.get("/excel/aoa", function (req, res) {
  // Definisikan header
  const header = ["ID", "Nama", "Umur"];

  // Definisikan data
  const data = [
    ["1", "Denny", "24"],
    ["2", "Aditya", "25"],
    ["3", "Pradipta", "26"],
    ["4", "Ardhie", "27"],
    ["5", "Putra", "28"],
    ["6", "Prananta", "29"],
  ];

  // Definisikan rows untuk ditulis ke dalam spreadsheet
  const rows = [header, ...data];

  // Buat Workbook
  const fileName = "AOA_XLS";
  let wb = XLSX.utils.book_new();
  wb.Props = {
    Title: fileName,
    Author: "Denny Pradipta",
    CreatedDate: new Date(),
  };

  // Buat Sheet
  wb.SheetNames.push("Sheet 1");

  // Buat Sheet dengan Data
  let ws = XLSX.utils.aoa_to_sheet(rows);
  wb.Sheets["Sheet 1"] = ws;
  __dirname;

  // Cek apakah folder downloadnya ada
  const downloadFolder = path.resolve(__dirname, "../downloads");
  if (!fs.existsSync(downloadFolder)) {
    fs.mkdirSync(downloadFolder);
  }

  try {
    // Simpan filenya
    XLSX.writeFile(
      wb,
      `${(__dirname, "../downloads")}${path.sep}${fileName}.xls`
    );

    res.download(`${downloadFolder}${path.sep}${fileName}.xls`);
  } catch (e) {
    console.log(e.message);
    throw e;
  }
});

/* GET Excel files from AOO data Files */
router.get("/excel/aoo", function (req, res) {
  // Definisikan data
  const data = [
    { id: 1, name: "Denny", age: 24 },
    { id: 2, name: "Aditya", age: 25 },
    { id: 3, name: "Pradipta", age: 26 },
    { id: 4, name: "Ardhie", age: 27 },
    { id: 5, name: "Putra", age: 28 },
    { id: 6, name: "Prananta", age: 29 },
  ];

  // Buat Workbook
  const fileName = "AOO_XLS";
  let wb = XLSX.utils.book_new();
  wb.Props = {
    Title: fileName,
    Author: "Denny Pradipta",
    CreatedDate: new Date(),
  };

  // Buat Sheet
  wb.SheetNames.push("Sheet 1");

  // Buat Sheet dengan Data
  let ws = XLSX.utils.json_to_sheet(data);
  wb.Sheets["Sheet 1"] = ws;

  // Cek apakah folder downloadnya ada
  const downloadFolder = path.resolve(__dirname, "../downloads");
  if (!fs.existsSync(downloadFolder)) {
    fs.mkdirSync(downloadFolder);
  }

  try {
    // Simpan filenya
    XLSX.writeFile(wb, `${downloadFolder}${path.sep}${fileName}.xls`);

    res.download(`${downloadFolder}${path.sep}${fileName}.xls`);
  } catch (e) {
    console.log(e.message);
    throw e;
  }
});

module.exports = router;
