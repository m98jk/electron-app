const XLSX = require("xlsx");
const workbook = XLSX.readFile("./db.xlsx");
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
//
function bringFile() {
  // Convert the worksheet to a JSON object
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 1 });
  const columnsAB = jsonData.map((row) => ({ name: row[0], email: row[1] }));

  let view = document.getElementById("xls");
  view.innerHTML = "";
  columnsAB.forEach((cell) => {
    view.innerHTML += `<tr><td>${cell.name}</td><td>${cell.email} </td><tr>`;
  });
}
//
function addFormData(event) {
  event.preventDefault();
  const listForm = document.getElementById("listForm");
  const fname = document.getElementById("fname").value;
  const email = document.getElementById("email").value;
  // Find the last row in the worksheet
  const lastRow = worksheet["!ref"].split(":")[1];
  const rowIndex = parseInt(lastRow.match(/\d+/)[0], 10);

  // Modify the worksheet by adding a new row
  XLSX.utils.sheet_add_aoa(worksheet, [[fname, email]], {
    origin: rowIndex,
  });
  XLSX.writeFile(workbook, "db.xlsx");
  bringFile();
  listForm.reset();
}
//
// document.getElementById("hit").addEventListener("click", bringFile);
document.getElementById("listForm").addEventListener("submit", addFormData);
