const XLSX = require("xlsx");
const file_name = "./db.xlsx";
const workbook = XLSX.readFile(file_name);
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
//
function bringFile() {
  // Convert the worksheet to a JSON object
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 1 });
  const columnsAB = jsonData.map((row) => ({ name: row[0], email: row[1] }));

  let view = document.getElementById("xls");
  view.innerHTML = "";
  columnsAB.forEach((cell,index) => {
    view.innerHTML += `<tr><td>${cell.name}</td><td>${cell.email} </td><td><button  onclick="stbadge(${index})">Print</button></td><tr>`;
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
  XLSX.writeFile(workbook, file_name);
  bringFile();
  listForm.reset();
}
//
document.getElementById("m98jk").addEventListener("click", () => {
  const { shell } = require("electron");

  shell.openExternal("https://github.com/m98jk/electron-app");
});
document.getElementById("listForm").addEventListener("submit", addFormData);

function makeBadge() {
  const badge = document.getElementById("badge");
  const badgeName = "Mohammed Jawad";
  badge.innerHTML = `<div class="badge"><span>${badgeName}</span></div>`;
}
const badges = document.getElementsByClassName("stbadge");
for (let i = 0; i < badges.length; i++) {
  badges[i].addEventListener("click", makeBadge);
  console.log(badges[i]);
  
}

function stbadge(id){
console.log(`Print Badge ${id}`);

}