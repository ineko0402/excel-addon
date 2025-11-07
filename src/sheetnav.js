/* global Excel */
export function initSheetNav() {
  loadSheets();
}

async function loadSheets() {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const container = document.getElementById("sheet-panel");
    container.innerHTML = "";

    sheets.items.forEach((sheet) => {
      const div = document.createElement("div");
      div.textContent = sheet.name;
      div.className = "sheet-item";
      div.onclick = () => activateSheet(sheet.name);
      container.appendChild(div);
    });
  });
}

function activateSheet(name) {
  Excel.run(async (context) => {
    context.workbook.worksheets.getItem(name).activate();
    await context.sync();
  }).catch(console.error);
}
