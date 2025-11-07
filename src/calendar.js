/* global Excel, flatpickr */
export function initCalendar() {
  const el = document.getElementById("calendar-panel");
  flatpickr(el, {
    inline: true,
    dateFormat: "Y/m/d",
    onChange: ([date], dateStr) => {
      if (dateStr) insertDate(dateStr);
    },
  });
}

function insertDate(dateStr) {
  Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.values = [[dateStr]];
    range.numberFormat = [["yyyy/mm/dd"]];
    await context.sync();
  }).catch(console.error);
}
