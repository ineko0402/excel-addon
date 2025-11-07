import { initCalendar } from "./calendar.js";
import { initSheetNav } from "./sheetnav.js";

Office.onReady(() => {
  initCalendar();
  initSheetNav();
  setupTabs();
});

function setupTabs() {
  const calendarPanel = document.getElementById("calendar-panel");
  const sheetPanel = document.getElementById("sheet-panel");

  document.getElementById("tab-calendar").addEventListener("click", () => {
    switchPanel(calendarPanel, sheetPanel, "tab-calendar");
  });
  document.getElementById("tab-sheet").addEventListener("click", () => {
    switchPanel(sheetPanel, calendarPanel, "tab-sheet");
  });
}

function switchPanel(show, hide, activeId) {
  hide.classList.add("hidden");
  show.classList.remove("hidden");
  document
    .querySelectorAll("#toolbar button")
    .forEach((btn) => btn.classList.remove("active"));
  document.getElementById(activeId).classList.add("active");
}
