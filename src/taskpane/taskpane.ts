/* global Office, Excel */

/**
 * Excelユーティリティ
 * シート移動と日付入力をタスクペインから行う。
 */

Office.onReady((officeInfo) => {
  if (officeInfo.host !== Office.HostType.Excel) {
    return;
  }

  initializeTabs();
  initializeSheetNavigator();
  initializeCalendar();

  void loadSheets();
});

/**
 * 機能切替タブを初期化する。
 */
function initializeTabs(): void {
  const sheetTab: HTMLElement | null =
    document.getElementById("tab-sheet");

  const calendarTab: HTMLElement | null =
    document.getElementById("tab-calendar");

  sheetTab?.addEventListener("click", () => {
    switchPanel("sheet");
    void loadSheets();
  });

  calendarTab?.addEventListener("click", () => {
    switchPanel("calendar");
  });
}

/**
 * シートナビの操作を初期化する。
 */
function initializeSheetNavigator(): void {
  const refreshButton: HTMLElement | null =
    document.getElementById("refresh-sheets");

  refreshButton?.addEventListener("click", () => {
    void loadSheets();
  });
}

/**
 * カレンダーの操作を初期化する。
 */
function initializeCalendar(): void {
  const todayButton: HTMLElement | null =
    document.getElementById("set-today");

  const insertButton: HTMLElement | null =
    document.getElementById("insert-date");

  todayButton?.addEventListener("click", () => {
    setToday();
  });

  insertButton?.addEventListener("click", () => {
    void insertDate();
  });

  setToday();
}

/**
 * 表示する機能を切り替える。
 */
function switchPanel(panelName: "sheet" | "calendar"): void {
  const sheetPanel: HTMLElement | null =
    document.getElementById("sheet-panel");

  const calendarPanel: HTMLElement | null =
    document.getElementById("calendar-panel");

  const sheetTab: HTMLElement | null =
    document.getElementById("tab-sheet");

  const calendarTab: HTMLElement | null =
    document.getElementById("tab-calendar");

  const sheetSelected: boolean = panelName === "sheet";

  sheetPanel?.classList.toggle("hidden", !sheetSelected);
  calendarPanel?.classList.toggle("hidden", sheetSelected);

  sheetTab?.classList.toggle("active", sheetSelected);
  calendarTab?.classList.toggle("active", !sheetSelected);

  sheetTab?.setAttribute(
    "aria-selected",
    String(sheetSelected)
  );

  calendarTab?.setAttribute(
    "aria-selected",
    String(!sheetSelected)
  );

  setMessage("");
}

/**
 * ワークシート一覧を読み込んで表示する。
 */
async function loadSheets(): Promise<void> {
  setMessage("シートを読み込んでいます。");

  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const worksheets: Excel.WorksheetCollection =
        context.workbook.worksheets;

      const activeWorksheet: Excel.Worksheet =
        worksheets.getActiveWorksheet();

      worksheets.load("items/name");
      activeWorksheet.load("name");

      await context.sync();

      displaySheets(
        worksheets.items,
        activeWorksheet.name
      );
    });

    setMessage("");
  } catch (error: unknown) {
    handleError(error);
  }
}

/**
 * ワークシート一覧をタスクペインへ表示する。
 */
function displaySheets(
  worksheets: Excel.Worksheet[],
  activeSheetName: string
): void {
  const sheetList: HTMLElement | null =
    document.getElementById("sheet-list");

  if (sheetList === null) {
    return;
  }

  sheetList.replaceChildren();

  worksheets.forEach((worksheet: Excel.Worksheet) => {
    const sheetButton: HTMLButtonElement =
      document.createElement("button");

    sheetButton.type = "button";
    sheetButton.textContent = worksheet.name;
    sheetButton.className = "sheet-item";

    if (worksheet.name === activeSheetName) {
      sheetButton.classList.add("active");
    }

    sheetButton.addEventListener("click", () => {
      void activateSheet(worksheet.name);
    });

    sheetList.appendChild(sheetButton);
  });
}

/**
 * 指定したワークシートをアクティブにする。
 */
async function activateSheet(
  sheetName: string
): Promise<void> {
  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const worksheet: Excel.Worksheet =
        context.workbook.worksheets.getItem(sheetName);

      worksheet.activate();

      await context.sync();
    });

    await loadSheets();
  } catch (error: unknown) {
    handleError(error);
  }
}

/**
 * 日付入力欄へ今日の日付を設定する。
 */
function setToday(): void {
  const dateInput: HTMLInputElement | null =
    getDateInput();

  if (dateInput === null) {
    return;
  }

  const today: Date = new Date();

  const year: string =
    String(today.getFullYear());

  const month: string =
    String(today.getMonth() + 1).padStart(2, "0");

  const day: string =
    String(today.getDate()).padStart(2, "0");

  dateInput.value = `${year}-${month}-${day}`;
}

/**
 * 選択範囲の左上セルへ日付を入力する。
 */
async function insertDate(): Promise<void> {
  const dateInput: HTMLInputElement | null =
    getDateInput();

  if (dateInput === null || dateInput.value === "") {
    setMessage("入力する日付を選択してください。", true);
    return;
  }

  const excelDate: number =
    convertToExcelDate(dateInput.value);

  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const selectedRange: Excel.Range =
        context.workbook.getSelectedRange();

      // 選択範囲の大きさによるエラーを避けるため、左上セルだけを対象にする。
      const targetCell: Excel.Range =
        selectedRange.getCell(0, 0);

      targetCell.values = [[excelDate]];
      targetCell.numberFormat = [["yyyy/mm/dd"]];

      await context.sync();
    });

    setMessage(`${dateInput.value} を入力しました。`);
  } catch (error: unknown) {
    handleError(error);
  }
}

/**
 * YYYY-MM-DDをExcelの日付シリアル値へ変換する。
 */
function convertToExcelDate(dateText: string): number {
  const dateParts: number[] =
    dateText.split("-").map(Number);

  const year: number = dateParts[0];
  const month: number = dateParts[1];
  const day: number = dateParts[2];

  // UTCで計算し、タイムゾーンや夏時間による日付のずれを防止する。
  const milliseconds: number =
    Date.UTC(year, month - 1, day);

  const millisecondsPerDay: number =
    24 * 60 * 60 * 1000;

  return milliseconds / millisecondsPerDay + 25569;
}

/**
 * 日付入力欄を取得する。
 */
function getDateInput(): HTMLInputElement | null {
  const dateElement: HTMLElement | null =
    document.getElementById("date-input");

  if (dateElement instanceof HTMLInputElement) {
    return dateElement;
  }

  return null;
}

/**
 * タスクペインへメッセージを表示する。
 */
function setMessage(
  message: string,
  isError: boolean = false
): void {
  const messageElement: HTMLElement | null =
    document.getElementById("message");

  if (messageElement === null) {
    return;
  }

  messageElement.textContent = message;
  messageElement.classList.toggle("error", isError);
}

/**
 * エラー内容を利用者向けに表示する。
 */
function handleError(error: unknown): void {
  console.error(error);

  if (error instanceof Error) {
    setMessage(`エラー：${error.message}`, true);
    return;
  }

  setMessage("処理中にエラーが発生しました。", true);
}