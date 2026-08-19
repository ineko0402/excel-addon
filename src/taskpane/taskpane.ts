/* global Office, Excel */

/**
 * Excelユーティリティ
 * シート移動、日付入力、郵便番号による住所入力をタスクペインから行う。
 */

interface ZipCloudResult {
  address1: string;
  address2: string;
  address3: string;
  zipcode: string;
}

interface ZipCloudResponse {
  message: string | null;
  results: ZipCloudResult[] | null;
  status: number;
}

interface PostalTarget {
  worksheetName: string;
  rowIndex: number;
  columnIndex: number;
  postalCode: string;
  currentAddress: string;
}

interface SheetNavigationItem {
  name: string;
  visibility: Excel.SheetVisibility;
  hasCharts: boolean;
}

type PostalStatusType = "normal" | "success" | "warning" | "error";
type PostalCodeFormat = "hyphen" | "plain";

const POSTAL_FORMAT_STORAGE_KEY: string =
  "postalCodeFormat";

let postalTarget: PostalTarget | null = null;
let sheetNavigationItems: SheetNavigationItem[] = [];
let activeSheetName: string = "";
let postalCodeFormatPreference: PostalCodeFormat =
  "hyphen";

Office.onReady((officeInfo) => {
  if (officeInfo.host !== Office.HostType.Excel) {
    return;
  }

  initializeTabs();
  initializeSheetNavigator();
  initializeCalendar();
  initializePostalSearch();

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

  const postalTab: HTMLElement | null =
    document.getElementById("tab-postal");

  sheetTab?.addEventListener("click", () => {
    switchPanel("sheet");
    void loadSheets();
  });

  calendarTab?.addEventListener("click", () => {
    switchPanel("calendar");
  });

  postalTab?.addEventListener("click", () => {
    switchPanel("postal");
  });
}

/**
 * シートナビの操作を初期化する。
 */
function initializeSheetNavigator(): void {
  const refreshButton: HTMLElement | null =
    document.getElementById("refresh-sheets");

  const searchElement: HTMLElement | null =
    document.getElementById("sheet-search");

  refreshButton?.addEventListener("click", () => {
    void loadSheets();
  });

  searchElement?.addEventListener("input", () => {
    displaySheets(sheetNavigationItems, activeSheetName);
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
 * 郵便番号検索の操作を初期化する。
 */
function initializePostalSearch(): void {
  const searchButton: HTMLElement | null =
    document.getElementById("search-postal");

  const verifyButton: HTMLElement | null =
    document.getElementById("verify-postal");

  const formatElement: HTMLElement | null =
    document.getElementById("postal-format");

  searchButton?.addEventListener("click", () => {
    void searchPostalAddress();
  });

  verifyButton?.addEventListener("click", () => {
    void verifyPostalAddress();
  });

  if (formatElement instanceof HTMLSelectElement) {
    postalCodeFormatPreference =
      loadPostalCodeFormat();

    formatElement.value =
      postalCodeFormatPreference;

    formatElement.addEventListener("change", () => {
      const postalCodeFormat: PostalCodeFormat =
        formatElement.value === "plain"
          ? "plain"
          : "hyphen";

      postalCodeFormatPreference =
        postalCodeFormat;

      savePostalCodeFormat(postalCodeFormat);
    });
  }
}

/**
 * 表示する機能を切り替える。
 */
function switchPanel(
  panelName: "sheet" | "calendar" | "postal"
): void {
  const panelNames: string[] = [
    "sheet",
    "calendar",
    "postal",
  ];

  panelNames.forEach((name: string) => {
    const panel: HTMLElement | null =
      document.getElementById(`${name}-panel`);

    const tab: HTMLElement | null =
      document.getElementById(`tab-${name}`);

    const selected: boolean = name === panelName;

    panel?.classList.toggle("hidden", !selected);
    tab?.classList.toggle("active", selected);
    tab?.setAttribute("aria-selected", String(selected));
  });

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

      worksheets.load("items/name,items/visibility");
      activeWorksheet.load("name");

      await context.sync();

      worksheets.items.forEach(
        (worksheet: Excel.Worksheet) => {
          worksheet.charts.load("count");
        }
      );

      await context.sync();

      sheetNavigationItems = worksheets.items.map(
        (worksheet: Excel.Worksheet): SheetNavigationItem => ({
          name: worksheet.name,
          visibility: worksheet.visibility,
          hasCharts: worksheet.charts.count > 0,
        })
      );

      activeSheetName = activeWorksheet.name;

      displaySheets(
        sheetNavigationItems,
        activeSheetName
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
  worksheets: SheetNavigationItem[],
  currentSheetName: string
): void {
  const sheetList: HTMLElement | null =
    document.getElementById("sheet-list");

  const searchElement: HTMLElement | null =
    document.getElementById("sheet-search");

  if (sheetList === null) {
    return;
  }

  const searchText: string =
    searchElement instanceof HTMLInputElement
      ? searchElement.value.trim().toLocaleLowerCase()
      : "";

  const filteredWorksheets: SheetNavigationItem[] =
    worksheets.filter((worksheet: SheetNavigationItem) =>
      worksheet.name.toLocaleLowerCase().includes(searchText)
    );

  sheetList.replaceChildren();

  if (filteredWorksheets.length === 0) {
    const emptyMessage: HTMLParagraphElement =
      document.createElement("p");

    emptyMessage.className = "sheet-empty";
    emptyMessage.textContent =
      "該当するシートがありません。";

    sheetList.appendChild(emptyMessage);
    return;
  }

  filteredWorksheets.forEach(
    (worksheet: SheetNavigationItem) => {
      const sheetButton: HTMLButtonElement =
        document.createElement("button");

      const isVisible: boolean =
        worksheet.visibility ===
        Excel.SheetVisibility.visible;

      sheetButton.type = "button";
      sheetButton.className = "sheet-item";
      sheetButton.disabled = !isVisible;

      if (worksheet.hasCharts) {
        const chartIcon: HTMLSpanElement =
          document.createElement("span");

        chartIcon.className = "sheet-icon";
        chartIcon.textContent = "📊";
        chartIcon.setAttribute("aria-hidden", "true");

        sheetButton.appendChild(chartIcon);
      }

      const sheetName: HTMLSpanElement =
        document.createElement("span");

      sheetName.className = "sheet-name";
      sheetName.textContent = worksheet.name;

      sheetButton.appendChild(sheetName);

      if (!isVisible) {
        const hiddenLabel: HTMLSpanElement =
          document.createElement("span");

        hiddenLabel.className = "sheet-visibility";
        hiddenLabel.textContent = "非表示";

        sheetButton.classList.add("hidden-sheet");
        sheetButton.appendChild(hiddenLabel);
      }

      if (
        isVisible &&
        worksheet.name === currentSheetName
      ) {
        sheetButton.classList.add("active");
      }

      if (isVisible) {
        sheetButton.addEventListener("click", () => {
          void activateSheet(worksheet.name);
        });
      }

      sheetList.appendChild(sheetButton);
    }
  );
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

/**
 * 選択セルの郵便番号から住所候補を検索する。
 */
async function searchPostalAddress(): Promise<void> {
  clearPostalResults();
  setPostalStatus("郵便番号を確認しています。");

  try {
    postalTarget = await readPostalTarget();

    const response: ZipCloudResponse =
      await requestPostalAddress(postalTarget.postalCode);

    const results: ZipCloudResult[] =
      getPostalResults(response);

    displayPostalCandidates(results, false);

    if (postalTarget.currentAddress !== "") {
      setPostalStatus(
        "右隣のセルには住所が入力されています。候補を選ぶと上書き確認を表示します。",
        "warning"
      );
      return;
    }

    setPostalStatus("住所候補を選択してください。");
  } catch (error: unknown) {
    handlePostalError(error);
  }
}

/**
 * 入力済み住所が郵便番号の住所候補と一致するか確認する。
 */
async function verifyPostalAddress(): Promise<void> {
  clearPostalResults();
  setPostalStatus("郵便番号と住所を確認しています。");

  try {
    postalTarget = await readPostalTarget();

    if (postalTarget.currentAddress === "") {
      throw new Error("右隣の住所セルが空欄です。");
    }

    const response: ZipCloudResponse =
      await requestPostalAddress(postalTarget.postalCode);

    const results: ZipCloudResult[] =
      getPostalResults(response);

    const currentAddress: string =
      normalizeAddress(postalTarget.currentAddress);

    const matched: boolean = results.some(
      (result: ZipCloudResult) =>
        currentAddress.startsWith(
          normalizeAddress(buildAddress(result))
        )
    );

    if (matched) {
      setPostalStatus(
        "✓ 郵便番号と住所は一致しています。",
        "success"
      );
      return;
    }

    setPostalStatus(
      "郵便番号と住所が一致しません。修正する場合は住所候補を選択してください。",
      "error"
    );

    displayPostalCandidates(results, true);
  } catch (error: unknown) {
    handlePostalError(error);
  }
}

/**
 * 選択セルと右隣セルの情報を取得する。
 */
async function readPostalTarget(): Promise<PostalTarget> {
  return Excel.run(
    async (
      context: Excel.RequestContext
    ): Promise<PostalTarget> => {
      const selectedRange: Excel.Range =
        context.workbook.getSelectedRange();

      const postalCell: Excel.Range =
        selectedRange.getCell(0, 0);

      const addressCell: Excel.Range =
        postalCell.getOffsetRange(0, 1);

      const worksheet: Excel.Worksheet =
        postalCell.worksheet;

      postalCell.load("text,rowIndex,columnIndex");
      addressCell.load("text");
      worksheet.load("name");

      await context.sync();

      const postalCode: string =
        normalizePostalCode(postalCell.text[0][0]);

      if (postalCode.length !== 7) {
        throw new Error(
          "選択セルの郵便番号は7桁で入力してください。"
        );
      }

      return {
        worksheetName: worksheet.name,
        rowIndex: postalCell.rowIndex,
        columnIndex: postalCell.columnIndex,
        postalCode,
        currentAddress: String(
          addressCell.text[0][0] ?? ""
        ).trim(),
      };
    }
  );
}

/**
 * ZipCloudをJSONPで呼び出す。
 * 認証情報を持たず、ブラウザーのCORS制限を受けないために使用する。
 */
function requestPostalAddress(
  postalCode: string
): Promise<ZipCloudResponse> {
  return new Promise(
    (
      resolve: (value: ZipCloudResponse) => void,
      reject: (reason?: unknown) => void
    ) => {
      const callbackName: string =
        `zipCloudCallback_${Date.now()}_${Math.floor(
          Math.random() * 100000
        )}`;

      const script: HTMLScriptElement =
        document.createElement("script");

      const callbackContainer: Record<string, unknown> =
        window as unknown as Record<string, unknown>;

      const timeoutId: number = window.setTimeout(() => {
        cleanup();
        reject(
          new Error(
            "郵便番号検索がタイムアウトしました。"
          )
        );
      }, 10000);

      function cleanup(): void {
        window.clearTimeout(timeoutId);
        script.remove();
        delete callbackContainer[callbackName];
      }

      callbackContainer[callbackName] = (
        response: ZipCloudResponse
      ): void => {
        cleanup();
        resolve(response);
      };

      script.onerror = (): void => {
        cleanup();
        reject(
          new Error(
            "郵便番号検索サービスへ接続できませんでした。"
          )
        );
      };

      script.src =
        "https://zipcloud.ibsnet.co.jp/api/search" +
        `?zipcode=${encodeURIComponent(postalCode)}` +
        `&callback=${encodeURIComponent(callbackName)}`;

      document.body.appendChild(script);
    }
  );
}

/**
 * APIレスポンスから利用できる住所候補を取得する。
 */
function getPostalResults(
  response: ZipCloudResponse
): ZipCloudResult[] {
  if (response.status !== 200) {
    throw new Error(
      response.message ??
        "郵便番号を検索できませんでした。"
    );
  }

  if (
    response.results === null ||
    response.results.length === 0
  ) {
    throw new Error(
      "該当する住所が見つかりませんでした。"
    );
  }

  return response.results;
}

/**
 * 住所候補をボタンとして表示する。
 */
function displayPostalCandidates(
  results: ZipCloudResult[],
  overwrite: boolean
): void {
  const resultContainer: HTMLElement | null =
    document.getElementById("postal-results");

  if (resultContainer === null) {
    return;
  }

  resultContainer.replaceChildren();

  results.forEach((result: ZipCloudResult) => {
    const address: string = buildAddress(result);

    const button: HTMLButtonElement =
      document.createElement("button");

    button.type = "button";
    button.className = overwrite
      ? "postal-candidate overwrite"
      : "postal-candidate";

    button.textContent = overwrite
      ? `上書き：${address}`
      : address;

    button.addEventListener("click", () => {
      void writePostalAddress(address, overwrite);
    });

    resultContainer.appendChild(button);
  });
}

/**
 * 郵便番号と住所を対象セルへ書き込む。
 */
async function writePostalAddress(
  address: string,
  overwrite: boolean
): Promise<void> {
  if (postalTarget === null) {
    setPostalStatus(
      "先に郵便番号を検索してください。",
      "error"
    );
    return;
  }

  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const worksheet: Excel.Worksheet =
        context.workbook.worksheets.getItem(
          postalTarget!.worksheetName
        );

      const postalCell: Excel.Range =
        worksheet.getCell(
          postalTarget!.rowIndex,
          postalTarget!.columnIndex
        );

      const addressCell: Excel.Range =
        worksheet.getCell(
          postalTarget!.rowIndex,
          postalTarget!.columnIndex + 1
        );

      addressCell.load("text");
      await context.sync();

      const existingAddress: string =
        String(addressCell.text[0][0] ?? "").trim();

      if (existingAddress !== "" && !overwrite) {
        setPostalStatus(
          "右隣のセルには住所があります。住所確認を行ってから上書きしてください。",
          "warning"
        );
        return;
      }

      postalCell.numberFormat = [["@"]];
      postalCell.values = [[
        formatPostalCode(postalTarget!.postalCode),
      ]];

      addressCell.values = [[address]];

      await context.sync();
    });

    postalTarget.currentAddress = address;
    clearPostalResults();
    setPostalStatus(
      "✓ 郵便番号と住所を入力しました。",
      "success"
    );
  } catch (error: unknown) {
    handlePostalError(error);
  }
}

/**
 * 全角数字や記号を除去して7桁の郵便番号へ整形する。
 */
function normalizePostalCode(value: string): string {
  const halfWidthValue: string = value.replace(
    /[０-９]/g,
    (character: string) =>
      String.fromCharCode(
        character.charCodeAt(0) - 0xfee0
      )
  );

  return halfWidthValue.replace(/[^0-9]/g, "");
}

/**
 * 選択された保存形式で郵便番号を整形する。
 */
function formatPostalCode(postalCode: string): string {
  if (postalCodeFormatPreference === "plain") {
    return postalCode;
  }

  return (
    postalCode.slice(0, 3) +
    "-" +
    postalCode.slice(3)
  );
}

/**
 * 保存済みの郵便番号形式を取得する。
 */
function loadPostalCodeFormat(): PostalCodeFormat {
  try {
    const storedFormat: string | null =
      window.localStorage.getItem(
        POSTAL_FORMAT_STORAGE_KEY
      );

    return storedFormat === "plain"
      ? "plain"
      : "hyphen";
  } catch {
    return "hyphen";
  }
}

/**
 * 郵便番号形式を端末内へ保存する。
 */
function savePostalCodeFormat(
  postalCodeFormat: PostalCodeFormat
): void {
  try {
    window.localStorage.setItem(
      POSTAL_FORMAT_STORAGE_KEY,
      postalCodeFormat
    );
  } catch {
    setPostalStatus(
      "保存形式を記憶できませんでした。この画面では選択した形式を使用します。",
      "warning"
    );
  }
}

/**
 * APIの住所項目を1つの住所へ結合する。
 */
function buildAddress(result: ZipCloudResult): string {
  return (
    result.address1 +
    result.address2 +
    result.address3
  );
}

/**
 * 比較に影響しない空白を除去する。
 */
function normalizeAddress(address: string): string {
  return address.replace(/[\s　]/g, "");
}

/**
 * 郵便番号検索の結果表示を初期化する。
 */
function clearPostalResults(): void {
  const resultContainer: HTMLElement | null =
    document.getElementById("postal-results");

  resultContainer?.replaceChildren();
}

/**
 * 郵便番号検索の状態を表示する。
 */
function setPostalStatus(
  message: string,
  statusType: PostalStatusType = "normal"
): void {
  const statusElement: HTMLElement | null =
    document.getElementById("postal-status");

  if (statusElement === null) {
    return;
  }

  statusElement.textContent = message;
  statusElement.className =
    `postal-status ${statusType}`;
}

/**
 * 郵便番号検索のエラーを表示する。
 */
function handlePostalError(error: unknown): void {
  console.error(error);

  if (error instanceof Error) {
    setPostalStatus(error.message, "error");
    return;
  }

  setPostalStatus(
    "郵便番号検索中にエラーが発生しました。",
    "error"
  );
}

