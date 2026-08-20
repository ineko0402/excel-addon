export type PostalCodeFormat = "hyphen" | "plain";

/**
 * 郵便番号を半角数字7桁へ正規化する。
 */
export function normalizePostalCode(value: string): string {
  return value
    .replace(/[０-９]/g, (character: string) =>
      String.fromCharCode(character.charCodeAt(0) - 0xfee0)
    )
    .replace(/[^0-9]/g, "");
}

/**
 * 郵便番号を選択した文字列形式へ整える。
 */
export function formatPostalCode(
  postalCode: string,
  format: PostalCodeFormat
): string {
  return format === "plain"
    ? postalCode
    : `${postalCode.slice(0, 3)}-${postalCode.slice(3)}`;
}

/**
 * 住所比較用に空白を除去する。
 */
export function normalizeAddress(address: string): string {
  return address.replace(/[\s　]/g, "");
}

/**
 * yyyy-mm-ddをExcelの日付シリアル値へ変換する。
 */
export function convertToExcelDate(dateText: string): number {
  const [year, month, day]: number[] =
    dateText.split("-").map(Number);

  return Math.floor(
    (Date.UTC(year, month - 1, day) -
      Date.UTC(1899, 11, 30)) /
      86400000
  );
}

/**
 * Excelのシート名として使用できるか検証する。
 */
export function getSheetNameValidationError(
  sheetName: string
): string | null {
  const trimmedSheetName: string = sheetName.trim();

  if (trimmedSheetName === "") {
    return "シート名を入力してください。";
  }

  if (trimmedSheetName.length > 31) {
    return "シート名は31文字以内で入力してください。";
  }

  if (/[:\\/?*\[\]]/.test(trimmedSheetName)) {
    return "シート名に : \\ / ? * [ ] は使用できません。";
  }

  return null;
}
