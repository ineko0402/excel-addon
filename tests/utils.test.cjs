const test = require("node:test");
const assert = require("node:assert/strict");

const {
  convertToExcelDate,
  formatPostalCode,
  getSheetNameValidationError,
  normalizeAddress,
  normalizePostalCode,
} = require("../.test-dist/utils.js");

test("郵便番号を全角やハイフンから7桁へ正規化する", () => {
  assert.equal(normalizePostalCode("１２３－４５６７"), "1234567");
  assert.equal(normalizePostalCode("123 4567"), "1234567");
});

test("郵便番号を選択した形式へ整える", () => {
  assert.equal(formatPostalCode("1234567", "hyphen"), "123-4567");
  assert.equal(formatPostalCode("1234567", "plain"), "1234567");
});

test("住所比較用に半角・全角空白を除去する", () => {
  assert.equal(normalizeAddress("東京都 千代田区　丸の内"), "東京都千代田区丸の内");
});

test("日付をExcel日付シリアル値へ変換する", () => {
  assert.equal(convertToExcelDate("1970-01-01"), 25569);
});

test("シート名を検証する", () => {
  assert.equal(getSheetNameValidationError("   "), "シート名を入力してください。");
  assert.equal(
    getSheetNameValidationError("12345678901234567890123456789012"),
    "シート名は31文字以内で入力してください。"
  );
  assert.equal(
    getSheetNameValidationError("売上/集計"),
    "シート名に : \\ / ? * [ ] は使用できません。"
  );
  assert.equal(getSheetNameValidationError("売上集計"), null);
});
