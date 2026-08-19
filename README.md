# Excelユーティリティ

Excelの操作を補助するOffice Add-inです。TypeScriptとOffice JavaScript APIで作成しています。

## 機能

- シート一覧のコンパクト表示とシート名検索
- シートをクリックして移動
- アクティブシートの強調表示
- 非表示シートのグレー表示
- シート設定メニューから表示・非表示を切り替え
- シート設定メニューからシート名を変更
- グラフを含むシートへの📊アイコン表示
- HTML標準カレンダーによる日付選択
- 選択範囲の左上セルへExcel日付を入力
- 日付の表示形式を選択して端末内へ保存
- 選択セルへTODAY関数を入力
- 選択した郵便番号セルから住所候補を検索
- 右隣の住所セルへ候補を入力
- 入力済みの郵便番号と住所を照合

## 日付入力

カレンダーで選択した日付をExcelの日付値として入力します。文字列ではないため、日付計算・並べ替え・集計に利用できます。

- `yyyy/mm/dd`、`yyyy/m/d`、`yyyy年m月d日`、`yyyy-mm-dd`、`m/d`から表示形式を選択
- 「既存のセル書式を優先」では、標準書式のセルだけ`yyyy/mm/dd`へ補正
- 選択した表示形式は端末内に保存し、次回起動時も維持
- 「TODAY関数を入力」は`=TODAY()`を設定し、再計算時に当日の日付へ更新
- 複数セルを選択している場合は左上セルだけに入力

## 郵便番号検索

郵便番号セルを選択し、右隣のセルを住所として検索・入力・照合します。郵便番号検索には[ZipCloud郵便番号検索API](https://zipcloud.ibsnet.co.jp/doc/api)を利用します。

- 郵便番号はハイフンあり・なし、全角・半角に対応
- 郵便番号は文字列の`123-4567`または`1234567`形式を選択して保存
- 既存住所は自動で上書きしない
- 入力済み住所は、郵便番号の住所候補との先頭一致で確認
- 番地や建物名が続く住所も一致として扱う
- 選択した保存形式は端末内に保存し、次回起動時も維持

## 必要環境

- Node.js 22
- Excelデスクトップ版
- Git
- Visual Studio Code（推奨）

## ローカル起動

リポジトリを取得し、依存関係をインストールします。

```powershell
git clone https://github.com/ineko0402/excel-addon.git
cd excel-addon
npm install
```

Excelでアドインを起動します。

```powershell
npm start
```

終了するときは次を実行します。

```powershell
npm stop
```

## 検査

```powershell
npm run typecheck
npm run validate
npm run build
```

## 公開

`main`ブランチへのpushを契機に、GitHub Actionsが型検査、マニフェスト検証、Webpackビルドを実行し、`dist`をGitHub Pagesへ公開します。

- 公開ページ: https://ineko0402.github.io/excel-addon/
- 公開用マニフェスト: https://ineko0402.github.io/excel-addon/manifest.xml

ローカル開発では`https://localhost:3000/`を使用し、本番ビルド時だけGitHub PagesのURLへ変換します。

## ライセンス

MIT License
