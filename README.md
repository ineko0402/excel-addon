# Excelユーティリティ

Excelの操作を補助するOffice Add-inです。TypeScriptとOffice JavaScript APIで作成しています。

## 機能

- シート一覧の表示
- シートをクリックして移動
- アクティブシートの強調表示
- HTML標準カレンダーによる日付選択
- 選択範囲の左上セルへExcel日付を入力

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
