##　開発環境の構築
`npm install`

## sheet.xlsxを準備
・今回使用するsheet.xlsxをxlsx/配下に設置。
・前回使用したものは、ファイル名を変更。
例：sheet_2025.xlsx

## main.jsファイルの値を更新
main.jsファイル修正箇所
・10行目：sheet_name変数の値を今回使用するsheet.xlsxファイル内のシート名に設定。
・17行目：EXCEL_END_ROWの値をsheet.xlsxファイルのデータ終了行値を設定。

## jsonデータの出力
`node main.js`

dest配下に以下のファイルが出力されます。
list.json（トップページ用）
list-lp.json（下層ページ用）