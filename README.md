# ナナイロ作業報告入力ツール

ナナイロ様へ毎月提出する作業報告書を、日別入力・月間工数の確認・Excel/PDF出力まで一画面で作成するWindowsデスクトップアプリです。

## 主な機能

- 対象月の日付、曜日、土日祝日を自動生成
- 通常勤務、公休日、有給休暇の区分に応じた入力と月間集計
- 当日までの実績工数、月末見込み、精算幅、有給取得目安を表示
- 複数日への勤務区分・作業内容・公休日・標準時間の一括設定
- エラー／警告チェックと該当日への移動
- UTF-8 JSON形式の `.nwr` 下書き保存、自動復旧、最近使ったファイル
- マクロなし `.xlsx` の出力・再読込
- デスクトップ版Microsoft Excelを利用した1ページPDF出力
- 内閣府の祝日CSVを同梱し、手動・バックグラウンド更新に対応

## 必要環境

- Windows 11 x64
- PDF出力時のみデスクトップ版Microsoft Excelが必要
- 配布版は自己完結型のため、利用者による.NETのインストールは不要

## 開発

```powershell
dotnet test NanairoWorkReportTool.slnx -c Release
dotnet publish src\NanairoWorkReportTool\NanairoWorkReportTool.csproj -c Release -r win-x64 --self-contained true -o publish\win-x64
```

技術構成は .NET 10 / WPF / MVVM / Open XML です。帳票テンプレートは実帳票の外観・印刷設定を引き継いだマクロなし形式を同梱しています。
