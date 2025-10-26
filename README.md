# 💡 Excel VBA マクロ集  
## セル表示モード切り替え ＆ Enterキー移動方向切り替え

このリポジトリでは、日常の Excel 操作をより快適にするための  
2つの便利な VBA マクロを紹介します。

---

### 🧩 収録マクロ一覧

1. **セルの表示モードを切り替えるマクロ**  
　→ 「通常表示」⇄「縮小して全体を表示」⇄「折り返して全体を表示」を順番にトグル  
2. **Enterキーの移動方向を切り替えるマクロ**  
　→ 「下方向」⇄「右方向」をワンタッチで切り替え  

どちらのマクロも **「個人用マクロブック（PERSONAL.XLSB）」** に登録することで、  
あなたの PC 上のすべての Excel ファイルで共通して使えるようになります。

---

## 🧭 ① セル表示モードを切り替えるマクロ

### 🔍 機能概要
選択中のセルの表示設定を、以下の3段階で順番に切り替えます。

| 実行回数 | 表示モード | 説明 |
|:--:|:--|:--|
| 1回目 | 縮小して全体を表示 | 文字を小さくしてセル内に収める |
| 2回目 | 折り返して全体を表示 | 複数行に折り返して表示 |
| 3回目 | 通常表示（初期状態） | どちらもオフの状態に戻る |

実行のたびにループし、現在の状態がステータスバーに3秒間表示されます。

---

### 💻 コード

```vba
Sub ToggleTextDisplayMode()
    Dim c As Range
    Dim mode As String
    
    On Error Resume Next
    
    For Each c In Selection
        ' 現在の状態を確認
        If Not c.ShrinkToFit And Not c.WrapText Then
            ' 初期 → 縮小
            c.ShrinkToFit = True
            c.WrapText = False
            mode = "縮小して全体を表示"
            
        ElseIf c.ShrinkToFit And Not c.WrapText Then
            ' 縮小 → 折り返し
            c.ShrinkToFit = False
            c.WrapText = True
            mode = "折り返して全体を表示"
            
        Else
            ' 折り返し → 初期（どちらもOFF）
            c.ShrinkToFit = False
            c.WrapText = False
            mode = "通常表示（初期状態）"
        End If
    Next c
    
    On Error GoTo 0
    
    ' ステータスバーに表示
    Application.StatusBar = "セル表示モード：" & mode
    Application.OnTime Now + TimeValue("00:00:03"), "ClearStatusBar"
End Sub

Sub ClearStatusBar()
    Application.StatusBar = False
End Sub
```
## ⌨️ ② Enterキーの移動方向を切り替えるマクロ

### 🔍 機能概要
Excelで Enter キーを押したときのカーソル移動方向を、  
「下方向」⇄「右方向」で切り替えます。  

入力フォームや名簿など、横方向に入力したい場面で便利です。

---

### 💻 コード

```vba
Sub ToggleEnterKeyDirection()
    Application.MoveAfterReturn = True
    If Application.MoveAfterReturnDirection = xlToRight Then
        Application.MoveAfterReturnDirection = xlDown
    Else
        Application.MoveAfterReturnDirection = xlToRight
    End If
End Sub
```

## 💾 「個人用マクロブック（PERSONAL.XLSB）」に登録する手順（推奨）

これらのマクロを 「個人用マクロブック（Personal.xlsb）」 に保存しておくと、
どのExcelファイルを開いても利用できるようになります。

### 手順

Excelを開く

「開発」タブ → ［マクロの記録］ をクリック

「マクロの保存先」で
　➡ 「個人用マクロブック（Personal Macro Workbook）」 を選択

適当に操作して記録をすぐに停止（これで Personal.xlsb が自動作成されます）

Alt + F11 で VBAエディタを開く

左の「プロジェクト」ウィンドウで
　📂 VBAProject (PERSONAL.XLSB) を選択

「挿入」→「標準モジュール」をクリック

このREADME内のマクロコードを貼り付け

Ctrl + S で保存 → Excelを再起動

## ⚙️ ショートカットキー設定（任意）

Excelで「開発」タブ →「マクロ」

対象のマクロを選択 →「オプション」クリック

任意のショートカットを設定

例：

マクロ名	推奨ショートカット
ToggleTextDisplayMode	Ctrl + Shift + T
ToggleEnterKeyDirection	Ctrl + Shift + E

→ 以後、どのブックでもショートカット1発で実行可能になります 🎯

## ⚠️ 注意点

これらのマクロは安全な操作のみを行いますが、
　初めて使用する際は「マクロを有効にする」設定が必要です。

Application.MoveAfterReturnDirection は Excel全体に影響 します。
　切り替え後は他のブックでも同じ方向になります。

Personal.xlsb はExcel起動時に自動で読み込まれます。
　誤って削除しないようご注意ください。
