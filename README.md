# Excel VBA マクロ集：セル表示モード切り替え ＆ Enterキー移動方向切り替え

このリポジトリでは、日常のExcel操作をより快適にするための2つのVBAマクロを紹介します。

1. **セルの表示モードを切り替えるマクロ**  
　→ 「通常表示」「縮小して全体を表示」「折り返して全体を表示」を順番にトグル  
2. **Enterキーの移動方向を切り替えるマクロ**  
　→ 「下方向」⇄「右方向」を切り替え  

どちらのマクロも **「個人用マクロブック（PERSONAL.XLSB）」** に登録しておくことで、  
すべてのExcelファイルで共通して使えるようになります。

---

## 🧩 ① セル表示モードを切り替えるマクロ

### 機能概要
選択中のセルの表示設定を以下の3段階で順番に切り替えます。

1. **通常表示（初期状態）**  
2. **縮小して全体を表示（Shrink to Fit）**  
3. **折り返して全体を表示（Wrap Text）**

実行のたびに上記の順でループし、現在の状態がステータスバーに3秒間表示されます。

---

### コード

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

## 🧭 ② Enterキーの移動方向を切り替えるマクロ

### 機能概要
Excelで Enter キーを押したときのカーソル移動方向を
「下」⇄「右」 で切り替えます。

入力フォームや名簿作成など、横方向に入力したい場合に便利です。

---

### コード

Sub ToggleEnterKeyDirection()
    Application.MoveAfterReturn = True
    If Application.MoveAfterReturnDirection = xlToRight Then
        Application.MoveAfterReturnDirection = xlDown
    Else
        Application.MoveAfterReturnDirection = xlToRight
    End If
End Sub

💾 「個人用マクロブック（PERSONAL.XLSB）」に登録する手順（推奨）

これらのマクロを「個人用マクロブック（Personal.xlsb）」に保存しておくことで、
あなたのPC上のすべてのExcelファイルで利用可能になります。

手順

Excelを開く

「開発」タブ → ［マクロの記録］ をクリック

「マクロの保存先」で
　➡ 「個人用マクロブック（Personal Macro Workbook）」 を選択

適当に操作してすぐに記録を停止
　→ Personal.xlsb ファイルが自動的に作成されます

Alt + F11 でVBAエディタを開く

左の「プロジェクト」ウィンドウで
　📂 VBAProject (PERSONAL.XLSB) を選択

「挿入」→「標準モジュール」をクリック

このREADMEにあるマクロコードを貼り付け

保存（Ctrl + S）して閉じる

Excelを再起動

⌨️ ショートカットキー設定（任意）

Excelで「開発」タブ →「マクロ」

対象のマクロを選択 →「オプション」

例：

ToggleTextDisplayMode → Ctrl + Shift + T

ToggleEnterKeyDirection → Ctrl + Shift + E

→ 以後、どのブックでもショートカット1発で実行可能です🎯

🧠 注意点

これらのマクロは安全な操作のみを行いますが、初めてマクロを有効化する場合は「マクロを有効にする」設定が必要です。

Application.MoveAfterReturnDirection は Excel全体の設定 を変更します。
切り替え後は、他のブックでも同じ方向になります。

Personal.xlsb はExcel起動時に自動で読み込まれます。削除しないようご注意ください。
