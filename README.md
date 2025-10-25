# Excel VBA：セルの表示モードを切り替えるマクロ

このVBAマクロは、選択中のセルの表示モードを以下の3段階で順番に切り替えます。

1. **通常表示（初期状態）**  
2. **縮小して全体を表示（Shrink to Fit）**  
3. **折り返して全体を表示（Wrap Text）**

実行するたびに上記の順でモードがループし、現在の状態がExcel画面左下のステータスバーに表示されます。

---

## 🧩 特徴

- 実行のたびに  
  **通常 → 縮小 → 折り返し → 通常**  
  の順で自動切り替え  
- 選択中の複数セルにも対応  
- 現在の状態をステータスバーに表示（3秒後に自動クリア）  
- メッセージボックス等は表示されず、操作がスムーズ  

---

## 💻 コード

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
