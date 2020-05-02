Attribute VB_Name = "GetRangeProperty"
'
' 表示テキストを返す
' (表示形式が反映された状態のテキストを返す)
' ex:
' `500` -> `\500``となるように表示形式を設定している場合、
' `\500`を返す
'
Public Function getText(ByVal rng As Range) As Variant
    getText = rng.Text
    
End Function

'
' 数式をテキストとして返す
'
Public Function getFormula(ByVal rng As Range) As Variant
    getFormula = rng.Formula
    
End Function

'
'指定セルがセル結合されていた場合、
'左上の内容を返す
'
Public Function getMergedValue(ByVal target As Range) As Variant
    getMergedValue = target(1, 1).MergeArea.Cells(1, 1).Value
    
End Function

'
' セルに挿入されたハイパーリンクを返す
'
Public Function getHyperlinkAddress(ByVal rng As Range)
    getHyperlinkAddress = rng.Hyperlinks(1).Address
End Function

