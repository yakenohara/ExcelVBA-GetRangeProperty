Attribute VB_Name = "GetRangeProperty"
'
' �\���e�L�X�g��Ԃ�
' (�\���`�������f���ꂽ��Ԃ̃e�L�X�g��Ԃ�)
' ex:
' `500` -> `\500``�ƂȂ�悤�ɕ\���`����ݒ肵�Ă���ꍇ�A
' `\500`��Ԃ�
'
Public Function getText(ByVal rng As Range) As Variant
    getText = rng.Text
    
End Function

'
' �������e�L�X�g�Ƃ��ĕԂ�
'
Public Function getFormula(ByVal rng As Range) As Variant
    getFormula = rng.Formula
    
End Function

'
'�w��Z�����Z����������Ă����ꍇ�A
'����̓��e��Ԃ�
'
Public Function getMergedValue(ByVal target As Range) As Variant
    getMergedValue = target(1, 1).MergeArea.Cells(1, 1).Value
    
End Function

'
' �Z���ɑ}�����ꂽ�n�C�p�[�����N��Ԃ�
'
Public Function getHyperlinkAddress(ByVal rng As Range)
    getHyperlinkAddress = rng.Hyperlinks(1).Address
End Function

