Sub Macro1()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("huhu")).Select
    Selection.ShapeRange.LockAspectRatio = msoFalse
    Selection.ShapeRange.ScaleWidth 1.4423076923, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 2.5089445438, msoFalse, msoScaleFromTopLeft
End Sub