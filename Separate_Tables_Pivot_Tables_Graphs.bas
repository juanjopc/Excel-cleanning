Attribute VB_Name = "Módulo1"
Sub MoverElementosAHojasSeparadas()
    Dim ws As Worksheet
    Dim ch As ChartObject
    Dim pt As PivotTable
    Dim lo As ListObject
    Dim newWS As Worksheet
    Dim baseName As String
    Dim newName As String
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    ' Mover gráficos
    For Each ws In ThisWorkbook.Worksheets
        i = 1
        For Each ch In ws.ChartObjects
            baseName = Left(ws.Name, 26) & "_G"
            newName = GetUniqueName(baseName, i)
            Set newWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            newWS.Name = newName
            
            ch.Cut
            newWS.Paste
            i = i + 1
        Next ch
    Next ws
    
    ' Mover tablas dinámicas
    For Each ws In ThisWorkbook.Worksheets
        i = 1
        For Each pt In ws.PivotTables
            baseName = Left(ws.Name, 26) & "_TD"
            newName = GetUniqueName(baseName, i)
            Set newWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            newWS.Name = newName
            pt.TableRange2.Cut Destination:=newWS.Range("A1")
            i = i + 1
        Next pt
    Next ws
    
    ' Mover tablas
    For Each ws In ThisWorkbook.Worksheets
        i = 1
        For Each lo In ws.ListObjects
            baseName = Left(ws.Name, 26) & "_T"
            newName = GetUniqueName(baseName, i)
            Set newWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            newWS.Name = newName
            lo.Range.Cut Destination:=newWS.Range("A1")
            i = i + 1
        Next lo
    Next ws
    
    Application.ScreenUpdating = True
    MsgBox "Proceso completado. Se han movido todos los elementos a hojas separadas."
End Sub

Function GetUniqueName(baseName As String, i As Integer) As String
    Dim newName As String
    newName = baseName & i
    If Len(newName) > 31 Then
        newName = Left(baseName, 31 - Len(CStr(i))) & i
    End If
    GetUniqueName = newName
End Function
