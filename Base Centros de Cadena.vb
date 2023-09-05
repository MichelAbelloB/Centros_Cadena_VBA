Sub Eliminar()

    Range("B10").ClearContents
    
End Sub

Sub Eliminar_Tabla()

    Range("A9:E28").ClearContents
    
End Sub

Sub Contar_Proveedor()

    Dim Contador As Long
    Dim Celda As Range
    
    Contador = 0
    
    For Each Celda In ThisWorkbook.Worksheets("BD").Range("D4:D50003")
        If Celda.Value = "ÉXITO" Then
            Contador = Contador + 1
        End If
    Next Celda
    
    ThisWorkbook.Worksheets("Punto1").Range("B10").Value = Contador
    
End Sub

Sub Promedio_Pizza()

    Dim Contador As Long
    Dim Celda As Range
    Dim Suma As Double
    Dim Promedio As Double

    Contador = 0
    Suma = 0

    For Each Celda In ThisWorkbook.Worksheets("BD").Range("C4:C50003")
        If Celda.Value = "PIZZA" Then
        Contador = Contador + 1
            Suma = Suma + Celda.Offset(0, 5).Value
        End If
    Next Celda

    If Contador > 0 Then
    Promedio = Suma / Contador
    Else
    Promedio = 0
    End If
    
    ThisWorkbook.Worksheets("Punto2").Range("B10").Value = Promedio
    
End Sub

Sub Fecha_Mas_Antigua()

    Dim FechaMasAntigua As Date
    
    FechaMasAntigua = WorksheetFunction.Min(Sheets("BD").Range("B:B"))
    Sheets("Punto3").Range("B10").Value = FechaMasAntigua
    
End Sub

Sub Tabla_Dinamica()

    Dim datos As Range
    Dim destino As Range
    Set datos = ThisWorkbook.Sheets("BD").Range("A1:J50001")
    Set destino = ThisWorkbook.Sheets("Punto4").Range("B10")
    
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "BD!R1C1:R50001C10", Version:=8).CreatePivotTable TableDestination:= _
        "Punto4!R10C2", TableName:="TablaDinámica3", DefaultVersion:=8
    Sheets("Punto4").Select
    Cells(10, 2).Select
    
    With ActiveSheet.PivotTables("TablaDinámica3").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("TablaDinámica3").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("CIUDAD DE VENTA")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("TablaDinámica3").AddDataField ActiveSheet.PivotTables( _
        "TablaDinámica3").PivotFields("TOTAL"), "Suma de TOTAL", xlSum
    With ActiveSheet.PivotTables("TablaDinámica3").PivotFields("Suma de TOTAL")
        .Caption = "Máx. de TOTAL"
        .Function = xlMax
    End With
    Range("C11:C16").Select
    Selection.Style = "Currency"
    
    
End Sub


Sub Tabla_Dinamica_Proveedor()

    Dim datos As Range
    Dim destino As Range
    Set datos = ThisWorkbook.Sheets("BD").Range("A1:J50001")
    Set destino = ThisWorkbook.Sheets("Punto5").Range("B10")
    
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Punto4").PivotTables("TablaDinámica3").PivotCache. _
        CreatePivotTable TableDestination:="Punto5!R10C2", TableName:= _
        "TablaDinámica5", DefaultVersion:=8
    Sheets("Punto5").Select
    Cells(10, 2).Select
    ActiveSheet.PivotTables("TablaDinámica5").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("TablaDinámica5").PivotFields("PROVEEDOR")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("TablaDinámica5").AddDataField ActiveSheet.PivotTables( _
        "TablaDinámica5").PivotFields("TOTAL"), "Suma de TOTAL", xlSum
    With ActiveSheet.PivotTables("TablaDinámica5").PivotFields("Suma de TOTAL")
        .Caption = "Promedio de TOTAL"
        .Function = xlAverage
    End With
    
    
End Sub


Sub Tabla_Dinamica_Pais()

    Dim datos As Range
    Dim destino As Range
    Set datos = ThisWorkbook.Sheets("BD").Range("A1:J50001")
    Set destino = ThisWorkbook.Sheets("Punto6").Range("B10")
    
    ActiveWorkbook.Worksheets("Punto4").PivotTables("TablaDinámica3").PivotCache. _
        CreatePivotTable TableDestination:="Punto6!R10C2", TableName:= _
        "TablaDinámica6", DefaultVersion:=8
    Sheets("Punto6").Select
    Cells(10, 2).Select
    
    With ActiveSheet.PivotTables("TablaDinámica6").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("TablaDinámica6").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("TablaDinámica6").PivotFields( _
        "PAÍS DE IMPORTACIÓN")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("TablaDinámica6").PivotFields( _
        "PAÍS DE IMPORTACIÓN")
        .PivotItems("NO APLICA").Visible = False
    End With
    With ActiveSheet.PivotTables("TablaDinámica6").PivotFields("IMPORTADO")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("TablaDinámica6").PivotFields("IMPORTADO")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("TablaDinámica6").AddDataField ActiveSheet.PivotTables( _
        "TablaDinámica6").PivotFields("IMPORTADO"), "Cuenta de IMPORTADO", xlCount
    With ActiveSheet.PivotTables("TablaDinámica6").PivotFields("IMPORTADO")
        .PivotItems("NO").Visible = False
    End With
    
    
End Sub
