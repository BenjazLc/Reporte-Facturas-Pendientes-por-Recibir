Attribute VB_Name = "M5_FINAL"
Option Explicit

Public Sub M5_Detalle_Nacionales_Extranjeros()

    Dim wb As Workbook, wsData As Worksheet
    Dim tc As Double: tc = 3.45

    Set wb = ObtenerLibroReporteFacturas()

    If wb Is Nothing Then
        MsgBox "No encontré abierto el archivo de reporte con la hoja 'Saldos por Facturar'.", vbCritical
        Exit Sub
    End If

    Set wsData = wb.Sheets("Saldos por Facturar")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next
    wb.Sheets("Nacionales").Delete
    wb.Sheets("Extranjeros").Delete
    On Error GoTo 0

    Application.DisplayAlerts = True

    CrearHojaProveedorTipo wb, wsData, "Nacionales", "1", tc
    CrearHojaProveedorTipo wb, wsData, "Extranjeros", "2", tc

    Application.ScreenUpdating = True


End Sub

Private Sub CrearHojaProveedorTipo(ByVal wb As Workbook, ByVal wsData As Worksheet, _
                                   ByVal nombreHoja As String, ByVal prefijo As String, _
                                   ByVal tc As Double)

    Dim ws As Worksheet
    Dim ultFila As Long, i As Long
    Dim codProv As String, nomProv As String, comprador As String
    Dim moneda As String, montoUSD As Double, dias As Double
    Dim clave As String

    Dim dictProv As Object, dictComprador As Object
    Dim arr As Variant, arrComp As Variant
    Dim fila As Long, n As Long
    Dim totalUSD As Double, totalFilas As Long
    Dim provUnicos As Long, ticketProm As Double
    Dim top5 As Double

    Set dictProv = CreateObject("Scripting.Dictionary")
    Set dictComprador = CreateObject("Scripting.Dictionary")

    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = nombreHoja
    ws.Tab.Color = RGB(0, 112, 192)

    ultFila = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultFila

        codProv = Trim(CStr(wsData.Cells(i, "D").Value))
        nomProv = Trim(CStr(wsData.Cells(i, "E").Value))
        comprador = Trim(CStr(wsData.Cells(i, "AB").Value))

        If nomProv = "" Then nomProv = codProv
        If comprador = "" Then comprador = "Sin comprador"

        If Left(codProv, 1) = prefijo Then

            If prefijo = "2" Then
                montoUSD = Val(wsData.Cells(i, "L").Value) / tc
            Else
                moneda = UCase(Trim(CStr(wsData.Cells(i, "S").Value)))

                If moneda = "USD" Then
                    montoUSD = Val(wsData.Cells(i, "R").Value)
                ElseIf moneda = "PEN" Then
                    montoUSD = Val(wsData.Cells(i, "R").Value) / tc
                Else
                    montoUSD = 0
                End If
            End If

            If montoUSD <> 0 Then

                clave = codProv & "|" & nomProv

                If Not dictProv.Exists(clave) Then
                    dictProv.Add clave, Array(codProv, nomProv, 0#, 0&)
                End If

                Dim tmp As Variant
                tmp = dictProv(clave)
                tmp(2) = tmp(2) + montoUSD
                tmp(3) = tmp(3) + 1
                dictProv(clave) = tmp

                AddMontoLocal dictComprador, comprador, montoUSD

                totalUSD = totalUSD + montoUSD
                totalFilas = totalFilas + 1

            End If

        End If

    Next i

    provUnicos = dictProv.Count
    If totalFilas > 0 Then ticketProm = totalUSD / totalFilas

    arr = DictProvToSortedArray(dictProv)

    For i = 1 To WorksheetFunction.Min(5, UBound(arr, 1))
        top5 = top5 + arr(i, 3)
    Next i

    '====================
    ' CABECERA
    '====================
    With ws.Range("A1")
        If prefijo = "1" Then
            .Value = "PROVEEDORES NACIONALES — SALDOS PENDIENTES " & Year(Date)
        Else
            .Value = "PROVEEDORES EXTRANJEROS — SALDOS PENDIENTES " & Year(Date)
        End If
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(31, 78, 121)
    End With

    With ws.Range("A2")
        .Value = "Proveedores con código iniciando en " & prefijo
        .Font.Italic = True
        .Font.Color = RGB(90, 90, 90)
    End With

    '====================
    ' RESUMEN 3 FILAS
    '====================
    ws.Range("A5:F5").Merge
    ws.Range("A5").Value = "RESUMEN"
    FormatoTituloLocal ws.Range("A5:F5"), RGB(31, 78, 121)

    ws.Range("A6:F6").Value = Array("Total Pendiente USD", "Filas Pendientes", "Proveedores", "Ticket Promedio USD", "Top 5 Concentración", "")
    FormatoSubTituloLocal ws.Range("A6:F6")

    ws.Range("A7").Value = totalUSD
    ws.Range("B7").Value = totalFilas
    ws.Range("C7").Value = provUnicos
    ws.Range("D7").Value = ticketProm
    If totalUSD > 0 Then ws.Range("E7").Value = top5 / totalUSD

    ws.Range("A7").NumberFormat = "_(""US$"" * #,##0.00_);_(""US$"" * (#,##0.00);_(""US$"" * ""-""??_);_(@_)"
    ws.Range("D7").NumberFormat = "_(""US$"" * #,##0.00_);_(""US$"" * (#,##0.00);_(""US$"" * ""-""??_);_(@_)"
    ws.Range("E7").NumberFormat = "0.0%"

    FormatoTablaLocal ws.Range("A6:F7")

    '====================
    ' TABLA PROVEEDORES
    '====================
    fila = 10

    ws.Range("A" & fila & ":F" & fila).Value = Array("#", "Código", "Proveedor", "Filas Pendientes", "Monto Pendiente USD", "% del Total")
    FormatoTituloLocal ws.Range("A" & fila & ":F" & fila), RGB(31, 78, 121)

    n = UBound(arr, 1)

    For i = 1 To n
        ws.Cells(fila + i, "A").Value = i
        ws.Cells(fila + i, "B").Value = arr(i, 1)
        ws.Cells(fila + i, "C").Value = arr(i, 2)
        ws.Cells(fila + i, "D").Value = arr(i, 4)
        ws.Cells(fila + i, "E").Value = arr(i, 3)

        If totalUSD > 0 Then
            ws.Cells(fila + i, "F").Value = arr(i, 3) / totalUSD
        End If
    Next i

    ws.Range("E" & fila + 1 & ":E" & fila + n).NumberFormat = "_(""US$"" * #,##0.00_);_(""US$"" * (#,##0.00);_(""US$"" * ""-""??_);_(@_)"
    ws.Range("F" & fila + 1 & ":F" & fila + n).NumberFormat = "0.0%"

    FormatoTablaLocal ws.Range("A" & fila & ":F" & fila + n)

    '====================
    ' SALDO POR COMPRADOR
    '====================
    fila = fila + n + 3
    arrComp = DictToSortedArrayLocal(dictComprador)

    ws.Range("A" & fila & ":D" & fila).Merge
    If prefijo = "1" Then
        ws.Range("A" & fila).Value = "SALDOS POR COMPRADOR — PROVEEDORES NACIONALES"
    Else
        ws.Range("A" & fila).Value = "SALDOS POR COMPRADOR — PROVEEDORES EXTRANJEROS"
    End If

    FormatoTituloLocal ws.Range("A" & fila & ":D" & fila), RGB(84, 130, 53)

    ws.Range("A" & fila + 1 & ":D" & fila + 1).Value = Array("#", "Comprador", "Monto Pendiente USD", "% del Total")
    FormatoSubTituloLocal ws.Range("A" & fila + 1 & ":D" & fila + 1)

    For i = 1 To UBound(arrComp, 1)
        ws.Cells(fila + 1 + i, "A").Value = i
        ws.Cells(fila + 1 + i, "B").Value = arrComp(i, 1)
        ws.Cells(fila + 1 + i, "C").Value = arrComp(i, 2)

        If totalUSD > 0 Then
            ws.Cells(fila + 1 + i, "D").Value = arrComp(i, 2) / totalUSD
        End If
    Next i

    ws.Range("C" & fila + 2 & ":C" & fila + 1 + UBound(arrComp, 1)).NumberFormat = "_(""US$"" * #,##0.00_);_(""US$"" * (#,##0.00);_(""US$"" * ""-""??_);_(@_)"
    ws.Range("D" & fila + 2 & ":D" & fila + 1 + UBound(arrComp, 1)).NumberFormat = "0.0%"

    FormatoTablaLocal ws.Range("A" & fila + 1 & ":D" & fila + 1 + UBound(arrComp, 1))

    ws.Columns("A:F").AutoFit
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 45
    ws.Columns("A").HorizontalAlignment = xlCenter
    ws.Columns("D").HorizontalAlignment = xlCenter
    ws.Columns("F").HorizontalAlignment = xlCenter

End Sub

Private Sub AddMontoLocal(ByVal dict As Object, ByVal clave As String, ByVal valor As Double)
    If clave = "" Then clave = "Sin dato"
    If Not dict.Exists(clave) Then dict.Add clave, 0#
    dict(clave) = dict(clave) + valor
End Sub

Private Function DictProvToSortedArray(ByVal dict As Object) As Variant

    Dim arr() As Variant, k As Variant, tmp As Variant
    Dim i As Long, j As Long
    Dim t1, t2, t3, t4

    ReDim arr(1 To dict.Count, 1 To 4)

    i = 1
    For Each k In dict.keys
        tmp = dict(k)
        arr(i, 1) = tmp(0)
        arr(i, 2) = tmp(1)
        arr(i, 3) = tmp(2)
        arr(i, 4) = tmp(3)
        i = i + 1
    Next k

    For i = 1 To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If arr(j, 3) > arr(i, 3) Then
                t1 = arr(i, 1): t2 = arr(i, 2): t3 = arr(i, 3): t4 = arr(i, 4)
                arr(i, 1) = arr(j, 1): arr(i, 2) = arr(j, 2): arr(i, 3) = arr(j, 3): arr(i, 4) = arr(j, 4)
                arr(j, 1) = t1: arr(j, 2) = t2: arr(j, 3) = t3: arr(j, 4) = t4
            End If
        Next j
    Next i

    DictProvToSortedArray = arr

End Function

Private Function DictToSortedArrayLocal(ByVal dict As Object) As Variant

    Dim arr() As Variant, k As Variant
    Dim i As Long, j As Long
    Dim t1, t2

    ReDim arr(1 To dict.Count, 1 To 2)

    i = 1
    For Each k In dict.keys
        arr(i, 1) = k
        arr(i, 2) = dict(k)
        i = i + 1
    Next k

    For i = 1 To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If arr(j, 2) > arr(i, 2) Then
                t1 = arr(i, 1): t2 = arr(i, 2)
                arr(i, 1) = arr(j, 1): arr(i, 2) = arr(j, 2)
                arr(j, 1) = t1: arr(j, 2) = t2
            End If
        Next j
    Next i

    DictToSortedArrayLocal = arr

End Function

Private Sub FormatoTituloLocal(ByVal rng As Range, ByVal colorFondo As Long)
    With rng
        .Interior.Color = colorFondo
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub FormatoSubTituloLocal(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(217, 230, 242)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub FormatoTablaLocal(ByVal rng As Range)
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(60, 60, 60)
        .Font.Name = "Calibri"
        .Font.Size = 11
    End With
End Sub

