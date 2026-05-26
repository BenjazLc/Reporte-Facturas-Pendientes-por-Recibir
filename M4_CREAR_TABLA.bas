Attribute VB_Name = "M4_CREAR_TABLA"
Option Explicit

Public Sub M4_Resumen_Facturas()

    Dim wb As Workbook
    Dim wsData As Worksheet, wsRes As Worksheet
    Dim ultFila As Long, i As Long
    Dim tc As Double: tc = 3.45

    Dim proveedorCod As String, proveedorNom As String, comprador As String
    Dim monto As Double, dias As Double, monedaTipo As String
    Dim mesIngreso As String

    Dim dictProvMonto As Object, dictProvDiasMax As Object, dictProvDiasProm As Object
    Dim dictComprador As Object, dictMonedaMonto As Object, dictMonedaLineas As Object
    Dim dictAntMonto As Object, dictAntLineas As Object
    Dim dictCompradorLineas As Object
    
Dim totalNac As Double, totalExt As Double
Dim filasNac As Long, filasExt As Long
Dim totalUSD As Double
Dim diasPromTotal As Double
Dim diasMaxTotal As Double
Dim sumaDias As Double
Dim contadorDias As Long

Dim provNac As Object
Dim provExt As Object

Dim monedaPago As String
Dim montoPago As Double

Dim totalPagoUSD As Double
Dim totalPagoPEN As Double

Set wb = ObtenerLibroReporteFacturas()

If wb Is Nothing Then
    MsgBox "No encontré abierto el archivo de reporte de facturas con la hoja 'Saldos por Facturar'.", vbCritical
    Exit Sub
End If

Set wsData = wb.Sheets("Saldos por Facturar")


    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("Resumen").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsRes = wb.Sheets.Add(After:=wsData)
    wsRes.Name = "Resumen"
    wsRes.Tab.Color = RGB(0, 112, 192)

    Set dictProvMonto = CreateObject("Scripting.Dictionary")
    Set dictProvDiasMax = CreateObject("Scripting.Dictionary")
    Set dictProvDiasProm = CreateObject("Scripting.Dictionary")
    Set dictComprador = CreateObject("Scripting.Dictionary")
    Set dictCompradorLineas = CreateObject("Scripting.Dictionary")
    Set dictMonedaMonto = CreateObject("Scripting.Dictionary")
    Set dictMonedaLineas = CreateObject("Scripting.Dictionary")
    Set dictAntMonto = CreateObject("Scripting.Dictionary")
    Set dictAntLineas = CreateObject("Scripting.Dictionary")
    
    Set provNac = CreateObject("Scripting.Dictionary")
Set provExt = CreateObject("Scripting.Dictionary")

    ultFila = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row





   For i = 2 To ultFila

    proveedorCod = Trim(CStr(wsData.Cells(i, "D").Value))
    proveedorNom = Trim(CStr(wsData.Cells(i, "E").Value))
    comprador = Trim(CStr(wsData.Cells(i, "AB").Value))

    If proveedorNom = "" Then proveedorNom = proveedorCod

    If Left(proveedorCod, 1) = "2" Then

        ' Extranjeros: usar L en PEN y convertir a USD
        monto = Val(wsData.Cells(i, "L").Value) / tc
        monedaTipo = "USD"

    Else

        ' Nacionales: usar R según moneda S y convertir todo a USD
        monedaTipo = UCase(Trim(CStr(wsData.Cells(i, "S").Value)))

        If monedaTipo = "USD" Then
            monto = Val(wsData.Cells(i, "R").Value)
        ElseIf monedaTipo = "PEN" Then
            monto = Val(wsData.Cells(i, "R").Value) / tc
        Else
            monto = 0
        End If

    End If

    dias = Abs(Val(wsData.Cells(i, "AC").Value))

If monto <> 0 Then
    
    If Left(proveedorCod, 1) = "2" Then
        totalExt = totalExt + monto
        filasExt = filasExt + 1
        
        If Not provExt.Exists(proveedorNom) Then provExt.Add proveedorNom, True
    Else
        totalNac = totalNac + monto
        filasNac = filasNac + 1
        
        If Not provNac.Exists(proveedorNom) Then provNac.Add proveedorNom, True
    End If
    

        AddMonto dictProvMonto, proveedorNom, monto
        AddMax dictProvDiasMax, proveedorNom, dias
        AddProm dictProvDiasProm, proveedorNom, dias

        AddMonto dictComprador, comprador, monto
        AddMonto dictCompradorLineas, comprador, 1
        
        


If Left(proveedorCod, 1) = "2" Then

    monedaPago = UCase(Trim(CStr(wsData.Cells(i, "I").Value)))

    If monedaPago = "USD" Then
        montoPago = Val(wsData.Cells(i, "L").Value) / tc
    ElseIf monedaPago = "PEN" Then
        montoPago = Val(wsData.Cells(i, "L").Value)
    Else
        montoPago = 0
    End If

Else

    monedaPago = UCase(Trim(CStr(wsData.Cells(i, "S").Value)))
    montoPago = Val(wsData.Cells(i, "R").Value)

End If

If montoPago <> 0 Then

AddMonto dictMonedaMonto, monedaPago, montoPago
AddMonto dictMonedaLineas, monedaPago, 1

If monedaPago = "USD" Then
    totalPagoUSD = totalPagoUSD + montoPago
ElseIf monedaPago = "PEN" Then
    totalPagoPEN = totalPagoPEN + montoPago
End If

End If



        AddAntiguedad dictAntMonto, dictAntLineas, dias, monto

        ' Indicadores clave
        sumaDias = sumaDias + dias
        contadorDias = contadorDias + 1

        If dias > diasMaxTotal Then
            diasMaxTotal = dias
        End If

    End If

Next i

totalUSD = totalNac + totalExt
If contadorDias > 0 Then
    diasPromTotal = sumaDias / contadorDias
End If
    '========================
    ' TITULO
    '========================
    With wsRes.Range("A1")
        .Value = "SALDOS PENDIENTES POR FACTURAR — INGRESOS " & Year(Date)
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(31, 78, 121)
    End With

    With wsRes.Range("A2")
        .Value = "Reporte de mercadería recibida pendiente de facturación"
        .Font.Italic = True
        .Font.Color = RGB(90, 90, 90)
    End With

    '========================
    ' RESUMEN GENERAL
    '========================
    wsRes.Range("A5:D5").Merge
    wsRes.Range("A5").Value = "RESUMEN GENERAL"
    FormatoTitulo wsRes.Range("A5:D5"), RGB(31, 78, 121)

    wsRes.Range("A6:D6").Value = Array("", "Nacionales", "Extranjeros", "Total")
    FormatoSubTitulo wsRes.Range("A6:D6")

    wsRes.Range("A7:A11").Value = Application.Transpose(Array( _
        "Saldo Pendiente (USD)", _
        "Filas Pendientes", _
        "Proveedores", _
        "Ticket Promedio (USD)", _
        "% del Total Pendiente"))



    wsRes.Range("B7").Value = totalNac
    wsRes.Range("C7").Value = totalExt
    wsRes.Range("D7").Value = totalNac + totalExt

    wsRes.Range("B8").Value = filasNac
    wsRes.Range("C8").Value = filasExt
    wsRes.Range("D8").Value = filasNac + filasExt

    wsRes.Range("B9").Value = provNac.Count
    wsRes.Range("C9").Value = provExt.Count
    wsRes.Range("D9").Value = provNac.Count + provExt.Count

    If filasNac > 0 Then wsRes.Range("B10").Value = totalNac / filasNac
    If filasExt > 0 Then wsRes.Range("C10").Value = totalExt / filasExt
    If filasNac + filasExt > 0 Then wsRes.Range("D10").Value = (totalNac + totalExt) / (filasNac + filasExt)

    If totalNac + totalExt > 0 Then
        wsRes.Range("B11").Value = totalNac / (totalNac + totalExt)
        wsRes.Range("C11").Value = totalExt / (totalNac + totalExt)
        wsRes.Range("D11").Value = 1
    End If

    wsRes.Range("B7:D10").NumberFormat = "#,##0"
    wsRes.Range("B11:D11").NumberFormat = "0.0%"
    FormatoTabla wsRes.Range("A6:D11")
    
'========================
' INDICADORES CLAVE
'========================
wsRes.Range("F5:J5").Merge
wsRes.Range("F5").Value = "INDICADORES CLAVE"
FormatoTitulo wsRes.Range("F5:J5"), RGB(47, 84, 150)

' Encabezados
wsRes.Range("F6:H6").Merge
wsRes.Range("F6").Value = "Indicador"

wsRes.Range("I6:J6").Merge
wsRes.Range("I6").Value = "Valor"

FormatoSubTitulo wsRes.Range("F6:J6")

' Filas
wsRes.Range("F7:H7").Merge
wsRes.Range("F7").Value = "Proveedores Únicos"
wsRes.Range("I7:J7").Merge
wsRes.Range("I7").Value = provNac.Count + provExt.Count

wsRes.Range("F8:H8").Merge
wsRes.Range("F8").Value = "Saldo Total USD"
wsRes.Range("I8:J8").Merge
wsRes.Range("I8").Value = totalPagoUSD

wsRes.Range("F9:H9").Merge
wsRes.Range("F9").Value = "Saldo Total PEN"
wsRes.Range("I9:J9").Merge
wsRes.Range("I9").Value = totalPagoPEN

wsRes.Range("F10:H10").Merge
wsRes.Range("F10").Value = "Promedio Días Pendientes"
wsRes.Range("I10:J10").Merge
wsRes.Range("I10").Value = diasPromTotal

wsRes.Range("F11:H11").Merge
wsRes.Range("F11").Value = "Máx. Días Atraso"
wsRes.Range("I11:J11").Merge
wsRes.Range("I11").Value = diasMaxTotal

' Formatos
' USD
wsRes.Range("I8:J8").NumberFormat = _
"_(""US$"" * #,##0.00_);_(""US$"" * (#,##0.00);_(""US$"" * ""-""??_);_(@_)"

' PEN
wsRes.Range("I9:J9").NumberFormat = _
"_(""S/"" * #,##0.00_);_(""S/"" * (#,##0.00);_(""S/"" * ""-""??_);_(@_)"


wsRes.Range("I10:J11").NumberFormat = "0"

wsRes.Range("I7:J11").HorizontalAlignment = xlCenter
wsRes.Range("F6:J11").VerticalAlignment = xlCenter

FormatoTabla wsRes.Range("F6:J11")

    '========================
    ' TOP 15 MONTO
    '========================
    CrearTopMonto wsRes, dictProvMonto, 14

    '========================
    ' TOP 15 DIAS
    '========================
    CrearTopDias wsRes, dictProvDiasMax, dictProvDiasProm, 14

    '========================
    ' RANKING COMPRADORES
    '========================
    CrearRankingCompradores wsRes, dictComprador, dictCompradorLineas, 33

    '========================
    ' DISTRIBUCION MONEDA
    '========================
    CrearDistribucionMoneda wsRes, dictMonedaMonto, dictMonedaLineas, 33

    '========================
    ' ANTIGUEDAD
    '========================
    CrearAntiguedad wsRes, dictAntMonto, dictAntLineas, 42

    wsRes.Columns("A:J").AutoFit
    
    wsRes.Columns("A").ColumnWidth = 20
    wsRes.Columns("B").ColumnWidth = 20
    wsRes.Columns("C").ColumnWidth = 20
    
    wsRes.Columns("G").ColumnWidth = 20
wsRes.Columns("H").ColumnWidth = 20

   

End Sub

Private Sub AddMonto(ByVal dict As Object, ByVal clave As String, ByVal valor As Double)
    If clave = "" Then clave = "Sin dato"
    If Not dict.Exists(clave) Then dict.Add clave, 0#
    dict(clave) = dict(clave) + valor
End Sub

Private Sub AddMax(ByVal dict As Object, ByVal clave As String, ByVal valor As Double)
    If clave = "" Then clave = "Sin dato"
    If Not dict.Exists(clave) Then
        dict.Add clave, valor
    ElseIf valor > dict(clave) Then
        dict(clave) = valor
    End If
End Sub

Private Sub AddProm(ByVal dict As Object, ByVal clave As String, ByVal valor As Double)

    Dim arr As Variant

    If clave = "" Then clave = "Sin dato"

    If Not dict.Exists(clave) Then
        dict.Add clave, Array(valor, 1)
    Else
        arr = dict.Item(clave)
        arr(0) = CDbl(arr(0)) + valor
        arr(1) = CLng(arr(1)) + 1
        dict.Item(clave) = arr
    End If

End Sub

Private Sub AddAntiguedad(ByVal dictMonto As Object, ByVal dictLineas As Object, ByVal dias As Double, ByVal monto As Double)

    Dim rango As String

    If dias <= 15 Then
        rango = "0-15 días"
    ElseIf dias <= 30 Then
        rango = "16-30 días"
    ElseIf dias <= 60 Then
        rango = "31-60 días"
    ElseIf dias <= 90 Then
        rango = "61-90 días"
    Else
        rango = ">90 días"
    End If

    AddMonto dictMonto, rango, monto
    AddMonto dictLineas, rango, 1

End Sub

Private Sub CrearTopMonto(ByVal ws As Worksheet, ByVal dict As Object, ByVal filaIni As Long)

    Dim arr As Variant, i As Long, n As Long
    arr = DictToSortedArray(dict, True)

    ws.Range("A" & filaIni & ":D" & filaIni).Merge
    ws.Range("A" & filaIni).Value = "TOP 15 PROVEEDORES POR MONTO PENDIENTE"
    FormatoTitulo ws.Range("A" & filaIni & ":D" & filaIni), RGB(47, 84, 150)

    ws.Range("A" & filaIni + 1).Value = "#"
ws.Range("B" & filaIni + 1 & ":C" & filaIni + 1).Merge
ws.Range("B" & filaIni + 1).Value = "Proveedor"
ws.Range("D" & filaIni + 1).Value = "Monto Pendiente"


    FormatoSubTitulo ws.Range("A" & filaIni + 1 & ":D" & filaIni + 1)

    n = WorksheetFunction.Min(15, UBound(arr, 1))

    For i = 1 To n
        ws.Cells(filaIni + 1 + i, "A").Value = i
        ws.Range("B" & filaIni + 1 + i & ":C" & filaIni + 1 + i).Merge
        ws.Cells(filaIni + 1 + i, "B").Value = arr(i, 1)
        ws.Cells(filaIni + 1 + i, "D").Value = arr(i, 2)
    Next i
    
    ws.Range("A" & filaIni + 1 & ":A" & filaIni + 16).HorizontalAlignment = xlCenter
ws.Range("A" & filaIni + 1 & ":A" & filaIni + 16).VerticalAlignment = xlCenter

    ws.Range("D" & filaIni + 2 & ":D" & filaIni + 16).NumberFormat = _
"_(""US$"" * #,##0.00_);_(""US$"" * (#,##0.00);_(""US$"" * ""-""??_);_(@_)"

    FormatoTabla ws.Range("A" & filaIni + 1 & ":D" & filaIni + 16)

End Sub

Private Sub CrearTopDias(ByVal ws As Worksheet, ByVal dictMax As Object, ByVal dictProm As Object, ByVal filaIni As Long)

    Dim arr As Variant, i As Long, n As Long
    Dim proveedor As String, promArr As Variant

    arr = DictToSortedArray(dictMax, True)

    ws.Range("F" & filaIni & ":J" & filaIni).Merge
    ws.Range("F" & filaIni).Value = "TOP 15 PROVEEDORES POR DÍAS DE ATRASO"
    FormatoTitulo ws.Range("F" & filaIni & ":J" & filaIni), RGB(192, 0, 0)

    ws.Cells(filaIni + 1, "F").Value = "#"

ws.Range("G" & filaIni + 1 & ":H" & filaIni + 1).Merge
ws.Cells(filaIni + 1, "G").Value = "Proveedor"

ws.Cells(filaIni + 1, "I").Value = "Máx Días"
ws.Cells(filaIni + 1, "J").Value = "Prom. Días"

    FormatoSubTitulo ws.Range("F" & filaIni + 1 & ":J" & filaIni + 1)

    n = WorksheetFunction.Min(15, UBound(arr, 1))

For i = 1 To n

    proveedor = arr(i, 1)
    promArr = dictProm(proveedor)

    ws.Cells(filaIni + 1 + i, "F").Value = i

    ws.Range("G" & filaIni + 1 + i & ":H" & filaIni + 1 + i).Merge
    ws.Cells(filaIni + 1 + i, "G").Value = proveedor

    ws.Cells(filaIni + 1 + i, "I").Value = arr(i, 2)

    If promArr(1) <> 0 Then
        ws.Cells(filaIni + 1 + i, "J").Value = promArr(0) / promArr(1)
    Else
        ws.Cells(filaIni + 1 + i, "J").Value = 0
    End If

Next i

' Centrar columnas F, I y J
ws.Range("F" & filaIni + 1 & ":F" & filaIni + 1 + n).HorizontalAlignment = xlCenter
ws.Range("I" & filaIni + 1 & ":I" & filaIni + 1 + n).HorizontalAlignment = xlCenter
ws.Range("J" & filaIni + 1 & ":J" & filaIni + 1 + n).HorizontalAlignment = xlCenter

' Centrado vertical general
ws.Range("F" & filaIni + 1 & ":J" & filaIni + 1 + n).VerticalAlignment = xlCenter

    ws.Range("I" & filaIni + 2 & ":J" & filaIni + 16).NumberFormat = "0"
    FormatoTabla ws.Range("F" & filaIni + 1 & ":J" & filaIni + 16)
    
    ws.Columns("G").ColumnWidth = 25
ws.Columns("H").ColumnWidth = 25

End Sub

Private Sub CrearRankingCompradores(ByVal ws As Worksheet, _
                                   ByVal dictMonto As Object, _
                                   ByVal dictLineas As Object, _
                                   ByVal filaIni As Long)

    Dim arr As Variant, i As Long, n As Long
    Dim totalMonto As Double

    arr = DictToSortedArray(dictMonto, True)

    ws.Range("A" & filaIni & ":D" & filaIni).Merge
    ws.Range("A" & filaIni).Value = "RANKING COMPRADORES"
    FormatoTitulo ws.Range("A" & filaIni & ":D" & filaIni), RGB(84, 130, 53)

    ws.Cells(filaIni + 1, "A").Value = "# Líneas"
    ws.Cells(filaIni + 1, "B").Value = "Comprador"
    ws.Cells(filaIni + 1, "C").Value = "Monto USD"
    ws.Cells(filaIni + 1, "D").Value = "% Participación"

    FormatoSubTitulo ws.Range("A" & filaIni + 1 & ":D" & filaIni + 1)

    n = WorksheetFunction.Min(10, UBound(arr, 1))

    ' Total para %
    For i = 1 To n
        totalMonto = totalMonto + arr(i, 2)
    Next i

    For i = 1 To n

        ws.Cells(filaIni + 1 + i, "B").Value = arr(i, 1)
        ws.Cells(filaIni + 1 + i, "C").Value = arr(i, 2)

        ' # líneas
        If dictLineas.Exists(arr(i, 1)) Then
            ws.Cells(filaIni + 1 + i, "A").Value = dictLineas(arr(i, 1))
        End If

        ' %
        If totalMonto > 0 Then
            ws.Cells(filaIni + 1 + i, "D").Value = arr(i, 2) / totalMonto
        End If

    Next i

    ' Formatos
    ws.Range("C" & filaIni + 2 & ":C" & filaIni + 1 + n).NumberFormat = _
    "_(""US$"" * #,##0.00_);_(""US$"" * (#,##0.00);_(""US$"" * ""-""??_);_(@_)"

    ws.Range("D" & filaIni + 2 & ":D" & filaIni + 1 + n).NumberFormat = "0.0%"

    ws.Range("A" & filaIni + 1 & ":A" & filaIni + 1 + n).HorizontalAlignment = xlCenter
    ws.Range("C" & filaIni + 1 & ":D" & filaIni + 1 + n).HorizontalAlignment = xlCenter
    ws.Range("A" & filaIni + 1 & ":D" & filaIni + 1 + n).VerticalAlignment = xlCenter

    FormatoTabla ws.Range("A" & filaIni + 1 & ":D" & filaIni + 1 + n)

End Sub

Private Sub CrearDistribucionMoneda(ByVal ws As Worksheet, ByVal dictMonto As Object, ByVal dictLineas As Object, ByVal filaIni As Long)

    Dim monedas As Variant, i As Long, totalLineas As Double
    monedas = Array("USD", "PEN")

    totalLineas = 0
    For i = LBound(monedas) To UBound(monedas)
        If dictLineas.Exists(monedas(i)) Then totalLineas = totalLineas + dictLineas(monedas(i))
    Next i

    ws.Range("F" & filaIni & ":J" & filaIni).Merge
    ws.Range("F" & filaIni).Value = "DISTRIBUCIÓN POR MONEDA"
    FormatoTitulo ws.Range("F" & filaIni & ":J" & filaIni), RGB(191, 143, 0)

    ws.Cells(filaIni + 1, "F").Value = "Moneda"

ws.Range("G" & filaIni + 1 & ":H" & filaIni + 1).Merge
ws.Cells(filaIni + 1, "G").Value = "Monto Total"

ws.Cells(filaIni + 1, "I").Value = "# Líneas"
ws.Cells(filaIni + 1, "J").Value = "% Líneas"


    FormatoSubTitulo ws.Range("F" & filaIni + 1 & ":J" & filaIni + 1)

For i = 0 To 1

    ws.Cells(filaIni + 2 + i, "F").Value = monedas(i)

    ws.Range("G" & filaIni + 2 + i & ":H" & filaIni + 2 + i).Merge
    If dictMonto.Exists(monedas(i)) Then
        ws.Cells(filaIni + 2 + i, "G").Value = dictMonto(monedas(i))
    End If

    If dictLineas.Exists(monedas(i)) Then
        ws.Cells(filaIni + 2 + i, "I").Value = dictLineas(monedas(i))
    End If

    If totalLineas > 0 Then
        ws.Cells(filaIni + 2 + i, "J").Value = ws.Cells(filaIni + 2 + i, "I").Value / totalLineas
    End If

Next i

ws.Range("G" & filaIni + 2 & ":H" & filaIni + 3).NumberFormat = "#,##0.00"

ws.Range("F" & filaIni + 1 & ":F" & filaIni + 3).HorizontalAlignment = xlCenter
ws.Range("F" & filaIni + 1 & ":F" & filaIni + 3).VerticalAlignment = xlCenter

    ws.Range("H" & filaIni + 2 & ":H" & filaIni + 3).NumberFormat = "#,##0"
    ws.Range("J" & filaIni + 2 & ":J" & filaIni + 3).NumberFormat = "0.0%"
    FormatoTabla ws.Range("F" & filaIni + 1 & ":J" & filaIni + 3)

End Sub

Private Sub CrearAntiguedad(ByVal ws As Worksheet, ByVal dictMonto As Object, ByVal dictLineas As Object, ByVal filaIni As Long)

    Dim rangos As Variant, i As Long, totalMonto As Double
    rangos = Array("0-15 días", "16-30 días", "31-60 días", "61-90 días", ">90 días")

    For i = LBound(rangos) To UBound(rangos)
        If dictMonto.Exists(rangos(i)) Then totalMonto = totalMonto + dictMonto(rangos(i))
    Next i

    ws.Range("F" & filaIni & ":J" & filaIni).Merge
    ws.Range("F" & filaIni).Value = "ANÁLISIS DE ANTIGÜEDAD (DÍAS PENDIENTES)"
    FormatoTitulo ws.Range("F" & filaIni & ":J" & filaIni), RGB(112, 48, 160)

    ws.Cells(filaIni + 1, "F").Value = "Rango de Días"

ws.Range("G" & filaIni + 1 & ":H" & filaIni + 1).Merge
ws.Cells(filaIni + 1, "G").Value = "Monto"

ws.Cells(filaIni + 1, "I").Value = "# Líneas"
ws.Cells(filaIni + 1, "J").Value = "% Monto"


    FormatoSubTitulo ws.Range("F" & filaIni + 1 & ":J" & filaIni + 1)

    For i = 0 To 4
        ws.Cells(filaIni + 2 + i, "F").Value = rangos(i)
        ws.Range("G" & filaIni + 2 + i & ":H" & filaIni + 2 + i).Merge

If dictMonto.Exists(rangos(i)) Then
    ws.Cells(filaIni + 2 + i, "G").Value = dictMonto(rangos(i))
End If


        If dictLineas.Exists(rangos(i)) Then ws.Cells(filaIni + 2 + i, "I").Value = dictLineas(rangos(i))
If totalMonto > 0 Then
    ws.Cells(filaIni + 2 + i, "J").Value = ws.Cells(filaIni + 2 + i, "G").Value / totalMonto
End If

Next i

ws.Range("G" & filaIni + 2 & ":H" & filaIni + 6).NumberFormat = _
"_([$USD] * #,##0.00_);_([$USD] * (#,##0.00);_([$USD] * ""-""??_);_(@_)"

ws.Range("J" & filaIni + 2 & ":J" & filaIni + 6).NumberFormat = "0.0%"

ws.Range("F" & filaIni + 1 & ":F" & filaIni + 6).HorizontalAlignment = xlCenter
ws.Range("I" & filaIni + 1 & ":I" & filaIni + 6).HorizontalAlignment = xlCenter
ws.Range("J" & filaIni + 1 & ":J" & filaIni + 6).HorizontalAlignment = xlCenter
ws.Range("F" & filaIni + 1 & ":J" & filaIni + 6).VerticalAlignment = xlCenter

    FormatoTabla ws.Range("F" & filaIni + 1 & ":J" & filaIni + 6)

End Sub

Private Function DictToSortedArray(ByVal dict As Object, ByVal desc As Boolean) As Variant

    Dim arr() As Variant
    Dim i As Long, j As Long
    Dim k As Variant
    Dim tmp1 As Variant, tmp2 As Variant

    ReDim arr(1 To dict.Count, 1 To 2)

    i = 1
    For Each k In dict.keys
        arr(i, 1) = k
        arr(i, 2) = dict(k)
        i = i + 1
    Next k

    For i = 1 To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If (desc And arr(j, 2) > arr(i, 2)) Or (Not desc And arr(j, 2) < arr(i, 2)) Then
                tmp1 = arr(i, 1): tmp2 = arr(i, 2)
                arr(i, 1) = arr(j, 1): arr(i, 2) = arr(j, 2)
                arr(j, 1) = tmp1: arr(j, 2) = tmp2
            End If
        Next j
    Next i

    DictToSortedArray = arr

End Function

Private Sub FormatoTitulo(ByVal rng As Range, ByVal colorFondo As Long)
    With rng
        .Interior.Color = colorFondo
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub FormatoSubTitulo(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(217, 230, 242)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub FormatoTabla(ByVal rng As Range)
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(60, 60, 60)
        .Font.Name = "Calibri"
        .Font.Size = 11
    End With
End Sub

Public Function ObtenerLibroReporteFacturas() As Workbook

    Dim wb As Workbook
    Dim ws As Worksheet

    For Each wb In Application.Workbooks

        ' Evita trabajar sobre el libro donde está la macro
        If wb.Name <> ThisWorkbook.Name Then

            On Error Resume Next
            Set ws = wb.Sheets("Saldos por Facturar")
            On Error GoTo 0

            If Not ws Is Nothing Then
                Set ObtenerLibroReporteFacturas = wb
                Exit Function
            End If

            Set ws = Nothing

        End If

    Next wb

End Function

