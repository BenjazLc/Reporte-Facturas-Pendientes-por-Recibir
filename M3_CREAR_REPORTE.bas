Attribute VB_Name = "M3_CREAR_REPORTE"
Option Explicit

Public Sub Armar_Reporte_Facturas()

    On Error GoTo EH

    Dim rutaFacturacion As String
    Dim rutaOTIF As String
    Dim archivoFacturas As String
    Dim archivoOC As String
    Dim archivoEM As String
    Dim rutaSalida As String

    Dim wbFac As Workbook, wsFac As Worksheet
    Dim wbOC As Workbook, wsOC As Worksheet
    Dim wbEM As Workbook, wsEM As Worksheet
    Dim wbRep As Workbook, wsRep As Worksheet

    Dim dictOC As Object, dictEM As Object
    Dim ultFac As Long, ultOC As Long, ultEM As Long
    Dim i As Long, filaOut As Long
    Dim id As String

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    rutaFacturacion = "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\FACTURACION\"
    

  archivoFacturas = Dir(rutaFacturacion & "FACTURAS_TEMP*.*")
archivoOC = Dir(rutaFacturacion & "OC_TEMP*.*")
archivoEM = Dir(rutaFacturacion & "ENTREGA_MERCANCIAS_TEMP*.*")

If archivoFacturas = "" Then Err.Raise vbObjectError + 1, , "No se encontró FACTURAS_TEMP."
If archivoOC = "" Then Err.Raise vbObjectError + 2, , "No se encontró OC_TEMP."
If archivoEM = "" Then Err.Raise vbObjectError + 3, , "No se encontró ENTREGA_MERCANCIAS_TEMP."

Set wbFac = AbrirOUsarWorkbook(rutaFacturacion & archivoFacturas)
Set wbOC = AbrirOUsarWorkbook(rutaFacturacion & archivoOC)
Set wbEM = AbrirOUsarWorkbook(rutaFacturacion & archivoEM)

    Set wsFac = wbFac.Sheets(1)
    Set wsOC = wbOC.Sheets(1)
    Set wsEM = wbEM.Sheets(1)

    Set dictOC = CreateObject("Scripting.Dictionary")
    Set dictEM = CreateObject("Scripting.Dictionary")

    '========================
    ' Diccionario OC_TEMP
    ' Key: A+B
    ' E = Nombre
    ' G = Descripcion
    ' A = Fecha Orden
    ' M = Comprador
    '========================
    ultOC = wsOC.Cells(wsOC.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultOC
        id = CrearID(wsOC.Cells(i, "C").Value, wsOC.Cells(i, "D").Value)

        If id <> "" Then
            If Not dictOC.Exists(id) Then
                dictOC.Add id, Array( _
                    wsOC.Cells(i, "E").Value, _
                    wsOC.Cells(i, "G").Value, _
                    wsOC.Cells(i, "A").Value, _
                    wsOC.Cells(i, "M").Value)
            End If
        End If
    Next i

    '========================
    ' Diccionario ENTREGA_MERCANCIAS_TEMP
    ' Key: C+D
    ' N = Guia Remisión
    ' I = Fecha Ingreso
    '========================
    ultEM = wsEM.Cells(wsEM.Rows.Count, "C").End(xlUp).Row

    For i = 2 To ultEM
        id = CrearID(wsEM.Cells(i, "C").Value, wsEM.Cells(i, "D").Value)

        If id <> "" Then
            If Not dictEM.Exists(id) Then
                dictEM.Add id, Array( _
                    wsEM.Cells(i, "N").Value, _
                    wsEM.Cells(i, "I").Value)
            End If
        End If
    Next i

    '========================
    ' Crear reporte
    '========================
    Set wbRep = Workbooks.Add
    Set wsRep = wbRep.Sheets(1)
    wsRep.Name = "Saldos por Facturar"
    
    wsRep.Columns("C").NumberFormat = "@"

    ' Encabezados
    wsRep.Range("A1:AC1").Value = Array( _
        "Documento compras", "Posición", "ID", "Proveedor", "Nombre", _
        "Material", "Descripcion", "Precio neto pedido", "Moneda", _
        "Cantidad base", "Unidad medida pedido", "Valor entrada mcía.", _
        "Impte.fact.reg.ML", "Moneda", "Cantidad de pedido", _
        "Cantidad entrada", "Cantidad facturada", _
        "Saldo Por Facturar (Moneda Proveedor)", "Moneda", _
        "Saldo por entregar", "UM precio pedido", "Entrega final", _
        "Organización compras", "Grupo de compras", _
        "Guia Remisión (Puede No ser preciso)", "Fecha de Ingreso", _
        "Fecha Orden", "Comprador", "Conteo Días")

    ultFac = wsFac.Cells(wsFac.Rows.Count, "A").End(xlUp).Row
    filaOut = 2

    For i = 2 To ultFac

        id = CrearID(wsFac.Cells(i, "A").Value, wsFac.Cells(i, "B").Value)

        If id <> "" Then

            ' FACTURAS_TEMP base
            wsRep.Cells(filaOut, "A").Value = wsFac.Cells(i, "A").Value
            wsRep.Cells(filaOut, "B").Value = wsFac.Cells(i, "B").Value
            wsRep.Cells(filaOut, "C").Value = CStr(id)
            wsRep.Cells(filaOut, "D").Value = wsFac.Cells(i, "C").Value

            wsRep.Cells(filaOut, "F").Value = wsFac.Cells(i, "D").Value
            wsRep.Cells(filaOut, "H").Value = wsFac.Cells(i, "E").Value
            wsRep.Cells(filaOut, "I").Value = wsFac.Cells(i, "F").Value
            wsRep.Cells(filaOut, "J").Value = wsFac.Cells(i, "G").Value
            wsRep.Cells(filaOut, "K").Value = wsFac.Cells(i, "H").Value
            wsRep.Cells(filaOut, "L").Value = wsFac.Cells(i, "I").Value
            wsRep.Cells(filaOut, "M").Value = wsFac.Cells(i, "J").Value
            wsRep.Cells(filaOut, "N").Value = wsFac.Cells(i, "K").Value
            wsRep.Cells(filaOut, "O").Value = wsFac.Cells(i, "L").Value
            wsRep.Cells(filaOut, "P").Value = wsFac.Cells(i, "M").Value
            wsRep.Cells(filaOut, "Q").Value = wsFac.Cells(i, "N").Value

            ' Fórmulas nuevas
If Val(wsRep.Cells(filaOut, "J").Value) <> 0 Then
    wsRep.Cells(filaOut, "R").Value = (Val(wsRep.Cells(filaOut, "H").Value) / Val(wsRep.Cells(filaOut, "J").Value)) * _
                                      (Val(wsRep.Cells(filaOut, "P").Value) - Val(wsRep.Cells(filaOut, "Q").Value))
Else
    wsRep.Cells(filaOut, "R").Value = 0
End If

wsRep.Cells(filaOut, "S").Value = wsRep.Cells(filaOut, "I").Value
wsRep.Cells(filaOut, "T").Value = Val(wsRep.Cells(filaOut, "O").Value) - Val(wsRep.Cells(filaOut, "P").Value)



            ' Continúa FACTURAS_TEMP
            wsRep.Cells(filaOut, "U").Value = wsFac.Cells(i, "O").Value
            
            If Trim(CStr(wsFac.Cells(i, "P").Value)) = "X" Then
    wsRep.Cells(filaOut, "V").Value = "Pedido Concluido"
Else
    wsRep.Cells(filaOut, "V").Value = wsFac.Cells(i, "P").Value
End If
            wsRep.Cells(filaOut, "W").Value = wsFac.Cells(i, "Q").Value
            wsRep.Cells(filaOut, "X").Value = wsFac.Cells(i, "R").Value



            ' Cruce OC_TEMP
            If dictOC.Exists(id) Then
                wsRep.Cells(filaOut, "E").Value = LimpiarNombreProveedor(dictOC(id)(0))
                wsRep.Cells(filaOut, "G").Value = dictOC(id)(1)
                wsRep.Cells(filaOut, "AA").Value = dictOC(id)(2)
                wsRep.Cells(filaOut, "AB").Value = dictOC(id)(3)
            End If

            ' Cruce ENTREGA_MERCANCIAS_TEMP
            If dictEM.Exists(id) Then
                wsRep.Cells(filaOut, "Y").Value = dictEM(id)(0)
                wsRep.Cells(filaOut, "Z").Value = dictEM(id)(1)
            End If
            
            If IsDate(wsRep.Cells(filaOut, "Z").Value) Then
    wsRep.Cells(filaOut, "AC").Value = CDate(wsRep.Cells(filaOut, "Z").Value) - Date
Else
    wsRep.Cells(filaOut, "AC").Value = ""
End If



            filaOut = filaOut + 1

        End If

    Next i

    If filaOut = 2 Then Err.Raise vbObjectError + 4, , "No se copiaron registros al reporte."

    '========================
    ' Ordenar por D, A, B
    '========================
    With wsRep.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsRep.Range("D2:D" & filaOut - 1), Order:=xlAscending
        .SortFields.Add key:=wsRep.Range("A2:A" & filaOut - 1), Order:=xlAscending
        .SortFields.Add key:=wsRep.Range("B2:B" & filaOut - 1), Order:=xlAscending
        .SetRange wsRep.Range("A1:AC" & filaOut - 1)
        .Header = xlYes
        .Apply
    End With

    '========================
    ' Formato profesional
    '========================
    With wsRep.Range("A1:AC1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    wsRep.Range("A1:AC" & filaOut - 1).Borders.LineStyle = xlContinuous
    wsRep.Range("A1:AC" & filaOut - 1).Borders.Color = RGB(200, 200, 200)

    wsRep.Range("H:H,L:M,R:R").NumberFormat = "#,##0.00"
    wsRep.Range("J:J,O:Q,T:T").NumberFormat = "#,##0.00"
    wsRep.Range("Z:AA").NumberFormat = "dd/mm/yyyy"
    wsRep.Range("AC:AC").NumberFormat = "General"

    wsRep.Columns("A:AC").AutoFit
    wsRep.Rows(1).RowHeight = 36
    wsRep.Range("A2").Select
    ActiveWindow.FreezePanes = True

    '========================
    ' Sombreado por OC
    '========================
    AplicarSombreadoPorOC wsRep, filaOut - 1

    '========================
    ' Guardar reporte
    '========================
    rutaSalida = rutaFacturacion & "REPORTE_FACTURAS_" & Format(Date, "dd-mm-yyyy") & ".xlsx"

    If Dir(rutaSalida) <> "" Then Kill rutaSalida

    wbRep.SaveAs Filename:=rutaSalida, FileFormat:=xlOpenXMLWorkbook

SALIDA:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.CutCopyMode = False


    Exit Sub

EH:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.CutCopyMode = False

    MsgBox "Error en Armar_Reporte_Facturas: " & Err.Description, vbCritical

End Sub

Private Function AbrirOUsarWorkbook(ByVal rutaCompleta As String) As Workbook

    Dim wb As Workbook
    Dim nombreArchivo As String

    nombreArchivo = Dir(rutaCompleta)

    For Each wb In Application.Workbooks
        If UCase(wb.Name) = UCase(nombreArchivo) Then
            Set AbrirOUsarWorkbook = wb
            Exit Function
        End If
    Next wb

    Set AbrirOUsarWorkbook = Workbooks.Open(rutaCompleta)

End Function

Private Function CrearID(ByVal docCompra As Variant, ByVal posicion As Variant) As String

    Dim a As String
    Dim b As String

    a = LimpiarParteID(docCompra)
    b = LimpiarParteID(posicion)

    If a = "" Or b = "" Then
        CrearID = ""
    Else
        CrearID = a & b
    End If

End Function

Private Function LimpiarParteID(ByVal valor As Variant) As String

    If IsError(valor) Then
        LimpiarParteID = ""
    ElseIf Trim(CStr(valor)) = "" Then
        LimpiarParteID = ""
    ElseIf IsNumeric(valor) Then
        LimpiarParteID = Format(CDbl(valor), "0")
    Else
        LimpiarParteID = Trim(CStr(valor))
    End If

End Function

Private Sub AplicarSombreadoPorOC(ByVal ws As Worksheet, ByVal UltimaFila As Long)

    Dim i As Long
    Dim ocActual As String
    Dim ocAnterior As String
    Dim colorGrupo As Boolean

    ocAnterior = ""
    colorGrupo = False

    For i = 2 To UltimaFila

        ocActual = CStr(ws.Cells(i, "A").Value)

        If ocActual <> ocAnterior Then
            colorGrupo = Not colorGrupo
            ocAnterior = ocActual
        End If

        If colorGrupo Then
            ws.Range("A" & i & ":AC" & i).Interior.Color = RGB(221, 235, 247)
        Else
            ws.Range("A" & i & ":AC" & i).Interior.Color = RGB(255, 255, 255)
        End If

    Next i

End Sub

Private Function LimpiarNombreProveedor(ByVal texto As String) As String

    Dim partes() As String

    texto = Trim(texto)

    ' Divide por espacios
    partes = Split(texto, " ")

    ' Si el primer bloque es numérico, lo elimina
    If UBound(partes) >= 1 Then
        If IsNumeric(partes(0)) Then
            LimpiarNombreProveedor = Trim(Mid(texto, Len(partes(0)) + 1))
        Else
            LimpiarNombreProveedor = texto
        End If
    Else
        LimpiarNombreProveedor = texto
    End If

End Function

