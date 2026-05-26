Attribute VB_Name = "M2_ME2N"
Option Explicit

Public Sub Descargar_ME2N()

    On Error GoTo EH

    Dim rutaCarpeta As String
    Dim archivo As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim yaAbierto As Boolean

    Dim ultFila As Long
    Dim i As Long
    Dim oc As String
    Dim textoOC As String
    Dim dictOC As Object

    Dim html As Object

    Dim SapGuiAuto As Object
    Dim SAPApp As Object
    Dim SAPCon As Object
    Dim session As Object

    '========================
    ' Configuración Excel
    '========================
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    Application.CutCopyMode = False

    '========================
    ' Ruta y archivo fuente
    '========================
    rutaCarpeta = "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\FACTURACION\"
    archivo = Dir(rutaCarpeta & "FACTURAS_TEMP*.*")

    If archivo = "" Then
        MsgBox "No se encontró el archivo FACTURAS_TEMP en la ruta FACTURACION.", vbCritical
        GoTo SALIDA
    End If

    '========================
    ' Usar archivo si ya está abierto, si no abrirlo
    '========================
    yaAbierto = False

    For Each wb In Application.Workbooks
        If UCase(wb.Name) Like UCase("FACTURAS_TEMP*") Then
            yaAbierto = True
            Exit For
        End If
    Next wb

    If Not yaAbierto Then
        Set wb = Workbooks.Open(rutaCarpeta & archivo)
    End If

    Set ws = wb.Sheets(1)

    '========================
    ' Leer OC columna A sin duplicados
    '========================
    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set dictOC = CreateObject("Scripting.Dictionary")

    For i = 2 To ultFila

        oc = Trim(CStr(ws.Cells(i, "A").Value))

        If oc <> "" Then
            If Not dictOC.Exists(oc) Then
                dictOC.Add oc, True
                textoOC = textoOC & oc & vbCrLf
            End If
        End If

    Next i

    If textoOC = "" Then
        MsgBox "No se encontraron órdenes de compra en la columna A de FACTURAS_TEMP.", vbExclamation
        GoTo SALIDA
    End If

    '========================
    ' Copiar OC al portapapeles
    '========================
    Set html = CreateObject("htmlfile")
    html.ParentWindow.ClipboardData.SetData "text", textoOC

    '========================
    ' Conectar a SAP
    '========================
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set SAPCon = SAPApp.Children(0)
    Set session = SAPCon.Children(0)

    '========================
    ' Enfocar SAP
    '========================
    session.findById("wnd[0]").Maximize

    On Error Resume Next
    AppActivate session.findById("wnd[0]").Text
    If Err.Number <> 0 Then
        Err.Clear
        AppActivate "SAP"
    End If
    On Error GoTo EH

    Application.Wait Now + TimeValue("0:00:01")

    '========================
    ' Ir a ME2N
    '========================
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nME2N"
    session.findById("wnd[0]").sendVKey 0

    Do While session.Busy
        DoEvents
    Loop

    '========================
    ' Abrir selección múltiple de OC
    '========================
    session.findById("wnd[0]/usr/btn%_EN_EBELN_%_APP_%-VALU_PUSH").Press

    Do While session.Busy
        DoEvents
    Loop

    '========================
    ' Pegar OC desde portapapeles
    '========================
    session.findById("wnd[1]").sendVKey 24

    Do While session.Busy
        DoEvents
    Loop

    'Aceptar selección múltiple
    session.findById("wnd[1]").sendVKey 8

    Do While session.Busy
        DoEvents
    Loop

    '========================
    ' Ejecutar reporte
    '========================
    session.findById("wnd[0]").sendVKey 8

    Do While session.Busy
        DoEvents
    Loop

    '========================
    ' Reactivar Excel antes de entregar a PAD
    '========================
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Application.CutCopyMode = False
    Application.StatusBar = False

    '========================
    ' Enfocar SAP antes de exportar
    '========================
    session.findById("wnd[0]").Maximize

    On Error Resume Next
    AppActivate session.findById("wnd[0]").Text
    If Err.Number <> 0 Then
        Err.Clear
        AppActivate "SAP"
    End If
    On Error GoTo EH

    Application.Wait Now + TimeValue("0:00:01")
    DoEvents

    '========================
    ' Exportar: Ctrl + Shift + F7
    ' En este módulo el 43 ES la exportación
    '========================
    SendKeys "^+{F7}", False

    DoEvents

    'Macro termina aquí para que PAD tome control
    Exit Sub

SALIDA:

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Application.CutCopyMode = False
    Application.StatusBar = False

    Exit Sub

EH:

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Application.CutCopyMode = False
    Application.StatusBar = False

    MsgBox "Error en Descargar_ME2N_FACTURAS_TEMP: " & Err.Description, vbCritical

End Sub
