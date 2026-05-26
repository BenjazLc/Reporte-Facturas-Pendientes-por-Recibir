Attribute VB_Name = "M1_MB5S"
Sub DESCARGAR_MB5S()

    Dim rutaCarpeta As String
    Dim archivo As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long
    Dim textoOC As String
    Dim oc As String
    Dim session As Object
    Dim WshShell As Object

    rutaCarpeta = "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\FACTURACION\"
    
    archivo = Dir(rutaCarpeta & "ENTREGA_MERCANCIAS_TEMP*.*")
    
    If archivo = "" Then
        MsgBox "No se encontró el archivo ENTREGA_MERCANCIAS_TEMP en la ruta OTIF.", vbCritical
        Exit Sub
    End If

Dim yaAbierto As Boolean
Dim nombreArchivo As String

nombreArchivo = archivo
yaAbierto = False

' Buscar si ya está abierto
For Each wb In Application.Workbooks
    If wb.Name = nombreArchivo Then
        yaAbierto = True
        Exit For
    End If
Next wb

' Si no está abierto, lo abre
If Not yaAbierto Then
    Set wb = Workbooks.Open(rutaCarpeta & nombreArchivo)
Else
    Set wb = Workbooks(nombreArchivo)
End If
Set ws = wb.Sheets(1)
    ultFila = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
Dim dictOC As Object
Set dictOC = CreateObject("Scripting.Dictionary")
For i = 2 To ultFila
    oc = Trim(CStr(ws.Cells(i, "C").Value))
    If oc <> "" Then
        If Not dictOC.Exists(oc) Then
            dictOC.Add oc, True
            textoOC = textoOC & oc & vbCrLf
        End If
    End If
Next i

    
    ' Traer SAP al frente (nivel Windows)
    
    ' Tomar sesión SAP activa
Dim SapGuiAuto As Object
Dim SAPApp As Object
Dim SAPCon As Object


Set SapGuiAuto = GetObject("SAPGUI")
Set SAPApp = SapGuiAuto.GetScriptingEngine
Set SAPCon = SAPApp.Children(0)
Set session = SAPCon.Children(0)

If session Is Nothing Then
    MsgBox "No se pudo tomar la sesión activa de SAP.", vbCritical
    Exit Sub
End If
session.findById("wnd[0]").Maximize

On Error Resume Next
AppActivate session.findById("wnd[0]").Text
If Err.Number <> 0 Then
    Err.Clear
    AppActivate "SAP"
End If
On Error GoTo 0

Application.Wait Now + TimeValue("0:00:01")
'---------------------------------------------------------------------------------

' Tomar sesión SAP

Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)

' Asegurar foco dentro de SAP (nivel SAP)
session.findById("wnd[0]").Maximize



    If textoOC = "" Then
        MsgBox "No se encontraron órdenes de compra en la columna C.", vbExclamation
        Exit Sub
    End If

' Copiar OC al portapapeles (método alternativo sin MSForms)
Dim html As Object
Set html = CreateObject("htmlfile")

html.ParentWindow.ClipboardData.SetData "text", textoOC

    ' Tomar sesión SAP activa
    Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)

    ' Entrar a MB5S
    session.findById("wnd[0]").Maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMB5S"
    session.findById("wnd[0]").sendVKey 0
    
    'Abrir Variante
    session.findById("wnd[0]").sendVKey 17
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 3
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = "3"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

    ' Abrir selección múltiple de Orden de Compra
    session.findById("wnd[0]/usr/btn%_EBELN_%_APP_%-VALU_PUSH").Press

    ' Pegar desde portapapeles
    session.findById("wnd[1]").sendVKey 24

    ' Aceptar selección múltiple
    session.findById("wnd[1]").sendVKey 8
    
' Desactivar alertas justo antes de ejecutar
Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.AskToUpdateLinks = False

DoEvents

    ' Ejecutar reporte
    session.findById("wnd[0]").sendVKey 8
    
    ' Reactivar alertas

    Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.AskToUpdateLinks = True

    ' Detalle de Transaccion
    session.findById("wnd[0]").sendVKey 43
'-----------------------------------------------------------------------------------
    
' Forzar foco en SAP antes de usar SendKeys de Windows
session.findById("wnd[0]").Maximize

On Error Resume Next
AppActivate session.findById("wnd[0]").Text

If Err.Number <> 0 Then
    Err.Clear
    AppActivate "SAP"
End If

On Error GoTo 0

DoEvents
Application.Wait Now + TimeValue("0:00:01")

' Exportar con Shift + F4
SendKeys "+{F4}", False

DoEvents

Exit Sub

End Sub
