Attribute VB_Name = "M0_MB51"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const RUTA_OTIF As String = _
    "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\OTIF\"

Private Const VARIANTE_MB51 As String = "BLAPA3"
Private Const ESPERA_CARGA_MS As Long = 2500

Public Sub OTIF_02()

    On Error GoTo EH

    Dim SapGuiAuto As Object
    Dim gui As Object
    Dim con As Object
    Dim ses As Object

    Dim fechaDesde As Date
    Dim fechaHasta As Date

    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False
    Application.CutCopyMode = False

    fechaDesde = DateSerial(Year(Date), 1, 1)
    fechaHasta = Date

    '========================
    ' Conectar a SAP
    '========================
    Set SapGuiAuto = GetObject("SAPGUI")
    Set gui = SapGuiAuto.GetScriptingEngine
    Set con = gui.Children(0)
    Set ses = con.Children(0)

    ses.findById("wnd[0]").Maximize
    ses.findById("wnd[0]").SetFocus
    AppActivate ses.findById("wnd[0]").Text
    DoEvents
    Sleep 300

    '========================
    ' Ir a MB51
    '========================
    ses.findById("wnd[0]/tbar[0]/okcd").Text = "/nMB51"
    ses.findById("wnd[0]").sendVKey 0
    Do While ses.Busy: DoEvents: Loop
    Sleep 500

    '========================
    ' Aplicar variante
    '========================
    ses.findById("wnd[0]").sendVKey 17
    Do While ses.Busy: DoEvents: Loop
    Sleep 300

    ses.findById("wnd[1]/usr/txtV-LOW").Text = VARIANTE_MB51
    ses.findById("wnd[1]/usr/txtV-LOW").CaretPosition = Len(VARIANTE_MB51)
    ses.findById("wnd[1]/tbar[0]/btn[8]").Press
    Do While ses.Busy: DoEvents: Loop
    Sleep 500

    '========================
    ' Fechas: desde inicio de año hasta hoy
    '========================
    ses.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = Format(fechaDesde, "dd.mm.yyyy")
    ses.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = Format(fechaHasta, "dd.mm.yyyy")
    ses.findById("wnd[0]/usr/ctxtBUDAT-HIGH").SetFocus
    ses.findById("wnd[0]/usr/ctxtBUDAT-HIGH").CaretPosition = 10
    ses.findById("wnd[0]").sendVKey 0
    Do While ses.Busy: DoEvents: Loop
    Sleep 300

    '========================
    ' Ejecutar
    '========================
    ses.findById("wnd[0]").sendVKey 8

    Sleep ESPERA_CARGA_MS

    '========================
    ' Cambiar de vista
    '========================
    ses.findById("wnd[0]").sendVKey 48
    Do While ses.Busy: DoEvents: Loop
    Sleep 500

    '========================
    ' Entregar control a PAD antes de exportar
    '========================
    ses.findById("wnd[0]").SetFocus
    AppActivate ses.findById("wnd[0]").Text
    DoEvents
    Sleep 500
    
    'MANDAR EXPORTAR
    SendKeys "+{F4}", True

SALIDA:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Application.CutCopyMode = False
    Exit Sub

EH:
    MsgBox "Error en OTIF_02: " & Err.Description, vbExclamation
    Resume SALIDA

End Sub


Public Sub EsperarALV(ByVal ses As Object, _
                      Optional ByVal timeoutMs As Long = 60000)

    Dim t As Double

    t = Timer

    Do

        DoEvents

        On Error Resume Next

        If ses.Busy = False Then

            If ses.findById("wnd[0]/usr/cntlGRID1/shellcont/shell") Is Nothing Then
                ' aún no carga
            Else
                Exit Do
            End If

        End If

        On Error GoTo 0

        Sleep 300

        If (Timer - t) * 1000 > timeoutMs Then
            Err.Raise vbObjectError + 1000, , _
                "Timeout esperando ALV."
        End If

    Loop

End Sub

