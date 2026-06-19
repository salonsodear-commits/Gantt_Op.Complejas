Attribute VB_Name = "SprintAuto"
' ============================================================================
'  SPRINTS POR REUNIÓN (opcional) — Proyecto Datos Op. Complejas
' ----------------------------------------------------------------------------
'  El libro .xlsx YA gestiona los sprints sin macros: un sprint es un PERÍODO
'  que cierra en una reunión (calendario editable en la hoja "Sprints"). Al
'  cerrar, las tareas no completadas pasan solas al sprint siguiente, con su
'  desvío; el motivo se elige de una lista y el comentario es texto libre.
'
'  Estas macros son OPCIONALES (requieren guardar como .xlsm) y dan un BOTÓN:
'   - PasarAlSiguiente : pone "Sí" en "Pasar al siguiente" de la fila activa
'     (y permite elegir el motivo). 1 clic.
'   - IndicarMotivo    : sólo fija el motivo del desvío de la fila activa.
'   - ReasignarSprintsFisico : write-back físico del nº de sprint + bitácora
'     persistente con fecha/hora y usuario.
'
'  Crear el botón: Insertar > (Controles de formulario) Botón -> asignar
'  la macro "PasarAlSiguiente".
' ============================================================================
Option Explicit

Const SH_GANTT As String = "Proyecto Datos (OKR´s)"
Const SH_SPR   As String = "Sprints"
Const SH_LOG   As String = "Log Movimientos"
Const COL_PASAR  As String = "H"   ' "Pasar al siguiente" (hoja Sprints)
Const COL_MOTIVO As String = "I"   ' "Motivo del desvío"  (hoja Sprints)
Const COL_GSPRINT As String = "E"  ' Sprint planificado (Gantt)
Const COL_GTAREA  As String = "G"
Const COL_GPROG   As String = "J"
Const COL_GINI    As String = "K"
Const ROW_FIRST  As Long = 7
Const ROW_LAST   As Long = 200

' BOTÓN: pasa la tarea de la fila activa al sprint siguiente y pide el motivo.
Public Sub PasarAlSiguiente()
    Dim sp As Worksheet: Set sp = ThisWorkbook.Worksheets(SH_SPR)
    Dim r As Long: r = ActiveCell.Row
    If ActiveSheet.Name <> SH_SPR Or r < ROW_FIRST Or r > ROW_LAST Then
        MsgBox "Ubíquese en la hoja 'Sprints', en la fila de la tarea.", vbExclamation: Exit Sub
    End If
    sp.Cells(r, COL_PASAR).Value = "Sí"
    PedirMotivo sp, r
    MsgBox "Tarea de la fila " & r & " marcada para pasar al sprint siguiente.", vbInformation, "Sprints"
End Sub

' BOTÓN: sólo fija el motivo del desvío de la fila activa.
Public Sub IndicarMotivo()
    Dim sp As Worksheet: Set sp = ThisWorkbook.Worksheets(SH_SPR)
    Dim r As Long: r = ActiveCell.Row
    If ActiveSheet.Name <> SH_SPR Or r < ROW_FIRST Or r > ROW_LAST Then
        MsgBox "Ubíquese en la hoja 'Sprints', en la fila de la tarea.", vbExclamation: Exit Sub
    End If
    PedirMotivo sp, r
End Sub

Private Sub PedirMotivo(sp As Worksheet, r As Long)
    Dim msg As String, i As Long, n As Long
    For i = 5 To 12      ' lista de motivos editable en P5:P12
        If Len(sp.Cells(i, "P").Value) > 0 Then n = n + 1: msg = msg & n & ") " & sp.Cells(i, "P").Value & vbCrLf
    Next i
    Dim opt As String
    opt = InputBox("Motivo del desvío (número o texto libre):" & vbCrLf & msg, "Motivo del desvío")
    If opt = "" Then Exit Sub
    Dim idx As Long: idx = Val(opt)
    If idx >= 1 And idx <= n Then sp.Cells(r, COL_MOTIVO).Value = sp.Cells(4 + idx, "P").Value _
                             Else sp.Cells(r, COL_MOTIVO).Value = opt
End Sub

' Opcional: reasigna FÍSICAMENTE el nº de sprint y deja bitácora persistente.
Public Sub ReasignarSprintsFisico()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim g As Worksheet: Set g = wb.Worksheets(SH_GANTT)
    Dim sp As Worksheet: Set sp = wb.Worksheets(SH_SPR)
    Dim lg As Worksheet: Set lg = GetOrCreateLog(wb)
    Dim sActual As Long: sActual = sp.Range("M14").Value     ' sprint actual (calculado)
    Dim r As Long, moved As Long
    For r = ROW_FIRST To ROW_LAST
        Dim tarea As Variant, prog As Variant, sprNum As Variant, ki As Variant, cierre As Variant
        tarea = g.Range(COL_GTAREA & r).Value
        ki = g.Range(COL_GINI & r).Value
        sprNum = g.Range(COL_GSPRINT & r).Value
        prog = g.Range(COL_GPROG & r).Value
        If Len(CStr(tarea)) > 0 And IsNumeric(sprNum) And IsDate(ki) Then
            cierre = CierreDeSprint(sp, CLng(sprNum))
            If Nz(prog) < 1 And IsDate(cierre) Then
                If Date > cierre And CLng(sprNum) < sActual Then
                    LogRow lg, "T" & Format(r, "000"), CStr(tarea), CStr(sprNum), CStr(sActual), _
                           sp.Range(COL_MOTIVO & r).Value
                    g.Range(COL_GSPRINT & r).Value = sActual
                    moved = moved + 1
                End If
            End If
        End If
    Next r
    MsgBox moved & " tarea(s) reasignada(s) al Sprint " & sActual & ". Bitácora en '" & SH_LOG & "'.", _
           vbInformation, "Sprints"
End Sub

Private Function CierreDeSprint(sp As Worksheet, sprNum As Long) As Variant
    Dim i As Long
    For i = 5 To 12      ' calendario: nº de sprint en L, fecha de cierre en N
        If sp.Cells(i, "L").Value = sprNum Then CierreDeSprint = sp.Cells(i, "N").Value: Exit Function
    Next i
    CierreDeSprint = Empty
End Function

Private Function Nz(v As Variant) As Double
    If IsNumeric(v) Then Nz = CDbl(v) Else Nz = 0
End Function

Private Function GetOrCreateLog(wb As Workbook) As Worksheet
    On Error Resume Next
    Set GetOrCreateLog = wb.Worksheets(SH_LOG)
    On Error GoTo 0
    If GetOrCreateLog Is Nothing Then
        Set GetOrCreateLog = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateLog.Name = SH_LOG
        GetOrCreateLog.Range("A1:G1").Value = Array("Fecha/Hora", "Usuario", _
            "ID", "Tarea", "Sprint origen", "Sprint destino", "Motivo del desvío")
    End If
End Function

Private Sub LogRow(lg As Worksheet, id As String, tarea As String, _
                   origen As String, destino As String, motivo As String)
    Dim nr As Long
    nr = lg.Cells(lg.Rows.Count, 1).End(xlUp).Row + 1
    lg.Cells(nr, 1).Value = Now
    lg.Cells(nr, 2).Value = Application.UserName
    lg.Cells(nr, 3).Value = id
    lg.Cells(nr, 4).Value = tarea
    lg.Cells(nr, 5).Value = origen
    lg.Cells(nr, 6).Value = destino
    lg.Cells(nr, 7).Value = IIf(Len(motivo) > 0, motivo, "No completada al cierre del sprint")
End Sub
