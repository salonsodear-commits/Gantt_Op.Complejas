Attribute VB_Name = "SprintAuto"
' ============================================================================
'  SPRINTS POR REUNIÓN (opcional) — Proyecto Datos Op. Complejas
' ----------------------------------------------------------------------------
'  El libro .xlsx YA gestiona los sprints automáticamente por fórmula:
'  un sprint es un PERÍODO que cierra en una reunión (calendario editable en
'  la hoja "Sprints"). Las tareas no completadas al cierre pasan solas al
'  sprint siguiente, mostrando el desvío; el motivo se elige de una lista.
'
'  Estas macros son OPCIONALES (requieren guardar como .xlsm):
'   - IndicarMotivo: BOTÓN de 1 clic para fijar el motivo del desvío en la
'     fila de la tarea seleccionada (escribe en la columna I de "Sprints").
'   - ReasignarSprintsFisico: write-back físico del nº de sprint en el Gantt
'     + bitácora persistente con fecha/hora y usuario.
'
'  Crear el botón: Insertar > (Controles de formulario) Botón -> asignar
'  la macro "IndicarMotivo".
' ============================================================================
Option Explicit

Const SH_GANTT As String = "Proyecto Datos (OKR´s)"
Const SH_SPR   As String = "Sprints"          ' calendario (K6:M13) + motivos (O5:O12) + col I motivo
Const SH_LOG   As String = "Log Movimientos"
Const COL_SPRINT As String = "E"   ' Sprint planificado (Gantt)
Const COL_TAREA  As String = "G"
Const COL_PROG   As String = "J"
Const COL_INI    As String = "K"
Const ROW_FIRST  As Long = 7
Const ROW_LAST   As Long = 200

' Botón: fija el motivo del desvío en la fila seleccionada de la hoja "Sprints".
Public Sub IndicarMotivo()
    Dim sp As Worksheet: Set sp = ThisWorkbook.Worksheets(SH_SPR)
    Dim r As Long: r = ActiveCell.Row
    If ActiveSheet.Name <> SH_SPR Or r < ROW_FIRST Or r > ROW_LAST Then
        MsgBox "Ubíquese en la hoja 'Sprints', en la fila de la tarea (filas " & _
               ROW_FIRST & " a " & ROW_LAST & ").", vbExclamation: Exit Sub
    End If
    Dim msg As String, i As Long, n As Long
    n = 0
    For i = 5 To 12
        If Len(sp.Cells(i, "O").Value) > 0 Then n = n + 1: msg = msg & n & ") " & sp.Cells(i, "O").Value & vbCrLf
    Next i
    Dim opt As String
    opt = InputBox("Motivo del desvío para la tarea de la fila " & r & ":" & vbCrLf & msg, _
                   "Indicar motivo del desvío")
    If opt = "" Then Exit Sub
    Dim idx As Long: idx = Val(opt)
    If idx >= 1 And idx <= n Then sp.Cells(r, "I").Value = sp.Cells(4 + idx, "O").Value Else sp.Cells(r, "I").Value = opt
    MsgBox "Motivo registrado: " & sp.Cells(r, "I").Value, vbInformation, "Sprints"
End Sub

' Opcional: reasigna FÍSICAMENTE el nº de sprint y deja bitácora persistente.
Public Sub ReasignarSprintsFisico()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim g As Worksheet: Set g = wb.Worksheets(SH_GANTT)
    Dim sp As Worksheet: Set sp = wb.Worksheets(SH_SPR)
    Dim lg As Worksheet: Set lg = GetOrCreateLog(wb)
    Dim sActual As Long: sActual = sp.Range("M15").Value     ' sprint actual (calculado)
    Dim r As Long, moved As Long
    For r = ROW_FIRST To ROW_LAST
        Dim tarea As Variant, prog As Variant, sprNum As Variant, ki As Variant, cierre As Variant
        tarea = g.Range(COL_TAREA & r).Value
        ki = g.Range(COL_INI & r).Value
        sprNum = g.Range(COL_SPRINT & r).Value
        prog = g.Range(COL_PROG & r).Value
        If Len(CStr(tarea)) > 0 And IsNumeric(sprNum) And IsDate(ki) Then
            cierre = CierreDeSprint(sp, CLng(sprNum))
            If Nz(prog) < 1 And IsDate(cierre) Then
                If Date > cierre And CLng(sprNum) < sActual Then
                    LogRow lg, "T" & Format(r, "000"), CStr(tarea), CStr(sprNum), CStr(sActual), _
                           sp.Range("I" & r).Value
                    g.Range(COL_SPRINT & r).Value = sActual
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
    For i = 6 To 13
        If sp.Cells(i, "K").Value = sprNum Then CierreDeSprint = sp.Cells(i, "M").Value: Exit Function
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
