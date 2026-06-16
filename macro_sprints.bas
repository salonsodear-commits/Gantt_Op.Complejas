Attribute VB_Name = "SprintAuto"
' ============================================================================
'  AUTOMATIZACIÓN DE SPRINTS (opcional) — Proyecto Datos Op. Complejas
' ----------------------------------------------------------------------------
'  El libro .xlsx YA reasigna los sprints de forma automática por fórmula
'  (columna "Sprint vigente" en la hoja _Datos + hoja "Historial de Sprints").
'
'  Esta macro es OPCIONAL y agrega, para quien la necesite:
'    - Reasignación FÍSICA del número de sprint en el Gantt (write-back).
'    - Un registro PERSISTENTE (no se recalcula) con FECHA/HORA y USUARIO.
'
'  Cómo usarla:
'    1) Guarde el libro como .xlsm (Libro de Excel habilitado para macros).
'    2) Alt+F11  ->  Insertar > Módulo  ->  pegue este código.
'    3) Ejecute "ReasignarSprints" (o descomente Workbook_Open para que
'       se ejecute solo al abrir el archivo).
' ============================================================================
Option Explicit

Const SH_GANTT As String = "Proyecto Datos (OKR´s)"   ' hoja del Gantt
Const SH_CONF  As String = "Historial de Sprints"     ' config (C5 inicio, C6 duración)
Const SH_LOG   As String = "Log Movimientos"          ' bitácora persistente
Const COL_SPRINT As String = "E"   ' Sprint (planificado)
Const COL_TAREA  As String = "G"   ' Tarea
Const COL_PROG   As String = "J"   ' Progreso (0..1)
Const COL_INI    As String = "K"   ' Inicio (fecha)
Const ROW_FIRST  As Long = 7
Const ROW_LAST   As Long = 200

Public Sub ReasignarSprints()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim g As Worksheet, cf As Worksheet, lg As Worksheet
    Set g = wb.Worksheets(SH_GANTT)
    Set cf = wb.Worksheets(SH_CONF)
    Set lg = GetOrCreateLog(wb)

    Dim ini As Date, dur As Long, sActual As Long
    ini = cf.Range("C5").Value
    dur = cf.Range("C6").Value
    If dur < 1 Then dur = 14
    sActual = Application.WorksheetFunction.Max(1, Int((Date - ini) / dur) + 1)

    Dim r As Long, moved As Long
    For r = ROW_FIRST To ROW_LAST
        Dim tarea As Variant, prog As Variant, sp As Variant, ki As Variant
        tarea = g.Range(COL_TAREA & r).Value
        ki = g.Range(COL_INI & r).Value
        sp = g.Range(COL_SPRINT & r).Value
        prog = g.Range(COL_PROG & r).Value
        ' Solo tareas reales (tienen Inicio) con nº de sprint y no completadas
        If Len(CStr(tarea)) > 0 And IsNumeric(sp) And IsDate(ki) Then
            If Nz(prog) < 1 And CLng(sp) < sActual Then
                Dim cierre As Date: cierre = ini + CLng(sp) * dur - 1
                If Date > cierre Then
                    LogRow lg, "T" & Format(r, "000"), CStr(tarea), CStr(sp), CStr(sActual), _
                           "No completada al cierre del Sprint " & sp
                    g.Range(COL_SPRINT & r).Value = sActual   ' reasignación física
                    moved = moved + 1
                End If
            End If
        End If
    Next r
    MsgBox moved & " tarea(s) reasignada(s) al Sprint " & sActual & _
           ". Registro en la hoja '" & SH_LOG & "'.", vbInformation, "Sprints"
End Sub

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
            "ID", "Tarea", "Sprint origen", "Sprint destino", "Motivo")
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
    lg.Cells(nr, 7).Value = motivo
End Sub

' ============================================================================
'  BOTÓN "POSTERGAR TAREA"  (1 clic, con selección de motivo)
'  Posterga 1 sprint la tarea de la fila seleccionada y registra el motivo.
'  Sólo rellena las celdas de la hoja "Postergar"; las fórmulas hacen el resto
'  (Sprint vigente se recalcula y el movimiento aparece en "Historial de Sprints").
'  Para crear el botón: Insertar > (Controles de formulario) Botón ->
'  asignar la macro "PostergarTareaActual".
' ============================================================================
Public Sub PostergarTareaActual()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim pg As Worksheet: Set pg = wb.Worksheets("Postergar")
    Dim r As Long: r = ActiveCell.Row
    If r < ROW_FIRST Or r > ROW_LAST Then
        MsgBox "Seleccione la fila de una tarea (filas " & ROW_FIRST & " a " & ROW_LAST & ").", vbExclamation
        Exit Sub
    End If
    Dim motivos As Variant
    motivos = Array("Dependencia bloqueada", "Sobrecarga del responsable", _
        "Falta de información / insumos", "Reestimación / mayor alcance", _
        "Recurso no disponible", "Prioridad reasignada", "Bloqueo técnico", "Otro")
    Dim msg As String, i As Integer
    For i = LBound(motivos) To UBound(motivos)
        msg = msg & (i + 1) & ") " & motivos(i) & vbCrLf
    Next i
    Dim opt As String
    opt = InputBox("Motivo del cuello de botella (1-" & (UBound(motivos) + 1) & "):" & vbCrLf & msg, _
                   "Postergar tarea al próximo sprint")
    If opt = "" Then Exit Sub
    Dim idx As Integer: idx = Val(opt)
    Dim motivo As String
    If idx >= 1 And idx <= UBound(motivos) + 1 Then motivo = motivos(idx - 1) Else motivo = opt
    pg.Range("F" & r).Value = Nz(pg.Range("F" & r).Value) + 1   ' +1 sprint
    pg.Range("G" & r).Value = motivo
    MsgBox "Tarea (fila " & r & ") postergada 1 sprint. Motivo: " & motivo & _
           vbCrLf & "Vea el detalle en 'Historial de Sprints'.", vbInformation, "Sprints"
End Sub

' --- Para ejecución automática al abrir: pegue esto en el objeto "ThisWorkbook" ---
' Private Sub Workbook_Open()
'     ReasignarSprints
' End Sub
