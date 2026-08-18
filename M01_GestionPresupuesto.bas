Attribute VB_Name = "M01_GestionPresupuesto"
Option Explicit

' =====================================================================
' M01_GestionPresupuesto (v3 - tablas reales, sin tope de filas)
'
' Hoja "Presupuesto":
'   Encabezado: C4=Numero E4=Fecha C5=Cliente C6=Telefono C7=Moneda C8=Estado
'   Configurador: L4=Producto L5=Talla L6=Tela L7=Tecnica L8=Volumen L9=Cantidad
'                 O4:O8=códigos resueltos (auxiliar, no tocar a mano)
'                 L16=Costo total unitario USD   L17=Precio unitario en moneda mostrada
'                 L26=Línea en edición (automático, ya no se escribe a mano)
'   Tabla T_Items (crece sola): ITEM|CANT|DESCRIPCIÓN|COSTO UNIT|TOTAL|Auxiliar
'   C30=Buscar Comanda (desplegable "Numero - Cliente")
'   F26=Total a pagar (=SUM(T_Items[TOTAL]))  F27=Abono
'   C34=Tasa COP por USD  C35=Tasa VES por USD
'
' Hoja "Comandas":
'   T_Comandas: Numero|Fecha|Cliente|Telefono|Moneda|Total|Abono|Resta|Estado|Etiqueta
'   T_ComandasDetalle: Numero|Item|Cantidad|Descripcion|CostoUnit|Total|Auxiliar
'
' Hoja "TasasCambio": B=Moneda C=Tasa_vs_USD (USD fila3, COP fila4, VES fila5)
'
' Si cambias nombres de hoja/tabla o el orden de columnas, actualiza las
' constantes y los ListColumns("...") de abajo.
' =====================================================================

Private Const HOJA As String = "Presupuesto"
Private Const HOJA_COM As String = "Comandas"
Private Const HOJA_TASAS As String = "TasasCambio"
Private Const TABLA_ITEMS As String = "T_Items"
Private Const TABLA_COM As String = "T_Comandas"
Private Const TABLA_DET As String = "T_ComandasDetalle"
Private Const FILA_EDITAR As Long = 26

Private Function ObtenerHoja() As Worksheet
    On Error GoTo Falla
    Set ObtenerHoja = ThisWorkbook.Worksheets(HOJA)
    Exit Function
Falla:
    MsgBox "No se encontró la hoja '" & HOJA & "'.", vbCritical
    Set ObtenerHoja = Nothing
End Function

Private Function ObtenerHojaComandas() As Worksheet
    On Error GoTo Falla
    Set ObtenerHojaComandas = ThisWorkbook.Worksheets(HOJA_COM)
    Exit Function
Falla:
    MsgBox "No se encontró la hoja '" & HOJA_COM & "'.", vbCritical
    Set ObtenerHojaComandas = Nothing
End Function

' =====================================================================
' AGREGAR / ACTUALIZAR ITEM (tabla real T_Items, sin tope de filas)
' =====================================================================
Public Sub AgregarOActualizarItem()
    Dim ws As Worksheet
    Set ws = ObtenerHoja()
    If ws Is Nothing Then Exit Sub

    Dim lo As ListObject
    Set lo = ws.ListObjects(TABLA_ITEMS)

    Application.Calculate  ' fuerza recálculo antes de leer el precio, evita valores viejos

    Dim producto As String, talla As String, tela As String, tecnica As String, volumen As String
    Dim cantRaw As Variant, cant As Double, precioUnit As Variant, descripcion As String

    producto = Trim(ws.Range("L4").Value)
    talla = Trim(ws.Range("L5").Value)
    tela = Trim(ws.Range("L6").Value)
    tecnica = Trim(ws.Range("L7").Value)
    volumen = Trim(ws.Range("L8").Value)
    cantRaw = ws.Range("L9").Value
    precioUnit = ws.Range("L17").Value

    If producto = "" Then MsgBox "Selecciona un Producto antes de agregar.", vbExclamation: Exit Sub
    If talla = "" Then MsgBox "Selecciona una Talla antes de agregar.", vbExclamation: Exit Sub
    If tecnica = "" Then MsgBox "Selecciona una Técnica antes de agregar.", vbExclamation: Exit Sub
    If volumen = "" Then MsgBox "Selecciona un Volumen antes de agregar.", vbExclamation: Exit Sub
    If Not IsNumeric(cantRaw) Then MsgBox "La Cantidad debe ser un número.", vbExclamation: Exit Sub
    cant = CDbl(cantRaw)
    If cant <= 0 Then MsgBox "La Cantidad debe ser mayor a 0.", vbExclamation: Exit Sub
    If Not IsNumeric(precioUnit) Then
        MsgBox "Esa combinación no tiene precio calculable. No se agregó la línea.", vbExclamation
        Exit Sub
    End If

    descripcion = producto & " | " & talla
    If tela <> "" Then descripcion = descripcion & " | " & tela
    descripcion = descripcion & " | " & tecnica

    Dim aux As String
    aux = producto & "|" & talla & "|" & tela & "|" & tecnica & "|" & volumen

    Dim filaEditar As Variant
    filaEditar = ws.Range("L" & FILA_EDITAR).Value

    Dim lr As ListRow
    Dim numeroItem As Long

    If filaEditar <> "" And IsNumeric(filaEditar) Then
        Set lr = BuscarListRowPorItem(lo, CLng(filaEditar))
        If lr Is Nothing Then
            MsgBox "No se encontró el ITEM Nº " & filaEditar & " en la tabla. No se actualizó nada.", vbExclamation
            Exit Sub
        End If
        numeroItem = CLng(filaEditar)
    Else
        Set lr = lo.ListRows.Add
        numeroItem = ProximoNumeroItem(lo)
    End If

    With lr.Range
        .Cells(1, lo.ListColumns("ITEM").Index).Value = numeroItem
        .Cells(1, lo.ListColumns("CANT").Index).Value = cant
        .Cells(1, lo.ListColumns("DESCRIPCIÓN").Index).Value = descripcion
        .Cells(1, lo.ListColumns("COSTO UNIT").Index).Value = precioUnit
        .Cells(1, lo.ListColumns("TOTAL").Index).Value = cant * precioUnit
        .Cells(1, lo.ListColumns("Auxiliar").Index).Value = aux
    End With

    Call LimpiarConfigurador
    MsgBox "Línea " & numeroItem & " guardada.", vbInformation
End Sub

' =====================================================================
' CARGAR ITEM PARA EDITAR (te paras en la fila de la tabla y clic)
' =====================================================================
Public Sub CargarItemParaEditar()
    Dim ws As Worksheet
    Set ws = ObtenerHoja()
    If ws Is Nothing Then Exit Sub

    Dim lo As ListObject
    Set lo = ws.ListObjects(TABLA_ITEMS)

    If ActiveSheet.Name <> HOJA Then
        MsgBox "Párate primero en una fila de la tabla de items.", vbExclamation
        Exit Sub
    End If
    If lo.DataBodyRange Is Nothing Or Intersect(ActiveCell, lo.DataBodyRange) Is Nothing Then
        MsgBox "Párate en una fila de la tabla de items (T_Items) antes de cargar.", vbExclamation
        Exit Sub
    End If

    Dim filaHoja As Long
    filaHoja = ActiveCell.Row
    Dim itemVal As Variant
    itemVal = ws.Cells(filaHoja, lo.ListColumns("ITEM").Range.Column).Value
    If itemVal = "" Then
        MsgBox "Esa fila está vacía, no hay nada que cargar.", vbExclamation
        Exit Sub
    End If

    Dim aux As String
    aux = ws.Cells(filaHoja, lo.ListColumns("Auxiliar").Range.Column).Value
    If Trim(aux) = "" Then
        MsgBox "Esa línea no tiene datos de configuración guardados. Edítala directo en la tabla.", vbExclamation
        Exit Sub
    End If

    Dim partes() As String
    partes = Split(aux, "|")
    If UBound(partes) < 4 Then
        MsgBox "Los datos auxiliares de esa línea están incompletos.", vbExclamation
        Exit Sub
    End If

    ws.Range("L4").Value = partes(0)
    ws.Range("L5").Value = partes(1)
    ws.Range("L6").Value = partes(2)
    ws.Range("L7").Value = partes(3)
    ws.Range("L8").Value = partes(4)
    ws.Range("L9").Value = ws.Cells(filaHoja, lo.ListColumns("CANT").Range.Column).Value
    ws.Range("L" & FILA_EDITAR).Value = itemVal
End Sub

' =====================================================================
' LIMPIAR CONFIGURADOR
' =====================================================================
Public Sub LimpiarConfigurador()
    Dim ws As Worksheet
    Set ws = ObtenerHoja()
    If ws Is Nothing Then Exit Sub

    ws.Range("L4").Value = ""
    ws.Range("L5").Value = ""
    ws.Range("L6").Value = ""
    ws.Range("L7").Value = ""
    ws.Range("L8").Value = ""
    ws.Range("L9").Value = 1
    ws.Range("L" & FILA_EDITAR).Value = ""
End Sub

' =====================================================================
' ELIMINAR ITEM (te paras en la fila y clic)
' =====================================================================
Public Sub EliminarItem()
    Dim ws As Worksheet
    Set ws = ObtenerHoja()
    If ws Is Nothing Then Exit Sub

    Dim lo As ListObject
    Set lo = ws.ListObjects(TABLA_ITEMS)

    If ActiveSheet.Name <> HOJA Then
        MsgBox "Párate primero en la fila que quieres eliminar.", vbExclamation
        Exit Sub
    End If
    If lo.DataBodyRange Is Nothing Or Intersect(ActiveCell, lo.DataBodyRange) Is Nothing Then
        MsgBox "Párate en una fila de la tabla de items antes de eliminar.", vbExclamation
        Exit Sub
    End If

    Dim filaHoja As Long
    filaHoja = ActiveCell.Row
    Dim itemVal As Variant
    itemVal = ws.Cells(filaHoja, lo.ListColumns("ITEM").Range.Column).Value
    If itemVal = "" Then
        MsgBox "Esa fila ya está vacía.", vbExclamation
        Exit Sub
    End If

    If MsgBox("¿Eliminar la línea ITEM Nº " & itemVal & "?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    Dim lr As ListRow
    Set lr = lo.ListRows(filaHoja - lo.DataBodyRange.Row + 1)
    lr.Delete
    Call LimpiarConfigurador
End Sub

' =====================================================================
' ACEPTAR PRESUPUESTO -> guarda encabezado + líneas en Comandas
' =====================================================================
Public Sub AceptarPresupuesto()
    Dim ws As Worksheet, wsC As Worksheet
    Set ws = ObtenerHoja()
    If ws Is Nothing Then Exit Sub
    Set wsC = ObtenerHojaComandas()
    If wsC Is Nothing Then Exit Sub

    Dim numero As String
    numero = Trim(ws.Range("C4").Value)
    If numero = "" Then
        MsgBox "Escribe el Número del presupuesto antes de aceptarlo.", vbExclamation
        Exit Sub
    End If

    Dim loItems As ListObject
    Set loItems = ws.ListObjects(TABLA_ITEMS)
    If loItems.DataBodyRange Is Nothing Then
        MsgBox "No hay líneas en el presupuesto.", vbExclamation
        Exit Sub
    End If

    Dim loCom As ListObject, loDet As ListObject
    Set loCom = wsC.ListObjects(TABLA_COM)
    Set loDet = wsC.ListObjects(TABLA_DET)

    ' --- Borrar detalle previo de esa misma comanda (se re-escribe completo) ---
    If Not loDet.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = loDet.ListRows.Count To 1 Step -1
            If loDet.ListRows(i).Range.Cells(1, loDet.ListColumns("Numero").Index).Value = numero Then
                loDet.ListRows(i).Delete
            End If
        Next i
    End If

    Dim total As Double, abono As Double
    total = 0: abono = 0
    On Error Resume Next
    total = ws.Range("F26").Value
    abono = ws.Range("F27").Value
    On Error GoTo 0

    Dim lrCom As ListRow
    Set lrCom = Nothing
    If Not loCom.DataBodyRange Is Nothing Then
        Dim r As Range
        For Each r In loCom.ListColumns("Numero").DataBodyRange
            If r.Value = numero Then
                Set lrCom = loCom.ListRows(r.Row - loCom.DataBodyRange.Row + 1)
                Exit For
            End If
        Next r
    End If
    If lrCom Is Nothing Then Set lrCom = loCom.ListRows.Add

    With lrCom.Range
        .Cells(1, loCom.ListColumns("Numero").Index).Value = numero
        .Cells(1, loCom.ListColumns("Fecha").Index).Value = ws.Range("E4").Value
        .Cells(1, loCom.ListColumns("Cliente").Index).Value = ws.Range("C5").Value
        .Cells(1, loCom.ListColumns("Telefono").Index).Value = ws.Range("C6").Value
        .Cells(1, loCom.ListColumns("Moneda").Index).Value = ws.Range("C7").Value
        .Cells(1, loCom.ListColumns("Total").Index).Value = total
        .Cells(1, loCom.ListColumns("Abono").Index).Value = abono
        .Cells(1, loCom.ListColumns("Resta").Index).Value = total - abono
        .Cells(1, loCom.ListColumns("Estado").Index).Value = "Aceptado"
    End With

    Dim itemCell As Range
    For Each itemCell In loItems.ListColumns("ITEM").DataBodyRange
        If itemCell.Value <> "" Then
            Dim filaHoja As Long
            filaHoja = itemCell.Row
            Dim lrDet As ListRow
            Set lrDet = loDet.ListRows.Add
            With lrDet.Range
                .Cells(1, loDet.ListColumns("Numero").Index).Value = numero
                .Cells(1, loDet.ListColumns("Item").Index).Value = itemCell.Value
                .Cells(1, loDet.ListColumns("Cantidad").Index).Value = ws.Cells(filaHoja, loItems.ListColumns("CANT").Range.Column).Value
                .Cells(1, loDet.ListColumns("Descripcion").Index).Value = ws.Cells(filaHoja, loItems.ListColumns("DESCRIPCIÓN").Range.Column).Value
                .Cells(1, loDet.ListColumns("CostoUnit").Index).Value = ws.Cells(filaHoja, loItems.ListColumns("COSTO UNIT").Range.Column).Value
                .Cells(1, loDet.ListColumns("Total").Index).Value = ws.Cells(filaHoja, loItems.ListColumns("TOTAL").Range.Column).Value
                .Cells(1, loDet.ListColumns("Auxiliar").Index).Value = ws.Cells(filaHoja, loItems.ListColumns("Auxiliar").Range.Column).Value
            End With
        End If
    Next itemCell

    ws.Range("C8").Value = "Aceptado"
    MsgBox "Presupuesto " & numero & " guardado en Comandas.", vbInformation
End Sub

' =====================================================================
' CARGAR COMANDA -> repuebla el Presupuesto completo desde Comandas
' =====================================================================
Public Sub CargarComanda()
    Dim ws As Worksheet, wsC As Worksheet
    Set ws = ObtenerHoja()
    If ws Is Nothing Then Exit Sub
    Set wsC = ObtenerHojaComandas()
    If wsC Is Nothing Then Exit Sub

    Dim etiqueta As String
    etiqueta = Trim(ws.Range("C30").Value)
    If etiqueta = "" Then
        MsgBox "Elige un presupuesto en 'Buscar Comanda' antes de cargar.", vbExclamation
        Exit Sub
    End If

    Dim loCom As ListObject, loDet As ListObject
    Set loCom = wsC.ListObjects(TABLA_COM)
    Set loDet = wsC.ListObjects(TABLA_DET)

    Dim numero As String
    numero = ""
    Dim r As Range
    If Not loCom.DataBodyRange Is Nothing Then
        For Each r In loCom.ListColumns("Etiqueta").DataBodyRange
            If r.Value = etiqueta Then
                Dim filaCom As Long
                filaCom = r.Row
                numero = wsC.Cells(filaCom, loCom.ListColumns("Numero").Range.Column).Value
                ws.Range("C4").Value = numero
                ws.Range("E4").Value = wsC.Cells(filaCom, loCom.ListColumns("Fecha").Range.Column).Value
                ws.Range("C5").Value = wsC.Cells(filaCom, loCom.ListColumns("Cliente").Range.Column).Value
                ws.Range("C6").Value = wsC.Cells(filaCom, loCom.ListColumns("Telefono").Range.Column).Value
                ws.Range("C7").Value = wsC.Cells(filaCom, loCom.ListColumns("Moneda").Range.Column).Value
                ws.Range("C8").Value = wsC.Cells(filaCom, loCom.ListColumns("Estado").Range.Column).Value
                ws.Range("F27").Value = wsC.Cells(filaCom, loCom.ListColumns("Abono").Range.Column).Value
                Exit For
            End If
        Next r
    End If
    If numero = "" Then
        MsgBox "No se encontró esa comanda.", vbExclamation
        Exit Sub
    End If

    Dim loItems As ListObject
    Set loItems = ws.ListObjects(TABLA_ITEMS)
    Do While Not loItems.DataBodyRange Is Nothing
        loItems.ListRows(1).Delete
    Loop

    If Not loDet.DataBodyRange Is Nothing Then
        Dim d As Range
        For Each d In loDet.ListColumns("Numero").DataBodyRange
            If d.Value = numero Then
                Dim filaDet As Long
                filaDet = d.Row
                Dim lr As ListRow
                Set lr = loItems.ListRows.Add
                With lr.Range
                    .Cells(1, loItems.ListColumns("ITEM").Index).Value = wsC.Cells(filaDet, loDet.ListColumns("Item").Range.Column).Value
                    .Cells(1, loItems.ListColumns("CANT").Index).Value = wsC.Cells(filaDet, loDet.ListColumns("Cantidad").Range.Column).Value
                    .Cells(1, loItems.ListColumns("DESCRIPCIÓN").Index).Value = wsC.Cells(filaDet, loDet.ListColumns("Descripcion").Range.Column).Value
                    .Cells(1, loItems.ListColumns("COSTO UNIT").Index).Value = wsC.Cells(filaDet, loDet.ListColumns("CostoUnit").Range.Column).Value
                    .Cells(1, loItems.ListColumns("TOTAL").Index).Value = wsC.Cells(filaDet, loDet.ListColumns("Total").Range.Column).Value
                    .Cells(1, loItems.ListColumns("Auxiliar").Index).Value = wsC.Cells(filaDet, loDet.ListColumns("Auxiliar").Range.Column).Value
                End With
            End If
        Next d
    End If

    Call LimpiarConfigurador
    MsgBox "Comanda " & numero & " cargada. Puedes editarla y volver a 'Aceptar Presupuesto' para guardar los cambios.", vbInformation
End Sub

' =====================================================================
' ACTUALIZAR TASAS -> guarda C34/C35 en TasasCambio, sin salir de Presupuesto
' =====================================================================
Public Sub ActualizarTasas()
    Dim ws As Worksheet, wsT As Worksheet
    Set ws = ObtenerHoja()
    If ws Is Nothing Then Exit Sub

    On Error GoTo Falla
    Set wsT = ThisWorkbook.Worksheets(HOJA_TASAS)
    On Error GoTo 0

    Dim cop As Variant, ves As Variant
    cop = ws.Range("C34").Value
    ves = ws.Range("C35").Value
    If Not IsNumeric(cop) Or cop <= 0 Then
        MsgBox "La tasa COP debe ser un número mayor a 0.", vbExclamation
        Exit Sub
    End If
    If Not IsNumeric(ves) Or ves <= 0 Then
        MsgBox "La tasa VES debe ser un número mayor a 0.", vbExclamation
        Exit Sub
    End If

    wsT.Range("C4").Value = cop
    wsT.Range("C5").Value = ves
    MsgBox "Tasas actualizadas: COP = " & cop & "   VES = " & ves, vbInformation
    Exit Sub

Falla:
    MsgBox "No se encontró la hoja '" & HOJA_TASAS & "'.", vbCritical
End Sub

' =====================================================================
' Helpers privados
' =====================================================================
Private Function ProximoNumeroItem(lo As ListObject) As Long
    Dim maxItem As Long, r As Range
    maxItem = 0
    If Not lo.DataBodyRange Is Nothing Then
        For Each r In lo.ListColumns("ITEM").DataBodyRange
            If IsNumeric(r.Value) Then
                If r.Value > maxItem Then maxItem = r.Value
            End If
        Next r
    End If
    ProximoNumeroItem = maxItem + 1
End Function

Private Function BuscarListRowPorItem(lo As ListObject, numeroItem As Long) As ListRow
    Dim r As Range
    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each r In lo.ListColumns("ITEM").DataBodyRange
        If r.Value = numeroItem Then
            Set BuscarListRowPorItem = lo.ListRows(r.Row - lo.DataBodyRange.Row + 1)
            Exit Function
        End If
    Next r
End Function
