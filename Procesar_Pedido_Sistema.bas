Option Explicit

' Constantes globales de formato
Const FONT_NAME As String = "Calibri"
Const COLOR_HEADER As Long = 5855577
Const COLOR_CALCULATED As Long = 13619151
Const COLOR_TOTAL_CELL As Long = 12566463
Const IVA As Double = 1.18

'==================================================================================================
' MACRO PRINCIPAL - ASIGNAR ESTA MACRO AL BOTÓN "PROCESAR"
' Propósito: Muestra un menú para que el usuario elija el tipo de archivo a generar (PDF o XLSX)
'            y llama a la subrutina correspondiente.
'==================================================================================================
Public Sub Procesar_Pedido()
    Dim choice As VbMsgBoxResult
    Dim msg As String
    
    msg = "¿Qué tipo de archivo desea generar?" & vbNewLine & vbNewLine
    msg = msg & "▪ Sí = Carta de Cotización (PDF)" & vbNewLine
    msg = msg & "▪ No = Hoja de Pedido (Excel)"
    
    choice = MsgBox(msg, vbYesNoCancel + vbQuestion, "Seleccionar Tipo de Salida")
    
    If choice = vbCancel Then
        Exit Sub
    ElseIf choice = vbYes Then
        ' --- RUTA PDF ---
        ' Primero, verificar que los datos del vendedor en CONFIG estén completos
        On Error Resume Next
        Dim wsConfig As Worksheet
        Set wsConfig = ThisWorkbook.Sheets("CONFIG")
        On Error GoTo 0
        
        If wsConfig Is Nothing Then
            MsgBox "La hoja 'CONFIG' no se encontró. Por favor, créela y configúrela.", vbCritical
            Exit Sub
        End If
        
        If IsEmpty(wsConfig.Range("B15").Value) Or wsConfig.Range("B15").Value = "" Then
            MsgBox "Antes de generar una carta PDF, por favor configure su nombre de vendedor en la celda B15 de la hoja 'CONFIG'.", vbExclamation
            Exit Sub
        End If
        
        ' Llamar a la macro que genera el PDF
        Call GenerarCartaPDF
        
    ElseIf choice = vbNo Then
        ' --- RUTA XLSX ---
        ' Llamar a la macro que genera el Excel
        Call GenerarHojaXLSX
    End If
End Sub

'==================================================================================================
' MACRO AUXILIAR PRIVADA
' Propósito: Genera una CARTA DE COTIZACIÓN FORMAL en formato PDF.
'==================================================================================================
Private Sub GenerarCartaPDF()
    On Error GoTo ManejoDeErrores
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Definir hojas
    Dim wsPedidosSource As Worksheet, wsPedidoDest As Worksheet, wsConfig As Worksheet, wbNuevo As Workbook
    Set wsPedidosSource = ThisWorkbook.Sheets("PEDIDOS")
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")

    ' Eliminar hoja de trabajo anterior si existe
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("PEDIDO_TEMP").Delete
    On Error GoTo ManejoDeErrores
    Application.DisplayAlerts = True

    ' Crear hoja de trabajo
    Set wsPedidoDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsPedidoDest.Name = "PEDIDO_TEMP"
    
    ' 1. Construir Cabecera
    wsPedidoDest.Rows("1:9").RowHeight = 15
    Dim logoShape As Shape
    On Error Resume Next
    Set logoShape = wsConfig.Shapes("LOGO_CIP")
    On Error GoTo ManejoDeErrores
    If Not logoShape Is Nothing Then
        logoShape.Copy
        wsPedidoDest.Paste
        With wsPedidoDest.Shapes(wsPedidoDest.Shapes.Count)
            .LockAspectRatio = msoTrue: .Height = wsPedidoDest.Range("A1:A6").Height
            .Width = wsPedidoDest.Range("A1:C1").Width: .Top = wsPedidoDest.Range("A1").Top + 5
            .Left = wsPedidoDest.Range("A1").Left + 5: .Placement = xlMove
        End With
    End If
    With wsConfig
        wsPedidoDest.Range("D1:J1").Merge: wsPedidoDest.Range("D1").Value = .Range("B4").Value
        wsPedidoDest.Range("D2:J2").Merge: wsPedidoDest.Range("D2").Value = .Range("B5").Value
        wsPedidoDest.Range("D3").Value = .Range("B6").Value
        wsPedidoDest.Range("D4").Value = .Range("B7").Value
        wsPedidoDest.Range("D5").Value = .Range("B8").Value
    End With
    With wsPedidoDest.Range("D1:J5")
        .Font.Name = FONT_NAME: .Font.Size = 9: .VerticalAlignment = xlCenter
    End With
    With wsPedidoDest.Range("D1")
        .Font.Bold = True: .Font.Size = 14: .HorizontalAlignment = xlCenter
    End With
    wsPedidoDest.Range("D2").HorizontalAlignment = xlCenter
    wsPedidoDest.Range("G7:H7").Merge: wsPedidoDest.Range("G7").Value = wsPedidosSource.Range("D2").Value
    wsPedidoDest.Range("J7").Value = wsPedidosSource.Range("D3").Value
    wsPedidoDest.Range("G8:H8").Merge: wsPedidoDest.Range("G8").Value = wsConfig.Range("B15").Value
    wsPedidoDest.Range("J8").Value = wsConfig.Range("B17").Value
    With wsPedidoDest
        .Range("F7").Value = "CLIENTE:": .Range("I7").Value = "PEDIDO N°:"
        .Range("F8").Value = "VENDEDOR:": .Range("I8").Value = "CONTACTO:"
    End With
    With wsPedidoDest.Range("F7:J8")
        .Font.Name = FONT_NAME: .Font.Size = 10: .Font.Bold = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    wsPedidoDest.Range("F7:F8, I7:I8").HorizontalAlignment = xlRight
    
    ' 2. Texto de introducción de la carta
    Dim introRow As Long: introRow = 11
    wsPedidoDest.Range("A" & introRow & ":J" & introRow).Merge
    wsPedidoDest.Range("A" & introRow).Value = "Estimados Sres. " & wsPedidosSource.Range("D2").Value & ","
    wsPedidoDest.Range("A" & introRow + 1 & ":J" & introRow + 1).Merge
    wsPedidoDest.Range("A" & introRow + 1).Value = "Atendiendo a su solicitud, tenemos el agrado de presentarles la siguiente cotización:"
    wsPedidoDest.Range("A" & introRow & ":A" & introRow + 1).Font.Name = FONT_NAME
    
    ' 3. Copiar datos y crear tabla
    Dim firstDataRowDest As Long, lastRowSource As Long, dataRowCount As Long
    firstDataRowDest = 15 ' La tabla empieza más abajo por la intro
    lastRowSource = wsPedidosSource.Cells(wsPedidosSource.Rows.Count, "D").End(xlUp).Row
    If lastRowSource < 4 Then dataRowCount = 0 Else dataRowCount = lastRowSource - 4 + 1
    If dataRowCount > 0 Then
        ' Mapeo corregido de columnas desde PEDIDOS a PEDIDO_TEMP
        ' Destino            <-- Origen (en PEDIDOS)
        ' ===============================================
        ' A: CANT.           <-- Columna F
        wsPedidosSource.Range("F4:F" & lastRowSource).Copy Destination:=wsPedidoDest.Range("A" & firstDataRowDest)
        ' B: U/M             <-- Columna G
        wsPedidosSource.Range("G4:G" & lastRowSource).Copy Destination:=wsPedidoDest.Range("B" & firstDataRowDest)
        ' C: ARTICULO        <-- Columna C
        wsPedidosSource.Range("C4:C" & lastRowSource).Copy Destination:=wsPedidoDest.Range("C" & firstDataRowDest)
        ' D: DESCRIPCIÓN     <-- Columna D
        wsPedidosSource.Range("D4:D" & lastRowSource).Copy Destination:=wsPedidoDest.Range("D" & firstDataRowDest)
        ' E: V. VENTA UNIT.  <-- Columna E
        wsPedidosSource.Range("E4:E" & lastRowSource).Copy Destination:=wsPedidoDest.Range("E" & firstDataRowDest)
        ' F: DESC 1          <-- Columna H
        wsPedidosSource.Range("H4:H" & lastRowSource).Copy Destination:=wsPedidoDest.Range("F" & firstDataRowDest)
        ' G: DESC 2          <-- Columna I (Asunción)
        wsPedidosSource.Range("I4:I" & lastRowSource).Copy Destination:=wsPedidoDest.Range("G" & firstDataRowDest)
    End If
    Dim headers As Variant
    headers = Array("CANT.", "U/M", "ARTICULO", "DESCRIPCIÓN", "V. VENTA UNIT.", "DESC 1", "DESC 2", "VALOR VENTA", "PRECIO UNIT.", "PRECIO VENTA")
    Dim i As Integer
    With wsPedidoDest
        For i = LBound(headers) To UBound(headers)
            With .Cells(14, i + 1)
                .Value = headers(i): .Interior.Color = COLOR_HEADER: .Font.Color = &HFFFFFF
                .Font.Bold = True: .Font.Name = FONT_NAME: .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter: .WrapText = True
                .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
            End With
        Next i
        .Rows(14).RowHeight = 35
    End With
    Dim tbl As ListObject, lastRowDestData As Long
    lastRowDestData = wsPedidoDest.Cells(wsPedidoDest.Rows.Count, "A").End(xlUp).Row
    If lastRowDestData < 14 Then lastRowDestData = 14
    Set tbl = wsPedidoDest.ListObjects.Add(xlSrcRange, wsPedidoDest.Range("A14:J" & lastRowDestData), , xlYes)
    tbl.TableStyle = ""
    
    ' 4. Fórmulas y formato de tabla
    If dataRowCount > 0 Then
        With tbl
            .ListColumns("VALOR VENTA").DataBodyRange.FormulaR1C1 = "=[@[CANT.]]*[@[V. VENTA UNIT.]]*(1-[@[DESC 1]])*(1-[@[DESC 2]])"
            .ListColumns("PRECIO UNIT.").DataBodyRange.FormulaR1C1 = "=1*[@[V. VENTA UNIT.]]*(1-[@[DESC 1]])*(1-[@[DESC 2]])*" & Replace(IVA, ",", ".")
            .ListColumns("PRECIO VENTA").DataBodyRange.FormulaR1C1 = "=[@[VALOR VENTA]]*" & Replace(IVA, ",", ".")
            .DataBodyRange.Font.Size = 10
            .ListColumns("PRECIO VENTA").DataBodyRange.Interior.Color = COLOR_CALCULATED
        End With
    End If
    Dim formatoSoles As String: formatoSoles = "[$S/-409] #,##0.00"
    With tbl
        .ListColumns("DESC 1").Range.NumberFormat = "0.00%": .ListColumns("DESC 2").Range.NumberFormat = "0.00%"
        .ListColumns("V. VENTA UNIT.").Range.NumberFormat = formatoSoles
        .ListColumns("VALOR VENTA").Range.NumberFormat = formatoSoles
        .ListColumns("PRECIO UNIT.").Range.NumberFormat = formatoSoles
        .ListColumns("PRECIO VENTA").Range.NumberFormat = formatoSoles
    End With

    ' 5. Total y Condiciones
    wsPedidoDest.Cells(9, "I").Value = "Importe Total:"
    With wsPedidoDest.Cells(9, "J")
        If dataRowCount > 0 Then .Formula = "=SUM(" & tbl.ListColumns("PRECIO VENTA").DataBodyRange.Address(External:=False) & ")" Else .Value = 0
        .NumberFormat = formatoSoles: .Font.Bold = True: .Font.Size = 13: .Interior.Color = COLOR_TOTAL_CELL
    End With
    Dim finalRow As Long: finalRow = wsPedidoDest.Cells(wsPedidoDest.Rows.Count, "A").End(xlUp).Row + 2
    wsPedidoDest.Range("A" & finalRow).Value = "Condiciones Comerciales:": wsPedidoDest.Range("A" & finalRow).Font.Bold = True
    wsPedidoDest.Range("A" & finalRow + 1).Value = "  - " & wsConfig.Range("B20").Value
    wsPedidoDest.Range("A" & finalRow + 2).Value = "  - " & wsConfig.Range("B21").Value
    wsPedidoDest.Range("A" & finalRow + 3).Value = "  - " & wsConfig.Range("B22").Value
    finalRow = finalRow + 5
    wsPedidoDest.Range("A" & finalRow).Value = "Agradeciendo de antemano su preferencia, nos despedimos."
    wsPedidoDest.Range("A" & finalRow + 2).Value = "Atentamente,"
    
    ' 6. Configuración de página y exportación a PDF
    With wsPedidoDest.PageSetup
        .CenterFooter = wsConfig.Range("B25").Value
        .FitToPagesWide = 1: .FitToPagesTall = False: .Orientation = xlPortrait
        .PrintArea = "$A$1:$J$" & finalRow + 4
    End With
    wsPedidoDest.Columns("A:J").AutoFit
    Dim nombreArchivo As String, desktopPath As String, pdfFilePath As String
    nombreArchivo = "Cotizacion - " & LimpiarNombreArchivo(wsPedidosSource.Range("D2").Value & "-" & wsPedidosSource.Range("D3").Value)
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    pdfFilePath = desktopPath & "\" & nombreArchivo & ".pdf"
    If Dir(pdfFilePath) <> "" And MsgBox("El archivo PDF '" & nombreArchivo & ".pdf' ya existe. ¿Desea reemplazarlo?", vbYesNo + vbQuestion) = vbNo Then
        MsgBox "Operación cancelada.", vbInformation
        GoTo LimpiezaFinal
    End If
    wsPedidoDest.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfFilePath, Quality:=xlQualityStandard, OpenAfterPublish:=False
    
    If MsgBox("PDF guardado en: " & vbNewLine & pdfFilePath & vbNewLine & vbNewLine & "¿Desea abrir el archivo PDF ahora?", vbYesNo + vbQuestion, "Guardado Exitoso") = vbYes Then
        ThisWorkbook.FollowHyperlink pdfFilePath
    End If
    
    Call LimpiarHojaPedidos

LimpiezaFinal:
    Application.DisplayAlerts = False
    If Not wsPedidoDest Is Nothing Then wsPedidoDest.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
ManejoDeErrores:
    MsgBox "Ocurrió un error inesperado:" & vbNewLine & Err.Description, vbCritical, "Error en Macro"
    Resume LimpiezaFinal
End Sub


'==================================================================================================
' MACRO AUXILIAR PRIVADA
' Propósito: Genera una HOJA DE PEDIDO en formato XLSX, sin formato de carta.
'==================================================================================================
Private Sub GenerarHojaXLSX()
    On Error GoTo ManejoDeErrores
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Definir hojas
    Dim wsPedidosSource As Worksheet, wsPedidoDest As Worksheet, wsConfig As Worksheet, wbNuevo As Workbook
    Set wsPedidosSource = ThisWorkbook.Sheets("PEDIDOS")
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    
    If wsConfig Is Nothing Then MsgBox "La hoja 'CONFIG' no se encontró.", vbCritical: GoTo LimpiezaFinal
    
    ' Validar datos obligatorios en CONFIG
    If IsEmpty(wsConfig.Range("B4").Value) Or wsConfig.Range("B4").Value = "" Then
        MsgBox "La celda B4 en CONFIG está vacía. Configure el nombre de la empresa.", vbExclamation
        GoTo LimpiezaFinal
    End If
    If IsEmpty(wsConfig.Range("B15").Value) Or wsConfig.Range("B15").Value = "" Then
        MsgBox "La celda B15 en CONFIG está vacía. Configure el nombre del vendedor.", vbExclamation
        GoTo LimpiezaFinal
    End If

    ' Eliminar hoja de trabajo anterior si existe
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("PEDIDO_TEMP").Delete
    On Error GoTo ManejoDeErrores
    Application.DisplayAlerts = True

    ' Crear hoja de trabajo
    Set wsPedidoDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsPedidoDest.Name = "PEDIDO_TEMP"

    ' 1. Construir Cabecera (sin textos de carta)
    wsPedidoDest.Rows("1:9").RowHeight = 15
    Dim logoShape As Shape
    On Error Resume Next
    Set logoShape = wsConfig.Shapes("LOGO_CIP")
    On Error GoTo ManejoDeErrores
    If Not logoShape Is Nothing Then
        logoShape.Copy
        wsPedidoDest.Paste
        With wsPedidoDest.Shapes(wsPedidoDest.Shapes.Count)
            .LockAspectRatio = msoTrue: .Height = wsPedidoDest.Range("A1:A6").Height
            .Width = wsPedidoDest.Range("A1:C1").Width: .Top = wsPedidoDest.Range("A1").Top + 5
            .Left = wsPedidoDest.Range("A1").Left + 5: .Placement = xlMove
        End With
    End If
    With wsConfig
        wsPedidoDest.Range("D1:J1").Merge: wsPedidoDest.Range("D1").Value = .Range("B4").Value
        wsPedidoDest.Range("D2:J2").Merge: wsPedidoDest.Range("D2").Value = .Range("B5").Value
        wsPedidoDest.Range("D3").Value = .Range("B6").Value
        wsPedidoDest.Range("D4").Value = .Range("B7").Value
        wsPedidoDest.Range("D5").Value = .Range("B8").Value
    End With
    With wsPedidoDest.Range("D1:J5")
        .Font.Name = FONT_NAME: .Font.Size = 9: .VerticalAlignment = xlCenter
    End With
    With wsPedidoDest.Range("D1")
        .Font.Bold = True: .Font.Size = 14: .HorizontalAlignment = xlCenter
    End With
    wsPedidoDest.Range("D2").HorizontalAlignment = xlCenter
    wsPedidoDest.Range("G7:H7").Merge: wsPedidoDest.Range("G7").Value = wsPedidosSource.Range("D2").Value
    wsPedidoDest.Range("J7").Value = wsPedidosSource.Range("D3").Value
    With wsPedidoDest
        .Range("F7").Value = "CLIENTE:": .Range("I7").Value = "PEDIDO N°:"
    End With
    With wsPedidoDest.Range("F7:J7")
        .Font.Name = FONT_NAME: .Font.Size = 10: .Font.Bold = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    wsPedidoDest.Range("F7:F8, I7:I8").HorizontalAlignment = xlRight
    ' Bordes para la cabecera
    wsPedidoDest.Range("A1:J8").Borders.LineStyle = xlContinuous
    wsPedidoDest.Range("A1:J8").Borders.Weight = xlThin
    
    ' 2. Copiar datos y crear tabla (empieza antes, no hay carta)
    Dim firstDataRowDest As Long, lastRowSource As Long, dataRowCount As Long
    firstDataRowDest = 11 ' La tabla empieza justo después de la cabecera
    lastRowSource = wsPedidosSource.Cells(wsPedidosSource.Rows.Count, "D").End(xlUp).Row
    If lastRowSource < 4 Then dataRowCount = 0 Else dataRowCount = lastRowSource - 4 + 1
    If dataRowCount = 0 Then
        MsgBox "No hay datos en la hoja PEDIDOS para generar el archivo Excel.", vbExclamation
        GoTo LimpiezaFinal
    End If
    If dataRowCount > 0 Then
        ' Mapeo corregido de columnas desde PEDIDOS a PEDIDO_TEMP
        ' Destino            <-- Origen (en PEDIDOS)
        ' ===============================================
        ' A: CANT.           <-- Columna F
        wsPedidosSource.Range("F4:F" & lastRowSource).Copy Destination:=wsPedidoDest.Range("A" & firstDataRowDest)
        ' B: U/M             <-- Columna G
        wsPedidosSource.Range("G4:G" & lastRowSource).Copy Destination:=wsPedidoDest.Range("B" & firstDataRowDest)
        ' C: ARTICULO        <-- Columna C
        wsPedidosSource.Range("C4:C" & lastRowSource).Copy Destination:=wsPedidoDest.Range("C" & firstDataRowDest)
        ' D: DESCRIPCIÓN     <-- Columna D
        wsPedidosSource.Range("D4:D" & lastRowSource).Copy Destination:=wsPedidoDest.Range("D" & firstDataRowDest)
        ' E: V. VENTA UNIT.  <-- Columna E
        wsPedidosSource.Range("E4:E" & lastRowSource).Copy Destination:=wsPedidoDest.Range("E" & firstDataRowDest)
        ' F: DESC 1          <-- Columna H
        wsPedidosSource.Range("H4:H" & lastRowSource).Copy Destination:=wsPedidoDest.Range("F" & firstDataRowDest)
        ' G: DESC 2          <-- Columna I (Asunción)
        wsPedidosSource.Range("I4:I" & lastRowSource).Copy Destination:=wsPedidoDest.Range("G" & firstDataRowDest)
    End If
    Dim headers As Variant
    headers = Array("CANT.", "U/M", "ARTICULO", "DESCRIPCIÓN", "V. VENTA UNIT.", "DESC 1", "DESC 2", "VALOR VENTA", "PRECIO UNIT.", "PRECIO VENTA")
    Dim i As Integer
    With wsPedidoDest
        For i = LBound(headers) To UBound(headers)
            With .Cells(10, i + 1) ' Fila de encabezados
                .Value = headers(i): .Interior.Color = COLOR_HEADER: .Font.Color = &HFFFFFF
                .Font.Bold = True: .Font.Name = FONT_NAME: .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter: .WrapText = True
                .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
            End With
        Next i
        .Rows(10).RowHeight = 35
    End With
    Dim tbl As ListObject, lastRowDestData As Long
    lastRowDestData = wsPedidoDest.Cells(wsPedidoDest.Rows.Count, "A").End(xlUp).Row
    If lastRowDestData < 10 Then lastRowDestData = 10
    Set tbl = wsPedidoDest.ListObjects.Add(xlSrcRange, wsPedidoDest.Range("A10:J" & lastRowDestData), , xlYes)
    tbl.TableStyle = "TableStyleMedium2"

    ' 3. Fórmulas y formato de tabla
    If dataRowCount > 0 Then
        With tbl
            .ListColumns("VALOR VENTA").DataBodyRange.FormulaR1C1 = "=[@[CANT.]]*[@[V. VENTA UNIT.]]*(1-[@[DESC 1]])*(1-[@[DESC 2]])"
            .ListColumns("PRECIO UNIT.").DataBodyRange.FormulaR1C1 = "=1*[@[V. VENTA UNIT.]]*(1-[@[DESC 1]])*(1-[@[DESC 2]])*" & Replace(IVA, ",", ".")
            .ListColumns("PRECIO VENTA").DataBodyRange.FormulaR1C1 = "=[@[VALOR VENTA]]*" & Replace(IVA, ",", ".")
            .DataBodyRange.Font.Size = 10
            .ListColumns("PRECIO VENTA").DataBodyRange.Interior.Color = COLOR_CALCULATED
        End With
    End If
    Dim formatoSoles As String: formatoSoles = "[$S/-409] #,##0.00"
    With tbl
        .ListColumns("DESC 1").Range.NumberFormat = "0.00%": .ListColumns("DESC 2").Range.NumberFormat = "0.00%"
        .ListColumns("V. VENTA UNIT.").Range.NumberFormat = formatoSoles
        .ListColumns("VALOR VENTA").Range.NumberFormat = formatoSoles
        .ListColumns("PRECIO UNIT.").Range.NumberFormat = formatoSoles
        .ListColumns("PRECIO VENTA").Range.NumberFormat = formatoSoles
    End With
    
    ' 4. Total
    wsPedidoDest.Cells(9, "I").Value = "Importe Total:"
    With wsPedidoDest.Cells(9, "J")
        If dataRowCount > 0 Then .Formula = "=SUM(" & tbl.ListColumns("PRECIO VENTA").DataBodyRange.Address(External:=False) & ")" Else .Value = 0
        .NumberFormat = formatoSoles: .Font.Bold = True: .Font.Size = 13: .Interior.Color = COLOR_TOTAL_CELL
    End With

    ' 5. Guardado como XLSX
    wsPedidoDest.Columns("A:J").AutoFit
    ' Ajuste manual de anchos de columna para mejor legibilidad
    wsPedidoDest.Columns("A").ColumnWidth = 8
    wsPedidoDest.Columns("B").ColumnWidth = 6
    wsPedidoDest.Columns("C").ColumnWidth = 12
    wsPedidoDest.Columns("D").ColumnWidth = 30
    wsPedidoDest.Columns("E").ColumnWidth = 15
    wsPedidoDest.Columns("F").ColumnWidth = 8
    wsPedidoDest.Columns("G").ColumnWidth = 8
    wsPedidoDest.Columns("H").ColumnWidth = 15
    wsPedidoDest.Columns("I").ColumnWidth = 15
    wsPedidoDest.Columns("J").ColumnWidth = 15
    ' Configuración de página para impresión
    With wsPedidoDest.PageSetup
        .Orientation = xlLandscape
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
    End With
    
    ' Copiar la hoja a un nuevo libro. Este es el método más robusto.
    wsPedidoDest.Copy
    Set wbNuevo = ActiveWorkbook ' El nuevo libro se convierte en el activo
    
    Dim nombreArchivo As String, desktopPath As String, xlsxFilePath As String
    nombreArchivo = "Pedido - " & LimpiarNombreArchivo(wbNuevo.Sheets(1).Range("G7").Value & "-" & wbNuevo.Sheets(1).Range("J7").Value)
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    xlsxFilePath = desktopPath & "\" & nombreArchivo & ".xlsx"
    
    If Dir(xlsxFilePath) <> "" Then
        If MsgBox("El archivo Excel '" & nombreArchivo & ".xlsx' ya existe en el escritorio. ¿Desea reemplazarlo?", vbYesNo + vbQuestion, "Confirmar Reemplazo") = vbNo Then
            MsgBox "Operación cancelada por el usuario.", vbInformation
            wbNuevo.Close SaveChanges:=False
            GoTo LimpiezaFinal
        End If
    End If
    
    ' Guardar el nuevo libro
    Application.DisplayAlerts = False ' Evitar el prompt de "sobrescribir" de Excel
    On Error Resume Next
    wbNuevo.SaveAs fileName:=xlsxFilePath, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        wbNuevo.Close SaveChanges:=False
        Application.DisplayAlerts = True
        MsgBox "No se pudo guardar el archivo:" & vbNewLine & Err.Description, vbCritical, "Error al Guardar"
        GoTo LimpiezaFinal
    End If
    On Error GoTo ManejoDeErrores
    Application.DisplayAlerts = True
    
    wbNuevo.Close SaveChanges:=False ' Se cierra el libro en memoria, el archivo ya está guardado
    
    MsgBox "Archivo Excel guardado en: " & vbNewLine & xlsxFilePath, vbInformation, "Guardado Exitoso"
    
    If MsgBox("¿Desea abrir el archivo Excel ahora?", vbYesNo + vbQuestion) = vbYes Then
        ThisWorkbook.FollowHyperlink xlsxFilePath
    End If
    
    Call LimpiarHojaPedidos

LimpiezaFinal:
    Application.DisplayAlerts = False
    If Not wsPedidoDest Is Nothing Then wsPedidoDest.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
ManejoDeErrores:
    MsgBox "Ocurrió un error inesperado:" & vbNewLine & Err.Description, vbCritical, "Error en Macro"
    Resume LimpiezaFinal
End Sub


'==================================================================================================
' FUNCIONES AUXILIARES
'==================================================================================================
Function LimpiarNombreArchivo(nombre As String) As String
    Dim caracteresNoPermitidos As String: caracteresNoPermitidos = "/\[]:*?<>|"""
    Dim i As Integer
    For i = 1 To Len(caracteresNoPermitidos)
        nombre = Replace(nombre, Mid(caracteresNoPermitidos, i, 1), "_")
    Next i
    LimpiarNombreArchivo = nombre
End Function

Sub LimpiarHojaPedidos()
    ' Solicita confirmación antes de limpiar la hoja de pedidos.
    If MsgBox("¿Está seguro de que desea limpiar los datos en la hoja PEDIDOS? Esta acción no se puede deshacer.", vbYesNo + vbQuestion, "Confirmar Limpieza") = vbNo Then Exit Sub
    
    With ThisWorkbook.Sheets("PEDIDOS")
        .Range("D2:D3").ClearContents
        If .Cells(.Rows.Count, "A").End(xlUp).Row > 4 Then
            .Range("A4:H" & .Cells(.Rows.Count, "A").End(xlUp).Row).ClearContents
        End If
        .Activate
        .Range("A4").Select
    End With
End Sub
