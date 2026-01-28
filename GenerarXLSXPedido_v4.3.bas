Option Explicit

'==================================================================================
' MACRO: GenerarXLSXPedido_v4.3
'==================================================================================
' Propósito: Crea una hoja "PEDIDO" formateada a partir de datos en "PEDIDOS",
'           con tabla, cálculos, imagen en A1:A3 y total superior.
' Entrada: Hoja "PEDIDOS" con datos en formato específico.
'         Logotipo y datos de empresa desde la hoja "CONFIG".
' Salida: Nuevo libro XLSX guardado en el escritorio con la hoja "PEDIDO".
' Versión: 4.3 (Enero 2026) - Unificación de versión con GenerarXLSXCarta
'==================================================================================

' Constantes globales
Private Const FONT_NAME As String = "Calibri"
Private Const COLOR_HEADER As Long = 5855577       ' Azul grisáceo oscuro (#595959)
Private Const COLOR_TOP_AREA As Long = 15461355    ' Gris muy claro (#EBEBEB)
Private Const COLOR_CALCULATED As Long = 13619151  ' Gris claro con tono azulado (#CFD5EA)
Private Const COLOR_TOTAL_CELL As Long = 12566463  ' Azul medio claro (#BFDFFF)
Private Const COLOR_TEXT As Long = 0               ' Negro
Private Const COLOR_TEXT_HEADER As Long = 16777215 ' Blanco (#FFFFFF)
Private Const COLOR_TEXT_TOP As Long = 4210752     ' Gris oscuro (#404040)
Private Const COLOR_INDEX As Long = 13421772       ' Gris muy claro (#D3D3D3)
Private Const IVA As Double = 1.18

'==================================================================================
' PROCEDIMIENTO PRINCIPAL
'==================================================================================
Public Sub GenerarXLSXPedido_v4_3()
    On Error GoTo ManejoDeErrores
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Definir hojas
    Dim wsPedidosSource As Worksheet
    Dim wsPedidoDest As Worksheet
    Dim wsConfig As Worksheet
    Dim wbNuevo As Workbook
    
    Set wsPedidosSource = ThisWorkbook.Sheets("PEDIDOS")
    
    ' Variables para moneda
    Dim simboloMoneda As String

    ' Verificar existencia de la hoja CONFIG
    If Not WorksheetExists("CONFIG") Then
        MsgBox "La hoja 'CONFIG' no se encontró en este libro.", vbCritical, "Error"
        GoTo LimpiezaFinal
    End If
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    
    ' Leer símbolo de moneda desde CONFIG!B26
    simboloMoneda = Trim(wsConfig.Range("B26").Value)
    ' Establecer símbolo predeterminado si está vacío
    If simboloMoneda = "" Then
        simboloMoneda = "S/. "  ' Soles peruanos por defecto
    End If

    ' Verificar existencia de la hoja PEDIDOS
    If Not WorksheetExists("PEDIDOS") Then
        MsgBox "La hoja 'PEDIDOS' no se encontró en este libro.", vbCritical, "Error"
        GoTo LimpiezaFinal
    End If

    ' Determinar rango de datos
    Dim firstDataRowSource As Long
    Dim lastRowSource As Long
    firstDataRowSource = 5  ' Datos desde fila 5 (fila 4 puede tener encabezados)
    lastRowSource = wsPedidosSource.Cells(wsPedidosSource.Rows.Count, "C").End(xlUp).Row
    
    ' Validar que existan datos en la hoja PEDIDOS
    If lastRowSource < firstDataRowSource Then
        MsgBox "No se encontraron datos en la hoja PEDIDOS." & vbNewLine & vbNewLine & _
               "Por favor, asegúrese de pegar los datos en la hoja PEDIDOS antes de ejecutar la macro." & vbNewLine & _
               "Los datos deben comenzar en la fila 5.", vbExclamation, "Sin Datos"
        GoTo LimpiezaFinal
    End If

    ' Eliminar hoja PEDIDO si existe
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("PEDIDO").Delete
    On Error GoTo ManejoDeErrores
    Application.DisplayAlerts = True

    ' Crear nueva hoja PEDIDO
    Set wsPedidoDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsPedidoDest.Name = "PEDIDO"

    ' Insertar logotipo desde la hoja CONFIG (ocupa exactamente 3 filas de alto: A1:A3)
    Dim logoShape As Shape
    On Error Resume Next
    Set logoShape = wsConfig.Shapes("logo_empresa")
    On Error GoTo ManejoDeErrores
    
    If Not logoShape Is Nothing Then
        logoShape.Copy
        wsPedidoDest.Paste
        With wsPedidoDest.Shapes(wsPedidoDest.Shapes.Count)
            ' Pegar en A1
            .Top = wsPedidoDest.Range("A1").Top
            .Left = wsPedidoDest.Range("A1").Left
            ' Ajustar tamaño con altura fija de 2.10 cm (aproximadamente 60 puntos)
            .Height = 60
            .Placement = xlMove  ' El logo se mueve con las celdas
        End With
    Else
        MsgBox "El logotipo 'logo_empresa' no se encontró en la hoja CONFIG. La hoja se generará sin logotipo.", vbExclamation
    End If

    ' C1:E1: Nombre de empresa desde CONFIG!B6 (celdas combinadas)
    wsPedidoDest.Range("C1:E1").Merge
    wsPedidoDest.Range("C1").Value = wsConfig.Range("B6").Value
    With wsPedidoDest.Range("C1")
        .Font.Bold = True
        .Font.Size = 16
        .Font.Name = FONT_NAME
        .Font.Color = RGB(0, 51, 102)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' Copiar información de cliente y pedido (desde D2 y D3, combinar en destino)
    With wsPedidosSource
        ' Cliente desde D2 (solo celda D2) - pegar en D2:E2 combinado
        .Range("D2").Copy Destination:=wsPedidoDest.Range("D2")
        wsPedidoDest.Range("D2:E2").Merge
        wsPedidoDest.Range("D2").NumberFormat = "@"  ' Formato texto
        
        ' Pedido desde D3 (solo celda D3) - pegar en D3:E3 combinado
        .Range("D3").Copy Destination:=wsPedidoDest.Range("D3")
        wsPedidoDest.Range("D3:E3").Merge
        wsPedidoDest.Range("D3").NumberFormat = "@"  ' Formato texto
    End With

    ' Formatear zona superior
    With wsPedidoDest.Range("C2:D3")
        .Borders.LineStyle = xlNone
        .Font.Color = COLOR_TEXT_TOP
        .Font.Bold = True
        .Font.Size = 12
        .Font.Name = FONT_NAME
        .Interior.Color = COLOR_TOP_AREA
        .VerticalAlignment = xlCenter
    End With

    ' Etiquetas CLIENTE y PEDIDO
    With wsPedidoDest
        .Cells(2, 3).Value = "CLIENTE"
        .Cells(2, 3).Font.Bold = True
        .Cells(2, 3).Font.Size = 11
        .Cells(2, 3).Font.Name = FONT_NAME
        .Cells(2, 3).Font.Color = COLOR_TEXT
        .Cells(2, 3).Interior.Color = COLOR_TOP_AREA
        .Cells(2, 3).VerticalAlignment = xlCenter

        .Cells(3, 3).Value = "PEDIDO"
        .Cells(3, 3).Font.Bold = True
        .Cells(3, 3).Font.Size = 11
        .Cells(3, 3).Font.Name = FONT_NAME
        .Cells(3, 3).Font.Color = COLOR_TEXT
        .Cells(3, 3).Interior.Color = COLOR_TOP_AREA
        .Cells(3, 3).VerticalAlignment = xlCenter
    End With

    ' Copiar datos de artículos con optimización y funcionalidad stock
    Dim firstDataRowDest As Long
    Dim dataRowCount As Long
    firstDataRowDest = 6  ' Encabezado en fila 5, datos desde fila 6
    If lastRowSource < firstDataRowSource Then
        dataRowCount = 0
    Else
        dataRowCount = lastRowSource - firstDataRowSource + 1
    End If

    If dataRowCount > 0 Then
        ' Optimización: Leer datos en array para procesamiento rápido
        Dim dataSource As Variant
        Dim dataDest As Variant
        ReDim dataDest(1 To dataRowCount, 1 To 12) ' 12 columnas en la tabla (A-L)

        ' Leer datos de origen en array
        dataSource = wsPedidosSource.Range("C" & firstDataRowSource & ":J" & lastRowSource).Value

        ' Procesar datos en array (más rápido que Range)
        Dim i As Long
        For i = 1 To dataRowCount
            ' Columna A: N° (Índice)
            dataDest(i, 1) = i

            ' Columna B: CANTIDAD (columna E del sistema RPE → índice 3 en array dataSource)
            dataDest(i, 2) = dataSource(i, 3)

            ' Columna C: U/M (columna G del sistema RPE → índice 5 en array dataSource)
            dataDest(i, 3) = dataSource(i, 5)

            ' Columna D: ARTICULO (columna C del sistema RPE → índice 1 en array dataSource)
            ' Convertir a texto explícitamente para preservar ceros a la izquierda
            dataDest(i, 4) = CStr(dataSource(i, 1))

            ' Columna E: DESCRIPCIÓN (columna D del sistema RPE → índice 2 en array dataSource)
            dataDest(i, 5) = dataSource(i, 2)

            ' Columna F: STOCK (comparar stock vs cantidad)
            Dim stockValue As Double
            Dim qtyValue As Double

            ' Stock (columna F del sistema RPE → índice 4 en array dataSource)
            If IsNumeric(dataSource(i, 4)) Then
                stockValue = CDbl(dataSource(i, 4))
            Else
                stockValue = 0
            End If

            ' Cantidad (columna E del sistema RPE → índice 3 en array dataSource)
            If IsNumeric(dataSource(i, 3)) Then
                qtyValue = CDbl(dataSource(i, 3))
            Else
                qtyValue = 0
            End If

            ' Determinar estado de stock (solo informativo, no modifica cantidades)
            If stockValue = 0 Then
                dataDest(i, 6) = "Sin Stock"
            ElseIf stockValue < qtyValue Then
                dataDest(i, 6) = "Stock Insuficiente"
            ElseIf stockValue >= qtyValue And stockValue <= qtyValue * 1.1 Then
                dataDest(i, 6) = "Stock Ajustado"
            Else
                dataDest(i, 6) = "Stock Disponible"
            End If
            ' Las cantidades se mantienen como solicitadas originalmente (stock solo informativo)

            ' Columna G: VALOR VENTA UNITARIO (columna H del sistema RPE → índice 6 en array dataSource)
            dataDest(i, 7) = dataSource(i, 6)

            ' Columna H: DESC 1 (columna I del sistema RPE → índice 7 en array dataSource)
            Dim desc1 As Double
            If IsNumeric(dataSource(i, 7)) Then
                desc1 = CDbl(dataSource(i, 7))
                If desc1 > 1 Then desc1 = desc1 / 100
            Else
                desc1 = 0
            End If
            dataDest(i, 8) = desc1

            ' Columna I: DESC 2 (columna J del sistema RPE → índice 8 en array dataSource)
            Dim desc2 As Double
            If IsNumeric(dataSource(i, 8)) Then
                desc2 = CDbl(dataSource(i, 8))
                If desc2 > 1 Then desc2 = desc2 / 100
            Else
                desc2 = 0
            End If
            dataDest(i, 9) = desc2

            ' Columna J: VALOR VENTA (calculado)
            Dim unitValue As Double
            If IsNumeric(dataDest(i, 7)) Then
                unitValue = CDbl(dataDest(i, 7))
            Else
                unitValue = 0
            End If
            dataDest(i, 10) = dataDest(i, 2) * unitValue * (1 - dataDest(i, 8)) * (1 - dataDest(i, 9))

            ' Columna K: PRECIO UNITARIO (IVA incluido)
            dataDest(i, 11) = unitValue * (1 - dataDest(i, 8)) * (1 - dataDest(i, 9)) * IVA

            ' Columna L: PRECIO VENTA (total con IVA)
            dataDest(i, 12) = dataDest(i, 10) * IVA
        Next i

        ' IMPORTANTE: Aplicar formato de texto ANTES de escribir los datos
        ' Esto preserva los ceros a la izquierda en códigos SKU
        ' Columna D: ARTICULO (texto - preserva ceros a la izquierda)
        wsPedidoDest.Range("D" & firstDataRowDest & ":D" & (firstDataRowDest + dataRowCount - 1)).NumberFormat = "@"
        ' Columna C: U/M (texto)
        wsPedidoDest.Range("C" & firstDataRowDest & ":C" & (firstDataRowDest + dataRowCount - 1)).NumberFormat = "@"
        ' Columna E: DESCRIPCIÓN (texto)
        wsPedidoDest.Range("E" & firstDataRowDest & ":E" & (firstDataRowDest + dataRowCount - 1)).NumberFormat = "@"
        ' Columna F: STOCK (texto)
        wsPedidoDest.Range("F" & firstDataRowDest & ":F" & (firstDataRowDest + dataRowCount - 1)).NumberFormat = "@"
        
        ' Escribir datos procesados en hoja (en bloque para mejor rendimiento)
        ' Las celdas ya están formateadas como texto, por lo que los valores se mantendrán como texto
        wsPedidoDest.Range("A" & firstDataRowDest & ":L" & (firstDataRowDest + dataRowCount - 1)).Value = dataDest
    End If

    ' Crear encabezados de tabla con saltos de línea (fila 5)
    Dim headers As Variant
    headers = Array("N°", "CANT.", "U/M", "ARTICULO", "DESCRIPCIÓN", "STOCK", "VALOR" & vbLf & "VENTA" & vbLf & "UNITARIO", "DESC" & vbLf & "1", "DESC" & vbLf & "2", "VALOR" & vbLf & "VENTA", "PRECIO" & vbLf & "UNITARIO", "PRECIO" & vbLf & "VENTA")
    
    With wsPedidoDest
        For i = LBound(headers) To UBound(headers)
            With .Cells(5, i + 1)
                .Value = headers(i)
                .Interior.Color = COLOR_HEADER
                .Font.Color = COLOR_TEXT_HEADER  ' Texto blanco para encabezados
                .Font.Bold = True
                .Font.Size = 10
                .Font.Name = FONT_NAME
                .HorizontalAlignment = xlCenter  ' Centrado horizontal
                .VerticalAlignment = xlCenter
                .WrapText = True  ' Habilitar ajuste de texto para mostrar saltos de línea

                ' Mejores bordes para los encabezados (más finos y elegantes)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Borders.Color = RGB(180, 180, 180)

                ' Añadir un sutil efecto de sombreado para dar profundidad
                .Interior.Pattern = xlSolid
            End With
        Next i
        .Rows(5).RowHeight = 45  ' Aumentado para acomodar múltiples líneas

        ' Destacar visualmente las columnas más importantes
        .Cells(5, 10).Interior.Color = RGB(75, 75, 75)  ' VALOR VENTA (un poco más oscuro)
        .Cells(5, 12).Interior.Color = RGB(75, 75, 75)  ' PRECIO VENTA (un poco más oscuro)
        .Cells(5, 6).Interior.Color = RGB(100, 150, 100)  ' STOCK (verde oscuro para seriedad)
    End With

    ' Crear tabla (encabezados en fila 5, datos desde fila 6)
    Dim tbl As ListObject
    Dim lastRowDestData As Long
    lastRowDestData = wsPedidoDest.Cells(wsPedidoDest.Rows.Count, "A").End(xlUp).Row
    If lastRowDestData < 5 Then lastRowDestData = 5
    
    Dim tableRange As Range
    If dataRowCount = 0 Then
        Set tableRange = wsPedidoDest.Range("A5:L5")
    Else
        Set tableRange = wsPedidoDest.Range("A5:L" & lastRowDestData)
    End If
    
    Set tbl = wsPedidoDest.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    tbl.TableStyle = "TableStyleMedium2"

    ' Personalizar estilo de la tabla para mantener el color de encabezado
    With tbl.HeaderRowRange
        .Interior.Color = COLOR_HEADER
        .Font.Color = COLOR_TEXT_HEADER
        .Font.Bold = True
    End With

    ' Mejora visual para toda la tabla
    With tbl
        .TableStyle = ""  ' Eliminar estilo predeterminado para aplicar formato personalizado
        .ShowTableStyleRowStripes = True  ' Filas alternas para mejor legibilidad
        .ShowTableStyleColumnStripes = False
        .ShowTableStyleFirstColumn = True  ' Resaltar primera columna (índice)
        .ShowTableStyleLastColumn = False
    End With

    ' Formatear cuerpo de tabla
    With tbl
        .DataBodyRange.Font.Size = 10
        .DataBodyRange.Font.Bold = False
        .DataBodyRange.Font.Name = FONT_NAME
        If dataRowCount > 0 And .ListColumns.Count >= 12 Then
            ' Destacar columnas importantes con colores sutiles
            .ListColumns(11).DataBodyRange.Interior.Color = COLOR_CALCULATED
            .ListColumns(12).DataBodyRange.Interior.Color = COLOR_CALCULATED
            .ListColumns(10).DataBodyRange.Interior.Color = RGB(242, 242, 242)  ' Gris muy suave
            .ListColumns(1).DataBodyRange.Interior.Color = COLOR_INDEX

            ' Alineación para mejor legibilidad
            .ListColumns(1).DataBodyRange.HorizontalAlignment = xlCenter
            .ListColumns(2).DataBodyRange.HorizontalAlignment = xlRight
            .ListColumns(6).DataBodyRange.HorizontalAlignment = xlRight
            .ListColumns(7).DataBodyRange.HorizontalAlignment = xlRight
            .ListColumns(8).DataBodyRange.HorizontalAlignment = xlRight
            .ListColumns(9).DataBodyRange.HorizontalAlignment = xlRight
            .ListColumns(10).DataBodyRange.HorizontalAlignment = xlRight
            .ListColumns(11).DataBodyRange.HorizontalAlignment = xlRight
            .ListColumns(12).DataBodyRange.HorizontalAlignment = xlRight

            ' Añadir bordes sutiles para mejorar separación visual
            .DataBodyRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .DataBodyRange.Borders(xlInsideHorizontal).Weight = xlThin
            .DataBodyRange.Borders(xlInsideHorizontal).Color = RGB(220, 220, 220)

            ' Formato condicional para la columna STOCK
            If .ListColumns.Count >= 6 Then
                Dim stockRange As Range
                Set stockRange = .ListColumns(6).DataBodyRange
                Dim condFormat As FormatCondition

                ' Rojo oscuro si "Sin Stock"
                Set condFormat = stockRange.FormatConditions.Add(xlCellValue, xlEqual, "Sin Stock")
                With condFormat
                    .Interior.Color = RGB(255, 150, 150)
                    .Font.Color = RGB(100, 0, 0)
                    .Font.Bold = True
                End With

                ' Rojo si "Stock Insuficiente"
                Set condFormat = stockRange.FormatConditions.Add(xlCellValue, xlEqual, "Stock Insuficiente")
                With condFormat
                    .Interior.Color = RGB(255, 200, 200)
                    .Font.Color = RGB(150, 50, 50)
                    .Font.Bold = True
                End With

                ' Amarillo si "Stock Ajustado"
                Set condFormat = stockRange.FormatConditions.Add(xlCellValue, xlEqual, "Stock Ajustado")
                With condFormat
                    .Interior.Color = RGB(255, 255, 180)
                    .Font.Color = RGB(150, 150, 0)
                    .Font.Bold = True
                End With

                ' Verde si "Stock Disponible"
                Set condFormat = stockRange.FormatConditions.Add(xlCellValue, xlEqual, "Stock Disponible")
                With condFormat
                    .Interior.Color = RGB(180, 255, 180)
                    .Font.Color = RGB(0, 150, 0)
                    .Font.Bold = True
                End With
            End If
            
            ' Formato condicional para la columna CANT. (resaltar cantidades 0 o negativas)
            If .ListColumns.Count >= 2 Then
                Dim qtyRange As Range
                Set qtyRange = .ListColumns(2).DataBodyRange
                
                ' Rojo oscuro si cantidad es 0
                Set condFormat = qtyRange.FormatConditions.Add(xlCellValue, xlEqual, 0)
                With condFormat
                    .Interior.Color = RGB(255, 200, 200)
                    .Font.Color = RGB(150, 0, 0)
                    .Font.Bold = True
                End With
                
                ' Rojo más oscuro si cantidad es negativa
                Set condFormat = qtyRange.FormatConditions.Add(xlCellValue, xlLess, 0)
                With condFormat
                    .Interior.Color = RGB(255, 150, 150)
                    .Font.Color = RGB(100, 0, 0)
                    .Font.Bold = True
                End With
            End If
        End If
    End With

    ' Aplicar formato numérico
    Dim formatoMoneda As String
    ' Determinar formato de moneda según símbolo
    Select Case Left(Trim(simboloMoneda), 1)
        Case "S"
            formatoMoneda = "[$S/-409] #,##0.00"  ' Soles peruanos
        Case "$"
            formatoMoneda = "[$$-409] #,##0.00"  ' Dólares estadounidenses
        Case "€"
            formatoMoneda = "[$€-2] #,##0.00"    ' Euros
        Case Else
            ' Para otras monedas, usar formato genérico con símbolo
            formatoMoneda = simboloMoneda & " #,##0.00"
    End Select
    
    If tbl.ListColumns.Count >= 12 Then
        With tbl
            .ListColumns(1).Range.NumberFormat = "0"  ' N° (índice)
            .ListColumns(2).Range.NumberFormat = "0"  ' CANT.
            .ListColumns(3).Range.NumberFormat = "@"  ' U/M (texto)
            .ListColumns(4).Range.NumberFormat = "@"  ' ARTICULO (texto - preserva ceros a la izquierda)
            .ListColumns(5).Range.NumberFormat = "@"  ' DESCRIPCIÓN (texto)
            .ListColumns(6).Range.NumberFormat = "@"  ' STOCK (texto)
            .ListColumns(7).Range.NumberFormat = formatoMoneda  ' VALOR VENTA UNITARIO
            .ListColumns(8).Range.NumberFormat = "0.00%"  ' DESC 1
            .ListColumns(9).Range.NumberFormat = "0.00%"  ' DESC 2
            .ListColumns(10).Range.NumberFormat = formatoMoneda  ' VALOR VENTA
            .ListColumns(11).Range.NumberFormat = formatoMoneda  ' PRECIO UNITARIO
            .ListColumns(12).Range.NumberFormat = formatoMoneda  ' PRECIO VENTA
        End With
    End If

    ' Agregar fórmulas interactivas en las columnas calculadas (al final del proceso)
    If dataRowCount > 0 Then
        On Error Resume Next
        Application.Calculation = xlCalculationAutomatic  ' Habilitar cálculo automático temporalmente

        ' VALOR VENTA = CANT * VALOR UNITARIO * (1-DESC1) * (1-DESC2)
        tbl.ListColumns(10).DataBodyRange.Formula = "=[@[CANT.]]*[@[VALOR" & vbLf & "VENTA" & vbLf & "UNITARIO]]*(1-[@[DESC" & vbLf & "1]])*(1-[@[DESC" & vbLf & "2]])"

        ' PRECIO UNITARIO = VALOR UNITARIO * (1-DESC1) * (1-DESC2) * IVA
        tbl.ListColumns(11).DataBodyRange.Formula = "=[@[VALOR" & vbLf & "VENTA" & vbLf & "UNITARIO]]*(1-[@[DESC" & vbLf & "1]])*(1-[@[DESC" & vbLf & "2]])*" & Replace(IVA, ",", ".")

        ' PRECIO VENTA = VALOR VENTA * IVA
        tbl.ListColumns(12).DataBodyRange.Formula = "=[@[VALOR" & vbLf & "VENTA]]*" & Replace(IVA, ",", ".")

        Application.Calculation = xlCalculationManual  ' Restaurar cálculo manual
        On Error GoTo ManejoDeErrores
    End If

    ' Agregar totales superiores (3 tipos: Total con Stock, Total General, Total Descuentos)
    ' Fila 1: Total con Stock
    wsPedidoDest.Cells(1, "K").Value = "Total con Stock:"
    With wsPedidoDest.Cells(1, "K")
        .Font.Bold = True
        .Font.Size = 11
        .Font.Name = FONT_NAME
        .Font.Color = RGB(0, 102, 0)  ' Verde para stock disponible
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Interior.Color = COLOR_TOP_AREA
    End With

    ' Fila 2: Total General
    wsPedidoDest.Cells(2, "K").Value = "Total General:"
    With wsPedidoDest.Cells(2, "K")
        .Font.Bold = True
        .Font.Size = 11
        .Font.Name = FONT_NAME
        .Font.Color = RGB(0, 51, 102)  ' Azul para total general
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Interior.Color = COLOR_TOP_AREA
    End With

    ' Fila 3: Total Descuentos
    wsPedidoDest.Cells(3, "K").Value = "Total Descuentos:"
    With wsPedidoDest.Cells(3, "K")
        .Font.Bold = True
        .Font.Size = 11
        .Font.Name = FONT_NAME
        .Font.Color = RGB(102, 0, 0)  ' Rojo para descuentos
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Interior.Color = COLOR_TOP_AREA
    End With

    ' Calcular y formatear los valores de los totales
    If dataRowCount > 0 Then
        On Error Resume Next
        ' Total con Stock (productos con Stock Disponible o Stock Ajustado)
        With wsPedidoDest.Cells(1, "L")
            .Formula = "=SUMPRODUCT(((F6:F" & (5 + dataRowCount) & "=""Stock Disponible"")+(F6:F" & (5 + dataRowCount) & "=""Stock Ajustado""))*(L6:L" & (5 + dataRowCount) & "))"
            .NumberFormat = formatoMoneda
            .Font.Bold = True
            .Font.Size = 13
            .Font.Name = FONT_NAME
            .Font.Color = RGB(0, 102, 0)
            .Interior.Color = RGB(220, 255, 220)  ' Verde claro
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThick
                .Color = RGB(0, 102, 0)
            End With
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlRight
        End With

        ' Total General (todos los productos)
        With wsPedidoDest.Cells(2, "L")
            .Formula = "=SUM(L6:L" & (5 + dataRowCount) & ")"
            .NumberFormat = formatoMoneda
            .Font.Bold = True
            .Font.Size = 13
            .Font.Name = FONT_NAME
            .Font.Color = RGB(0, 51, 102)
            .Interior.Color = COLOR_TOTAL_CELL
            With .Borders
                .LineStyle = xlDouble
                .Weight = xlThick
                .Color = RGB(0, 51, 102)
            End With
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlRight
        End With

        ' Total Descuentos (suma de descuentos aplicados por fila)
        With wsPedidoDest.Cells(3, "L")
            .Formula = "=SUMPRODUCT(G6:G" & (5 + dataRowCount) & "*B6:B" & (5 + dataRowCount) & "*(H6:H" & (5 + dataRowCount) & "+I6:I" & (5 + dataRowCount) & "))"
            .NumberFormat = formatoMoneda
            .Font.Bold = True
            .Font.Size = 13
            .Font.Name = FONT_NAME
            .Font.Color = RGB(102, 0, 0)
            .Interior.Color = RGB(255, 220, 220)  ' Rojo claro
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThick
                .Color = RGB(102, 0, 0)
            End With
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlRight
        End With
        On Error GoTo ManejoDeErrores
    Else
        ' Totales en cero cuando no hay datos
        With wsPedidoDest.Cells(1, "L")
            .Value = 0
            .NumberFormat = formatoMoneda
            .Font.Bold = True
            .Font.Size = 13
            .Font.Name = FONT_NAME
            .Font.Color = RGB(0, 102, 0)
            .Interior.Color = RGB(220, 255, 220)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThick
                .Color = RGB(0, 102, 0)
            End With
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlRight
        End With

        With wsPedidoDest.Cells(2, "L")
            .Value = 0
            .NumberFormat = formatoMoneda
            .Font.Bold = True
            .Font.Size = 13
            .Font.Name = FONT_NAME
            .Font.Color = RGB(0, 51, 102)
            .Interior.Color = COLOR_TOTAL_CELL
            With .Borders
                .LineStyle = xlDouble
                .Weight = xlThick
                .Color = RGB(0, 51, 102)
            End With
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlRight
        End With

        With wsPedidoDest.Cells(3, "L")
            .Value = 0
            .NumberFormat = formatoMoneda
            .Font.Bold = True
            .Font.Size = 13
            .Font.Name = FONT_NAME
            .Font.Color = RGB(102, 0, 0)
            .Interior.Color = RGB(255, 220, 220)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThick
                .Color = RGB(102, 0, 0)
            End With
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlRight
        End With
    End If

    ' Ajustar altura de filas de totales
    wsPedidoDest.Rows(1).RowHeight = 20
    wsPedidoDest.Rows(2).RowHeight = 20
    wsPedidoDest.Rows(3).RowHeight = 20

    ' Congelar paneles en la fila 5
    On Error Resume Next
    wsPedidoDest.Activate
    ActiveWindow.View = xlNormalView ' Asegurar vista normal
    wsPedidoDest.Range("A6").Select ' Seleccionar A6 para congelar las filas 1-5
    ActiveWindow.FreezePanes = True
    If Err.Number <> 0 Then
        MsgBox "No se pudieron congelar los paneles. Continuando sin congelar." & vbNewLine & _
               "Error " & Err.Number & ": " & Err.Description, vbExclamation
        Err.Clear
    End If
    On Error GoTo ManejoDeErrores

    ' Ajustar ancho de columnas
    wsPedidoDest.Columns.AutoFit

    ' Guardar en nuevo libro
    Set wbNuevo = Application.Workbooks.Add
    wsPedidoDest.Copy Before:=wbNuevo.Sheets(1)
    
    Dim nombreArchivo As String
    nombreArchivo = LimpiarNombreArchivo(wsPedidoDest.Range("D2").Value & "-" & wsPedidoDest.Range("D3").Value & ".xlsx")
    
    Dim desktopPath As String
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    
    Dim fullFilePath As String
    fullFilePath = desktopPath & "\" & nombreArchivo

    If Dir(fullFilePath) <> "" Then
        If MsgBox("El archivo '" & nombreArchivo & "' ya existe en el escritorio. ¿Desea reemplazarlo?", vbYesNo + vbQuestion, "Confirmar Guardar") = vbNo Then
            MsgBox "Operación de guardar cancelada.", vbInformation
            wbNuevo.Close SaveChanges:=False
            GoTo LimpiezaFinal
        End If
    End If
    
    On Error Resume Next
    wbNuevo.SaveAs Filename:=fullFilePath, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        MsgBox "No se pudo guardar el archivo en " & fullFilePath & "." & vbNewLine & _
               "Error " & Err.Number & ": " & Err.Description, vbCritical
        wbNuevo.Close SaveChanges:=False
        GoTo LimpiezaFinal
    End If
    On Error GoTo ManejoDeErrores
    
    wbNuevo.Close SaveChanges:=False
    
    ' ELIMINAR hoja PEDIDO del libro base después de guardar exitosamente
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("PEDIDO").Delete
    On Error GoTo ManejoDeErrores
    Application.DisplayAlerts = True

    ' Opciones post-guardado
    Dim response As VbMsgBoxResult
    Dim msgText As String
    msgText = "El archivo se ha guardado correctamente en:" & vbNewLine & fullFilePath & vbNewLine & vbNewLine & _
              "¿Qué desea hacer ahora?" & vbNewLine & vbNewLine & _
              "Sí: Abrir la carpeta contenedora" & vbNewLine & _
              "No: Abrir el archivo guardado" & vbNewLine & _
              "Cancelar: Cerrar sin acción adicional"
    response = MsgBox(msgText, vbYesNoCancel + vbQuestion, "Archivo Guardado - v4.3")
    
    Select Case response
        Case vbYes
            ' Abrir carpeta usando Shell con explorer.exe
            On Error Resume Next
            Shell "explorer.exe /select,""" & fullFilePath & """", vbNormalFocus
            If Err.Number <> 0 Then
                MsgBox "No se pudo abrir la carpeta. Error: " & Err.Description, vbExclamation
            End If
            On Error GoTo ManejoDeErrores
            
        Case vbNo
            ' Abrir el archivo
            On Error Resume Next
            Workbooks.Open Filename:=fullFilePath
            If Err.Number <> 0 Then
                MsgBox "No se pudo abrir el archivo. Error: " & Err.Description, vbExclamation
            End If
            On Error GoTo ManejoDeErrores
            
        Case vbCancel
            ' No hacer nada adicional
    End Select

LimpiezaFinal:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ManejoDeErrores:
    MsgBox "Ocurrió un error: " & Err.Number & " - " & Err.Description, vbCritical, "Error"
    Resume LimpiezaFinal
End Sub

'==================================================================================
' FUNCIÓN: WorksheetExists
'==================================================================================
Private Function WorksheetExists(sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function

'==================================================================================
' FUNCIÓN: LimpiarNombreArchivo
'==================================================================================
Private Function LimpiarNombreArchivo(nombre As String) As String
    Dim caracteresInvalidos As Variant
    caracteresInvalidos = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Integer
    For i = 0 To UBound(caracteresInvalidos)
        nombre = Replace(nombre, caracteresInvalidos(i), "_")
    Next i
    LimpiarNombreArchivo = nombre
End Function