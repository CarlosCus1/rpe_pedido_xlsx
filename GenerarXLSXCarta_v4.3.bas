Option Explicit

'==================================================================================
' v4.3 - GENERAR XLSX CARTA (SIN PDF AUTOMÁTICO)
'==================================================================================
' Propósito: Generar archivo XLSX con formato de carta profesional para cotizaciones
'           sin exportación automática a PDF (responsabilidad del usuario)
' Entrada: Hoja "PEDIDOS" con datos, hoja "CONFIG" con empresa/vendedor/condiciones
' Salida: Archivo XLSX guardado en escritorio
' Versión: 4.3 (Enero 2026) - Unificación de versión con GenerarXLSXPedido
'           Textos mejorados en introducción y despedida
'==================================================================================

'==================================================================================
' USER-DEFINED TYPES (UDTs)
'==================================================================================
Private Type EmpresaInfo
    Nombre As String
    RUC As String
    Direccion As String
    Web As String
    SimboloMoneda As String
End Type

Private Type VendedorInfo
    Nombre As String
    Cargo As String
    Telefono As String
    Email As String
End Type

Private Type CondicionesComerciales
    ValidezCotizacion As String
    TipoPago As String
    PlazoEntrega As String
    Garantia As String
    MediosPago As String
    TextoIntroduccion As String
    TextoDespedida As String
End Type

Private Type PedidoInfo
    ClienteNombre As String
    NumeroPedido As String
    FechaActual As String
    Productos() As Variant
End Type

'==================================================================================
' CONSTANTES
'==================================================================================
Private Const IGV_RATE As Double = 0.18
Private Const FONT_NAME As String = "Calibri"
Private Const FONT_SIZE_NORMAL As Integer = 11
Private Const FONT_SIZE_SMALL As Integer = 10
Private Const FONT_SIZE_HEADER As Integer = 12

' Colores
Private Const COLOR_HEADER_BG As Long = 2893878
Private Const COLOR_HEADER_TEXT As Long = 16777215
Private Const COLOR_ROW_ALT As Long = 15921906
Private Const COLOR_TOTAL_BG As Long = 4605510
Private Const COLOR_TOTAL_TEXT As Long = 16777215

'==================================================================================
' PROCEDIMIENTO PRINCIPAL
'==================================================================================
Public Sub GenerarXLSXCarta_v4_3()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wsConfig As Worksheet, wsPedidos As Worksheet, wsCarta As Worksheet
    Dim empresa As EmpresaInfo, vendedor As VendedorInfo, condiciones As CondicionesComerciales
    Dim pedido As PedidoInfo
    Dim success As Boolean
    Dim wbNuevo As Workbook
    
    ' --- VALIDACIÓN DE HOJAS ---
    If Not WorksheetExists("CONFIG") Then
        MsgBox "La hoja 'CONFIG' no existe en este libro.", vbCritical, "Error"
        GoTo CleanExit
    End If
    
    If Not WorksheetExists("PEDIDOS") Then
        MsgBox "La hoja 'PEDIDOS' no existe en este libro.", vbCritical, "Error"
        GoTo CleanExit
    End If
    
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    Set wsPedidos = ThisWorkbook.Sheets("PEDIDOS")
    
    ' --- 1. LECTURA DE DATOS ---
    success = LeerDatosConfig(wsConfig, empresa, vendedor, condiciones)
    If Not success Then GoTo CleanExit
    
    With wsPedidos
        pedido.ClienteNombre = .Range("D2").Value
        pedido.NumeroPedido = .Range("D3").Value
    End With
    pedido.FechaActual = Format(Date, "dd ""de"" MMMM ""de"" yyyy")

    ' --- VALIDACIONES ---
    If pedido.ClienteNombre = "" Or pedido.NumeroPedido = "" Then
        MsgBox "Faltan datos del Cliente o N° de Pedido en la hoja PEDIDOS." & vbCrLf & _
               "Verifique las celdas D2 (Cliente) y D3 (N° Pedido).", vbCritical, "Datos Incompletos"
        GoTo CleanExit
    End If

    pedido.Productos = LeerProductos(wsPedidos)
    If Not IsArray(pedido.Productos) Or UBound(pedido.Productos) = 0 Then
        MsgBox "No se encontraron productos en la hoja PEDIDOS." & vbCrLf & _
               "Los productos deben comenzar en la fila 5, columna C.", vbExclamation, "Sin Productos"
        GoTo CleanExit
    End If
    
    ' --- 2. PROCESAMIENTO Y GENERACIÓN ---
    ' Crear nuevo libro para la carta
    Set wbNuevo = Workbooks.Add
    Set wsCarta = wbNuevo.Sheets(1)
    wsCarta.Name = "CARTA"
    
    LlenarCarta wsCarta, empresa, vendedor, condiciones, pedido, wsConfig, success
    If Not success Then GoTo CleanExit
    
    ConfigurarPagina wsCarta, empresa
    
    ' --- 3. GUARDAR COMO XLSX (SIN PDF AUTOMÁTICO) ---
    Dim outputFolder As String, outputFileName As String, outputPath As String
    outputFolder = Environ("USERPROFILE") & "\Desktop"
    
    ' Verificar que existe la carpeta Desktop
    If Dir(outputFolder, vbDirectory) = "" Then
        MsgBox "No se pudo acceder a la carpeta Escritorio." & vbCrLf & _
               "Ruta: " & outputFolder, vbCritical, "Error de Ruta"
        GoTo CleanExit
    End If
    
    outputFileName = "Cotizacion_" & LimpiarNombreArchivo(pedido.NumeroPedido & "_" & pedido.ClienteNombre) & ".xlsx"
    outputPath = outputFolder & "\" & outputFileName
    
    ' Guardar como XLSX
    Application.DisplayAlerts = False
    wbNuevo.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    ' Cerrar el libro (ya está guardado)
    wbNuevo.Close SaveChanges:=False
    
    MsgBox "Carta XLSX v4.3 generada exitosamente en:" & vbCrLf & outputPath & vbCrLf & vbCrLf & _
           "NOTA: Puede imprimir o guardar como PDF manualmente desde Excel." & vbCrLf & _
           "Los totales usan fórmulas que se recalculan automáticamente." & vbCrLf & _
           "La tabla de productos es una Tabla de Excel con autoajuste automático." & vbCrLf & _
           "Fuente Calibri tamaño 11 para mejor visibilidad." & vbCrLf & _
           "Columna A (ITEM) con ancho 8 para mejor visualización." & vbCrLf & _
           "Textos mejorados en introducción y despedida.", vbInformation, "Éxito - v4.3"
    
CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Ocurrió un error inesperado en la macro:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Ubicación: " & Erl, vbCritical, "Error en Macro"
    Resume CleanExit
End Sub

'==================================================================================
' FUNCIÓN: LeerDatosConfig
' Lee todos los datos de CONFIG usando celdas específicas
'==================================================================================
Private Function LeerDatosConfig(ws As Worksheet, ByRef empresa As EmpresaInfo, ByRef vendedor As VendedorInfo, ByRef cond As CondicionesComerciales) As Boolean
    On Error GoTo ErrorHandler
    
    ' Leer datos de la empresa (desde celdas específicas - compatible con setup_config_sheet_ES.vba v2.5)
    empresa.Nombre = ws.Range("B6").Value
    empresa.Direccion = ws.Range("B7").Value
    empresa.Web = ws.Range("B10").Value
    empresa.RUC = ExtraerRUC(ws.Range("B25").Value)
    empresa.SimboloMoneda = IIf(ws.Range("B26").Value <> "", ws.Range("B26").Value, "S/.")
    
    ' Leer datos del vendedor
    vendedor.Nombre = ws.Range("B15").Value
    vendedor.Cargo = ""  ' No existe campo de cargo en CONFIG v2.5
    vendedor.Telefono = ws.Range("B16").Value
    vendedor.Email = ws.Range("B17").Value
    
    ' Leer condiciones
    cond.ValidezCotizacion = ws.Range("B20").Value
    cond.TipoPago = ws.Range("B21").Value
    cond.PlazoEntrega = ws.Range("B22").Value
    cond.Garantia = ws.Range("B23").Value
    cond.MediosPago = ws.Range("B28").Value
    
    ' Leer mensajes personalizables (v2.6) - Si están vacíos, usar valores predeterminados
    cond.TextoIntroduccion = ws.Range("B31").Value
    If cond.TextoIntroduccion = "" Then
        cond.TextoIntroduccion = "Estimados: | Es un gusto saludarles. Les envío la propuesta comercial sobre los productos consultados. | Quedamos a su disposición para cualquier consulta:"
    End If
    
    cond.TextoDespedida = ws.Range("B32").Value
    If cond.TextoDespedida = "" Then
        cond.TextoDespedida = "Agradecemos su interés y quedamos atentos a su aprobación. | Confiamos en que la calidad de nuestra marca sea de su agrado."
    End If
    
    On Error GoTo 0
    LeerDatosConfig = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error al leer la hoja 'CONFIG'." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Asegúrese de que la hoja CONFIG esté creada con el formato correcto (v2.5)." & vbCrLf & _
           "Ejecute la macro 'CrearHojaDeConfiguracion' si es necesario.", vbCritical, "Error de Configuración"
    LeerDatosConfig = False
End Function

'==================================================================================
' FUNCIÓN: WorksheetExists
' Verifica si existe una hoja en el libro activo
'==================================================================================
Private Function WorksheetExists(sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function

'==================================================================================
' FUNCIÓN: ExtraerRUC
'==================================================================================
Private Function ExtraerRUC(texto As String) As String
    Dim regex As Object, matches As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d{11}"
    If regex.Test(texto) Then
        Set matches = regex.Execute(texto)
        ExtraerRUC = matches(0).Value
    Else
        ExtraerRUC = ""
    End If
End Function

'==================================================================================
' FUNCIÓN: LeerProductos
'==================================================================================
Private Function LeerProductos(ws As Worksheet) As Variant
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    If lastRow < 5 Then
        LeerProductos = Array()
        Exit Function
    End If
    
    ' Leer el rango de productos
    LeerProductos = ws.Range("C5:J" & lastRow).Value
    Exit Function
    
ErrorHandler:
    MsgBox "Error al leer productos de la hoja PEDIDOS:" & vbCrLf & _
           Err.Description, vbCritical, "Error de Lectura"
    LeerProductos = Array()
End Function

'==================================================================================
' PROCEDIMIENTO: LlenarCarta
'==================================================================================
Private Sub LlenarCarta(ws As Worksheet, empresa As EmpresaInfo, vendedor As VendedorInfo, condiciones As CondicionesComerciales, pedido As PedidoInfo, wsConfig As Worksheet, ByRef success As Boolean)
    On Error GoTo ErrorHandler
    
    Dim currentRow As Long: currentRow = 1
    Dim i As Long, simboloMoneda As String: simboloMoneda = empresa.SimboloMoneda
    Dim tblStartRow As Long, tblEndRow As Long, subtotalRow As Long, igvRow As Long, totalRow As Long
    Dim tbl As ListObject
    
    success = False
    
    ' Configurar fuente y tamaño para toda la hoja
    ws.Columns("A:G").Font.Name = FONT_NAME
    ws.Columns("A:G").Font.Size = FONT_SIZE_NORMAL
    
    ' Configurar ancho fijo para columna A (ITEM) - ajuste manual al imprimir
    ws.Columns("A").ColumnWidth = 8
    
    ' --- ENCABEZADO ---
    CopiarLogo wsConfig, ws, ws.Range("A1"), success
    If Not success Then Exit Sub
    
    With ws.Range("C1:G2")
        .Merge
        .Value = empresa.Nombre
        .Font.Size = 15
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A3:G3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    currentRow = 4
    
    ' Fila vacía entre encabezado y fecha
    currentRow = currentRow + 1

    ' --- DATOS DE COTIZACIÓN Y CLIENTE ---
    ws.Cells(currentRow, 1).Value = "COTIZACIÓN N°: " & pedido.NumeroPedido
    ws.Cells(currentRow, 6).Value = "Fecha:"
    ws.Cells(currentRow, 7).Value = pedido.FechaActual
    ws.Range(ws.Cells(currentRow, 6), ws.Cells(currentRow, 7)).HorizontalAlignment = xlRight
    currentRow = currentRow + 2
    ws.Cells(currentRow, 1).Value = "SEÑOR(ES):"
    ws.Cells(currentRow, 1).Font.Bold = True
    ws.Range("B" & currentRow).Value = pedido.ClienteNombre
    currentRow = currentRow + 2
    
    ' --- INTRODUCCIÓN MEJORADA (v4.3) ---
    ' Cargar desde CONFIG (B31) o usar valor predeterminado
    currentRow = FormatearTextoConPipe(ws, currentRow, condiciones.TextoIntroduccion, "A", "G")
    currentRow = currentRow + 1
    
    ' --- TABLA DE PRODUCTOS (7 COLUMNAS) ---
    tblStartRow = currentRow
    Dim headers: headers = Array("ITEM", "CÓDIGO", "DESCRIPCIÓN", "CANT.", "U/M", "P. UNIT.", "TOTAL")
    
    ' Llenar encabezados
    With ws.Range("A" & currentRow & ":G" & currentRow)
        .Value = headers
        .Font.Bold = True
        .Font.Size = FONT_SIZE_SMALL
        .Interior.Color = COLOR_HEADER_BG
        .Font.Color = COLOR_HEADER_TEXT
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    currentRow = currentRow + 1
    
    ' Llenar datos de productos con FÓRMULAS
    For i = 1 To UBound(pedido.Productos, 1)
        Dim cantidad As Double: cantidad = Val(pedido.Productos(i, 3))
        Dim valorUnitario As Double: valorUnitario = Val(pedido.Productos(i, 6))
        Dim desc1 As Double: desc1 = Val(pedido.Productos(i, 7))
        Dim desc2 As Double: desc2 = Val(pedido.Productos(i, 8))
        Dim precioUnitConDesc As Double: precioUnitConDesc = valorUnitario * (1 - desc1 / 100) * (1 - desc2 / 100)
        
        ws.Cells(currentRow, 1).Value = i
        ws.Cells(currentRow, 2).Value = "'" & pedido.Productos(i, 1)
        ws.Cells(currentRow, 3).Value = pedido.Productos(i, 2)
        ws.Cells(currentRow, 4).Value = cantidad
        ws.Cells(currentRow, 5).Value = pedido.Productos(i, 5)
        ws.Cells(currentRow, 6).Value = precioUnitConDesc
        
        ' FÓRMULA para el total: =Cantidad × Precio Unitario
        ws.Cells(currentRow, 7).Formula = "=D" & currentRow & "*F" & currentRow
        
        ' Aplicar formato numérico con NumberFormatLocal (respeta configuración regional)
        On Error Resume Next
        ws.Cells(currentRow, 4).NumberFormatLocal = "#,##0"           ' CANT. - sin decimales
        ws.Cells(currentRow, 6).NumberFormatLocal = "#,##0.00"        ' P. UNIT. - 2 decimales
        ws.Cells(currentRow, 7).NumberFormatLocal = "#,##0.00"        ' TOTAL - 2 decimales
        On Error GoTo ErrorHandler
        
        If i Mod 2 = 0 Then ws.Range("A" & currentRow & ":G" & currentRow).Interior.Color = COLOR_ROW_ALT
        currentRow = currentRow + 1
    Next i
    
    tblEndRow = currentRow - 1
    
    ' --- CONVERTIR RANGO EN TABLA DE EXCEL (LISTOBJECT) ---
    On Error Resume Next
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A" & tblStartRow & ":G" & tblEndRow), , xlYes)
    On Error GoTo ErrorHandler
    
    If Not tbl Is Nothing Then
        With tbl
            .Name = "TablaProductos"
            .TableStyle = "TableStyleMedium9"
            .ShowHeaders = True
            .ShowAutoFilter = False  ' Ocultar filtros para impresión
            .ShowTableStyleRowStripes = False  ' Usar nuestro propio formato de filas alternas
            .ShowTableStyleFirstColumn = False
            .ShowTableStyleLastColumn = False
        End With
        
        ' Autoajuste de columnas B:G (todas excepto A)
        On Error Resume Next
        ws.Range("B" & tblStartRow & ":G" & tblEndRow).Columns.AutoFit
        On Error GoTo ErrorHandler
    End If
    
    ' Formatear bordes de la tabla
    With ws.Range("A" & tblStartRow & ":G" & tblEndRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(200, 200, 200)
    End With
    
    currentRow = currentRow + 1
    
    ' --- TOTALES CON FÓRMULAS E IGV VISIBLE ---
    subtotalRow = currentRow
    igvRow = currentRow + 1
    totalRow = currentRow + 2
    
    With ws.Range("F" & currentRow & ":G" & currentRow + 2)
        .Font.Bold = True
        .Font.Size = FONT_SIZE_HEADER
        .Interior.Color = COLOR_TOTAL_BG
        .Font.Color = COLOR_TOTAL_TEXT
        .HorizontalAlignment = xlRight
    End With
    
    ' FÓRMULAS de totales
    ws.Range("E" & subtotalRow).Value = "SUBTOTAL"
    ws.Range("G" & subtotalRow).Formula = "=SUM(TablaProductos[TOTAL])"
    
    ws.Range("E" & igvRow).Value = "IGV"
    ws.Range("F" & igvRow).Value = IGV_RATE
    ws.Range("F" & igvRow).NumberFormatLocal = "0%"
    ws.Range("G" & igvRow).Formula = "=G" & subtotalRow & "*F" & igvRow
    
    ws.Range("E" & totalRow).Value = "TOTAL"
    ws.Range("G" & totalRow).Formula = "=G" & subtotalRow & "+G" & igvRow
    
    ' Aplicar formato numérico a los totales
    On Error Resume Next
    ws.Range("G" & subtotalRow & ":G" & totalRow).NumberFormatLocal = "#,##0.00"
    On Error GoTo ErrorHandler
    
    currentRow = currentRow + 4
    
    ' --- CONDICIONES COMERCIALES Y MEDIOS DE PAGO ---
    ws.Cells(currentRow, 1).Value = "CONDICIONES COMERCIALES"
    ws.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    
    Dim arrCond(1 To 4) As String
    arrCond(1) = "Validez de la oferta: " & condiciones.ValidezCotizacion
    arrCond(2) = "Forma de pago: " & condiciones.TipoPago
    arrCond(3) = "Plazo de entrega: " & condiciones.PlazoEntrega
    arrCond(4) = "Garantía: " & condiciones.Garantia
    
    For i = 1 To 4
        ws.Cells(currentRow, 1).Value = "•"
        ws.Cells(currentRow, 2).Value = arrCond(i)
        ws.Range("B" & currentRow & ":D" & currentRow).Merge
        currentRow = currentRow + 1
    Next i
    currentRow = currentRow + 1

    If condiciones.MediosPago <> "" Then
        ws.Cells(currentRow, 1).Value = "MEDIOS DE PAGO"
        ws.Cells(currentRow, 1).Font.Bold = True
        currentRow = FormatearMediosDePago(ws, currentRow + 1, condiciones.MediosPago)
        currentRow = currentRow + 1
    End If
    
    ' --- DESPEDIDA MEJORADA (v4.3) ---
    ' Cargar desde CONFIG (B32) o usar valor predeterminado
    currentRow = FormatearTextoConPipe(ws, currentRow, condiciones.TextoDespedida, "A", "G")
    currentRow = currentRow + 2
    
    ws.Cells(currentRow, 1).Value = "Atentamente,"
    currentRow = currentRow + 4
    
    ws.Cells(currentRow, 1).Value = vendedor.Nombre
    ws.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    ws.Cells(currentRow, 1).Value = vendedor.Cargo
    currentRow = currentRow + 1
    ws.Cells(currentRow, 1).Value = "T: " & vendedor.Telefono & " | E: " & vendedor.Email
    ws.Range("A" & currentRow - 2 & ":D" & currentRow).Borders(xlEdgeTop).LineStyle = xlContinuous
    
    success = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al llenar la carta:" & vbCrLf & Err.Description, vbCritical, "Error"
    success = False
End Sub

'==================================================================================
' PROCEDIMIENTO: FormatearMediosDePago
'==================================================================================
Private Function FormatearMediosDePago(ws As Worksheet, startRow As Long, texto As String) As Long
    FormatearMediosDePago = FormatearTextoConPipe(ws, startRow, texto, "B", "G")
End Function

'==================================================================================
' PROCEDIMIENTO: FormatearTextoConPipe
' Formatea texto usando "|" como separador, similar a medios de pago
' Cada segmento va en una fila separada, combinando columnas especificadas
'==================================================================================
Private Function FormatearTextoConPipe(ws As Worksheet, startRow As Long, texto As String, colInicio As String, colFin As String) As Long
    Dim lineas: lineas = Split(texto, "|")
    Dim i As Long
    Dim rng As String
    rng = colInicio & startRow & ":" & colFin & startRow
    
    For i = 0 To UBound(lineas)
        With ws.Range(colInicio & (startRow + i) & ":" & colFin & (startRow + i))
            .Merge
            .Value = Trim(lineas(i))
            .VerticalAlignment = xlTop
            .WrapText = False  ' Sin ajuste de texto para mejor control
        End With
        ' Auto-ajustar altura de fila
        ws.Rows(startRow + i).AutoFit
    Next i
    FormatearTextoConPipe = startRow + UBound(lineas) + 1
End Function

'==================================================================================
' PROCEDIMIENTO: CopiarLogo
'==================================================================================
Private Sub CopiarLogo(wsSource As Worksheet, wsTarget As Worksheet, targetCell As Range, ByRef success As Boolean)
    On Error GoTo ErrorHandler
    
    Dim shp As Shape
    
    On Error Resume Next
    Set shp = wsSource.Shapes("logo_empresa")
    On Error GoTo ErrorHandler
    
    If shp Is Nothing Then
        MsgBox "Advertencia: No se encontró el objeto con el nombre 'logo_empresa' en la hoja 'CONFIG'." & vbCrLf & _
               "La carta se generará sin logotipo.", vbExclamation, "Advertencia"
        success = True  ' Continuar sin logo
        Exit Sub
    End If
    
    ' Copiar logo
    shp.Copy
    wsTarget.Paste
    
    ' Ajustar tamaño y posición
    With wsTarget.Shapes(wsTarget.Shapes.Count)
        .LockAspectRatio = msoTrue
        .Height = 45
        .Top = targetCell.Top + (targetCell.Height - .Height) / 2
        .Left = targetCell.Left + 5
    End With
    
    success = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al copiar el logotipo:" & vbCrLf & Err.Description, vbExclamation, "Error"
    success = True  ' Continuar sin logo
End Sub

'==================================================================================
' PROCEDIMIENTO: ConfigurarPagina
'==================================================================================
Private Sub ConfigurarPagina(ws As Worksheet, empresa As EmpresaInfo)
    On Error GoTo ErrorHandler
    
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 2
        .CenterHorizontally = True
        
        Dim footerText As String
        footerText = empresa.Direccion & " | RUC: " & empresa.RUC & " | " & empresa.Web
        .CenterFooter = "&""Calibri,Regular""&8" & footerText
        .RightFooter = "&""Calibri,Bold""&8Página &P de &N"
        
        .PrintGridlines = False
        .PrintHeadings = False
    End With
    
    On Error Resume Next
    ws.PageSetup.PrintArea = ws.UsedRange.Address
    On Error GoTo ErrorHandler
    ws.DisplayPageBreaks = False
    Exit Sub
    
ErrorHandler:
    ' Continuar aunque falle la configuración de página
    ' No es crítico para la generación del archivo
End Sub

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