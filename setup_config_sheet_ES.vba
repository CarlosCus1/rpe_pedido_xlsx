Option Explicit

'==================================================================================================
' MACRO DE CONFIGURACIÓN INICIAL v2.6 - EJECUTAR SÓLO UNA VEZ
' Propósito: Crea y formatea la hoja "CONFIG" con estructura v2.6
'            Alineado con CONFIG_ESTRUCTURA.txt y macros v4.3
'            Compatible con macros de XLSX (sin PDF automático)
'            NOVEDAD v2.6: Mensajes de carta personalizables (B31-B32)
'==================================================================================================

Public Sub CrearHojaDeConfiguracion()
    Dim ws As Worksheet
    Dim sheetName As String: sheetName = "CONFIG"

    ' Verificar si la hoja ya existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        If MsgBox("La hoja 'CONFIG' ya existe. ¿Desea borrarla y crear una nueva con el formato v2.5?", _
                  vbYesNo + vbExclamation, "Hoja Existente") = vbNo Then
            Exit Sub
        End If
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    ' Crear la nueva hoja
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = sheetName

    ' --- Aplicar formato y agregar contenido ---
    With ws
        ' Ajustar anchos de columna
        .Columns("A").ColumnWidth = 40
        .Columns("B").ColumnWidth = 60
        .Columns("A:B").Font.Name = "Calibri"
        .Columns("A:B").Font.Size = 10

        ' --- LOGO AREA (A1:B3) - Para logo de empresa ---
        .Range("A1:B3").Merge
        .Range("A1").Value = "[LOGO - Nombre: logo_empresa]"
        Call FormatoAreaLogo(.Range("A1"))
        .Range("A1:B3").RowHeight = 60

        ' --- SECCIÓN: DATOS DE LA EMPRESA (Constantes: B6:B10) ---
        .Range("A5").Value = "DATOS DE LA EMPRESA"
        Call FormatoTitulo(.Range("A5"), 12, "333333")
        
        PoblarCampo ws, "A6", "B6", "Nombre de la Empresa:", "CIP COMERCIAL S.A.C."
        PoblarCampo ws, "A7", "B7", "Dirección:", "Av. Principal 123, Lima, Perú"
        PoblarCampo ws, "A8", "B8", "Teléfono Empresa:", "+51 1 2345678"
        PoblarCampo ws, "A9", "B9", "Email Empresa:", "contacto@cipcomercial.com"
        PoblarCampo ws, "A10", "B10", "Website:", "www.cipcomercial.com"
        
        ' --- SECCIÓN: DATOS DEL VENDEDOR (Constantes: B15:B17) ---
        .Range("A13").Value = "DATOS DEL VENDEDOR"
        Call FormatoTitulo(.Range("A13"), 12, "333333")
        
        PoblarCampo ws, "A15", "B15", "Nombre del Vendedor:", "Juan Pérez García"
        PoblarCampo ws, "A16", "B16", "Teléfono del Vendedor:", "999 123 456"
        PoblarCampo ws, "A17", "B17", "Email del Vendedor:", "juan.perez@cipcomercial.com"
        
        ' --- SECCIÓN: CONDICIONES COMERCIALES ESTÁNDAR (Constantes: B20:B25) ---
        .Range("A19").Value = "CONDICIONES COMERCIALES ESTÁNDAR"
        Call FormatoTitulo(.Range("A19"), 12, "333333")
        
        PoblarCampo ws, "A20", "B20", "Validez de Cotización:", "30 días a partir de la fecha de este documento"
        PoblarCampo ws, "A21", "B21", "Tipo de Pago:", "Contado / Crédito 30 días"
        PoblarCampo ws, "A22", "B22", "Plazo de Entrega:", "3-5 días hábiles a partir de confirmación"
        PoblarCampo ws, "A23", "B23", "Condición Especial 1:", "Garantía 12 meses en todos los productos"
        PoblarCampo ws, "A24", "B24", "Condición Especial 2:", "Transporte incluido a domicilio dentro de Lima"
        
        ' --- PIE DE PÁGINA PARA PDF (Constante: B25) ---
        PoblarCampo ws, "A25", "B25", "Pie de Página PDF:", "CIP COMERCIAL S.A.C. - RUC: 20123456789 - Lima, Perú"

        ' --- SECCIÓN: CONFIGURACIÓN DE PAGO Y MONEDA ---
        .Range("A27").Value = "CONFIGURACIÓN DE PAGO Y MONEDA"
        Call FormatoTitulo(.Range("A27"), 12, "333333")
        
        PoblarCampo ws, "A26", "B26", "Moneda:", "S/. "
        PoblarCampo ws, "A28", "B28", "Medios de Pago:", "BCP Soles 191-12345678-0-00 | CCI: 002-191-001234567890-00 | Yape: 999 123 456"
        
        ' --- SECCIÓN: MENSAJES DE CARTA PERSONALIZABLES (v2.6) ---
        .Range("A30").Value = "MENSAJES DE CARTA PERSONALIZABLES"
        Call FormatoTitulo(.Range("A30"), 12, "2E7D32")  ' Verde para destacar
        
        PoblarCampo ws, "A31", "B31", "Texto de Introducción:", "Estimados: | Es un gusto saludarles. Les envío la propuesta comercial sobre los productos consultados. | Quedamos a su disposición para cualquier consulta:"
        PoblarCampo ws, "A32", "B32", "Texto de Despedida:", "Agradecemos su interés y quedamos atentos a su aprobación. | Confiamos en que la calidad de nuestra marca sea de su agrado."

        ' --- INSTRUCCIONES FINALES ---
        .Range("A34").Value = "INSTRUCCIONES IMPORTANTES:"
        Call FormatoTitulo(.Range("A34"), 11, "404040")
        
        .Range("A35").Value = "1. Logo:"
        .Range("B35").Value = "Inserte su logo en el área A1:B3, luego seleccione la imagen y nómbrela 'logo_empresa' en el cuadro de nombres."
        .Range("B35").WrapText = True: .Range("B35").Font.Italic = True
        
        .Range("A36").Value = "2. Datos:"
        .Range("B36").Value = "Complete TODAS las celdas de la columna B. Son obligatorias para generar cotizaciones."
        .Range("B36").WrapText = True: .Range("B36").Font.Italic = True
        
        .Range("A37").Value = "3. Moneda (B26):"
        .Range("B37").Value = "Símbolo de moneda que aparecerá en los precios. Ejemplos: 'S/. ' (soles), '$ ' (dólares). Dejar vacío para usar S/. por defecto."
        .Range("B37").WrapText = True: .Range("B37").Font.Italic = True

        .Range("A38").Value = "4. Medios de Pago (B28):"
        .Range("B38").Value = "Información de medios de pago (cuentas bancarias, Yape, etc.). Use '|' para separar líneas."
        .Range("B38").WrapText = True: .Range("B38").Font.Italic = True
        
        .Range("A39").Value = "5. Mensajes de Carta (B31-B32):"
        .Range("B39").Value = "Textos de introducción y despedida de la carta. Use '|' para separar párrafos. Si deja vacío, se usarán textos predeterminados."
        .Range("B39").WrapText = True: .Range("B39").Font.Italic = True
        
        .Rows(35).RowHeight = 30
        .Rows(36).RowHeight = 25
        .Rows(37).RowHeight = 25
        .Rows(38).RowHeight = 25
        .Rows(39).RowHeight = 30

    End With
    
    ws.Activate
    MsgBox "✓ Hoja CONFIG v2.6 creada exitosamente." & vbNewLine & vbNewLine & _
           "PRÓXIMOS PASOS:" & vbNewLine & _
           "1. Complete todos los datos en la columna B" & vbNewLine & _
           "2. Inserte el logo en A1:B3 y nómbrelo 'logo_empresa'" & vbNewLine & _
           "3. Personalice los mensajes de carta en B31 (intro) y B32 (despedida)" & vbNewLine & _
           "4. Guarde el archivo" & vbNewLine & vbNewLine & _
           "¡Listo para generar cotizaciones XLSX!", _
            vbInformation, "Configuración v2.6 Completada"

End Sub

' --- Subrutinas de ayuda para formatear ---

Private Sub PoblarCampo(ws As Worksheet, labelCell As String, valueCell As String, label As String, example As String)
    With ws.Range(labelCell)
        .Value = label
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlRight
        .Interior.Color = RGB(242, 242, 242)
        .VerticalAlignment = xlCenter
        
        With .Offset(0, 1)
            .Value = example
            .Font.Size = 10
            .Interior.Color = RGB(255, 255, 255)
            .VerticalAlignment = xlCenter
        End With
        
        .Resize(1, 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Resize(1, 2).Borders(xlEdgeBottom).Weight = xlThin
        .Resize(1, 2).Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
    End With
End Sub

Private Sub FormatoTitulo(rng As Range, fontSize As Integer, colorHex As String)
    Dim rgbColor As Long
    rgbColor = CLng("&H" & colorHex)
    
    With rng
        .Font.Bold = True
        .Font.Size = fontSize
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = rgbColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With
End Sub

Private Sub FormatoAreaLogo(rng As Range)
    With rng
        .Font.Italic = True
        .Font.Size = 10
        .Font.Color = RGB(150, 150, 150)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(180, 180, 180)
    End With
End Sub