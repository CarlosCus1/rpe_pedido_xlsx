Option Explicit

'==================================================================================
' UTILIDADES PARA GESTIÓN DE PEDIDOS v1.0
'==================================================================================
' Propósito: Proporcionar herramientas de utilidad para la gestión de pedidos
'            incluyendo limpieza de datos, restauración de formato, etc.
' Versión: 1.0 (Enero 2026)
'==================================================================================

'==================================================================================
' MACRO: LimpiarHojaPedidos
' Limpia todos los datos de la hoja PEDIDOS y la deja en estado inicial
'==================================================================================
Public Sub LimpiarHojaPedidos()
    Dim wsPedidos As Worksheet
    Dim respuesta As VbMsgBoxResult
    
    ' Verificar que existe la hoja PEDIDOS
    On Error Resume Next
    Set wsPedidos = ThisWorkbook.Sheets("PEDIDOS")
    On Error GoTo ErrorHandler
    
    If wsPedidos Is Nothing Then
        MsgBox "La hoja 'PEDIDOS' no existe en este libro." & vbCrLf & _
               "No se puede realizar la limpieza.", vbExclamation, "Hoja No Encontrada"
        Exit Sub
    End If
    
    ' Confirmar antes de limpiar
    respuesta = MsgBox("¿Está seguro de que desea limpiar la hoja PEDIDOS?" & vbCrLf & vbCrLf & _
                      "Se eliminarán:" & vbCrLf & _
                      "• Datos del cliente (D2, D3)" & vbCrLf & _
                      "• Todos los productos (desde A4 hasta AB + última fila con datos)" & vbCrLf & _
                      "• Formato aplicado a los datos" & vbCrLf & vbCrLf & _
                      "Esta acción NO se puede deshacer.", _
                      vbYesNo + vbExclamation + vbDefaultButton2, _
                      "Confirmar Limpieza")
    
    If respuesta = vbNo Then
        MsgBox "Limpieza cancelada.", vbInformation, "Cancelado"
        Exit Sub
    End If
    
    ' Desactivar actualización de pantalla para velocidad
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Limpiar datos del cliente y restaurar placeholders
    With wsPedidos
        .Range("D2").Value = "CLIENTE:"
        .Range("D2").Font.Italic = True
        .Range("D2").Font.Color = RGB(128, 128, 128)  ' Gris
        
        .Range("D3").Value = "PEDIDO:"
        .Range("D3").Font.Italic = True
        .Range("D3").Font.Color = RGB(128, 128, 128)  ' Gris
        
        ' Limpiar datos de productos (desde fila 5, columnas C hasta N)
        ' Primero determinar la última fila con datos
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        
        ' Limpiar desde A4 hasta AB (columna 28) - Incluye encabezados C4:J4
        If lastRow >= 4 Then
            ' Limpiar TODO: desde A4 hasta AB y hasta la última fila con datos
            .Range("A4:AB" & lastRow).ClearContents
            .Range("A4:AB" & lastRow).ClearFormats
        End If
        
        ' Restaurar encabezados en C4:J4 (K4:N4 NO son encabezados)
        .Range("C4").Value = "CÓDIGO"
        .Range("D4").Value = "DESCRIPCIÓN"
        .Range("E4").Value = "CANT."
        .Range("F4").Value = "STOCK"
        .Range("G4").Value = "U/M"
        .Range("H4").Value = "PRECIO"
        .Range("I4").Value = "DESC1"
        .Range("J4").Value = "DESC2"
        With .Range("C4:J4")
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217)
        End With
        
          ' Placeholder en A5 con alto contraste
          .Range("B5:AB5").ClearContents
          .Range("B5:AB5").ClearFormats
          
          .Range("A5").Value = "[PEGAR AQUÍ - Datos desde sistema RPE]"
          .Range("A5").Font.Italic = True
          .Range("A5").Font.Bold = True
          .Range("A5").Font.Size = 11
          .Range("A5").Font.Color = RGB(0, 51, 102)  ' Azul oscuro
          .Range("A5").Interior.Color = RGB(255, 255, 153)  ' Amarillo claro
         
         ' Configurar columna C como TEXTO para preservar ceros a la izquierda (ej: "02182")
         .Columns("C").NumberFormat = "@"
         .Range("C5").NumberFormat = "@"
        
    End With
    
    ' Reactivar actualización de pantalla
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "✓ Limpieza completada exitosamente." & vbCrLf & vbCrLf & _
           "La hoja PEDIDOS ha sido restaurada a su estado inicial." & vbCrLf & _
           "Puede ahora ingresar nuevos datos.", vbInformation, "Limpieza Completada"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Ocurrió un error durante la limpieza:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

'==================================================================================
' MACRO: RestaurarFormatoPedidos
' Restaura solo el formato de la hoja PEDIDOS sin borrar datos
'==================================================================================
Public Sub RestaurarFormatoPedidos()
    Dim wsPedidos As Worksheet
    Dim respuesta As VbMsgBoxResult
    
    ' Verificar que existe la hoja PEDIDOS
    On Error Resume Next
    Set wsPedidos = ThisWorkbook.Sheets("PEDIDOS")
    On Error GoTo ErrorHandler
    
    If wsPedidos Is Nothing Then
        MsgBox "La hoja 'PEDIDOS' no existe en este libro.", vbExclamation, "Hoja No Encontrada"
        Exit Sub
    End If
    
    ' Confirmar
    respuesta = MsgBox("¿Desea restaurar el formato de la hoja PEDIDOS?" & vbCrLf & vbCrLf & _
                      "Los datos NO se borrarán, solo se ajustará el formato.", _
                      vbYesNo + vbQuestion, "Restaurar Formato")
    
    If respuesta = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    With wsPedidos
        ' Restaurar formato de encabezados C4:J4
        .Range("C4").Value = "CÓDIGO"
        .Range("D4").Value = "DESCRIPCIÓN"
        .Range("E4").Value = "CANT."
        .Range("F4").Value = "STOCK"
        .Range("G4").Value = "U/M"
        .Range("H4").Value = "PRECIO"
        .Range("I4").Value = "DESC1"
        .Range("J4").Value = "DESC2"
        With .Range("C4:J4")
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = RGB(217, 217, 217)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
            .HorizontalAlignment = xlCenter
        End With
        
         ' Auto-ajustar columnas
         .Columns("C:AB").AutoFit
         
         ' Configurar columna C como TEXTO para preservar ceros a la izquierda (ej: "02182")
         .Columns("C").NumberFormat = "@"
         
         ' Ajustar ancho mínimo para columnas de texto
        If .Columns("D").ColumnWidth < 15 Then .Columns("D").ColumnWidth = 15
        If .Columns("E").ColumnWidth < 30 Then .Columns("E").ColumnWidth = 30
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "✓ Formato restaurado.", vbInformation, "Completado"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

'==================================================================================
' MACRO: PrepararNuevoPedido
' Prepara la hoja para un nuevo pedido limpiando datos pero manteniendo formato
'==================================================================================
Public Sub PrepararNuevoPedido()
    Dim wsPedidos As Worksheet
    Dim respuesta As VbMsgBoxResult
    
    On Error Resume Next
    Set wsPedidos = ThisWorkbook.Sheets("PEDIDOS")
    On Error GoTo ErrorHandler
    
    If wsPedidos Is Nothing Then
        MsgBox "La hoja 'PEDIDOS' no existe.", vbExclamation, "Error"
        Exit Sub
    End If
    
    respuesta = MsgBox("¿Preparar hoja para un NUEVO PEDIDO?" & vbCrLf & vbCrLf & _
                      "Se borrarán:" & vbCrLf & _
                      "• Cliente y N° de Pedido" & vbCrLf & _
                      "• Todos los productos" & vbCrLf & _
                      "• Se mantendrá el formato" & vbCrLf & vbCrLf & _
                      "¿Continuar?", vbYesNo + vbQuestion, "Nuevo Pedido")
    
    If respuesta = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    With wsPedidos
        ' Restaurar placeholders en lugar de solo limpiar
        .Range("D2").Value = "CLIENTE:"
        .Range("D2").Font.Italic = True
        .Range("D2").Font.Color = RGB(128, 128, 128)
        
        .Range("D3").Value = "PEDIDO:"
        .Range("D3").Font.Italic = True
        .Range("D3").Font.Color = RGB(128, 128, 128)
        
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        
        ' Limpiar desde A4 hasta AB (columna 28) - Incluye encabezados C4:J4
        If lastRow >= 4 Then
            .Range("A4:AB" & lastRow).ClearContents
            .Range("A4:AB" & lastRow).ClearFormats
        End If
        
        ' Restaurar encabezados en C4:J4 (K4:N4 NO son encabezados)
        .Range("C4").Value = "CÓDIGO"
        .Range("D4").Value = "DESCRIPCIÓN"
        .Range("E4").Value = "CANT."
        .Range("F4").Value = "STOCK"
        .Range("G4").Value = "U/M"
        .Range("H4").Value = "PRECIO"
        .Range("I4").Value = "DESC1"
        .Range("J4").Value = "DESC2"
        With .Range("C4:J4")
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217)
        End With
        
          ' Placeholder en A5 con alto contraste
          .Range("B5:AB5").ClearContents
          .Range("B5:AB5").ClearFormats
          
          .Range("A5").Value = "[PEGAR AQUÍ - Datos desde sistema RPE]"
          .Range("A5").Font.Italic = True
          .Range("A5").Font.Bold = True
          .Range("A5").Font.Size = 11
          .Range("A5").Font.Color = RGB(0, 51, 102)  ' Azul oscuro
          .Range("A5").Interior.Color = RGB(255, 255, 153)  ' Amarillo claro
         
         ' Configurar columna C como TEXTO para preservar ceros a la izquierda (ej: "02182")
         .Columns("C").NumberFormat = "@"
         .Range("C5").NumberFormat = "@"
        
        ' Posicionar cursor en D2 para empezar
        .Range("D2").Select
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "✓ Listo para nuevo pedido." & vbCrLf & vbCrLf & _
           "Ingrese el cliente en D2 y el pedido en D3.", vbInformation, "Nuevo Pedido"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

'==================================================================================
' MACRO: VerificarDatosPedido
' Verifica que los datos del pedido estén completos antes de procesar
'==================================================================================
Public Function VerificarDatosPedido() As Boolean
    Dim wsPedidos As Worksheet
    Dim hayErrores As Boolean
    Dim mensajeError As String
    
    On Error Resume Next
    Set wsPedidos = ThisWorkbook.Sheets("PEDIDOS")
    On Error GoTo 0
    
    If wsPedidos Is Nothing Then
        MsgBox "La hoja 'PEDIDOS' no existe.", vbCritical, "Error"
        VerificarDatosPedido = False
        Exit Function
    End If
    
    hayErrores = False
    mensajeError = "Se encontraron los siguientes problemas:" & vbCrLf & vbCrLf
    
    With wsPedidos
        ' Verificar cliente
        If Trim(.Range("D2").Value) = "" Then
            mensajeError = mensajeError & "• Falta el nombre del cliente (D2)" & vbCrLf
            hayErrores = True
        End If
        
        ' Verificar número de pedido
        If Trim(.Range("D3").Value) = "" Then
            mensajeError = mensajeError & "• Falta el número de pedido (D3)" & vbCrLf
            hayErrores = True
        End If
        
        ' Verificar que hay productos
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        If lastRow < 5 Then
            mensajeError = mensajeError & "• No hay productos ingresados (desde fila 5)" & vbCrLf
            hayErrores = True
        End If
    End With
    
    If hayErrores Then
        MsgBox mensajeError & vbCrLf & "Por favor, complete los datos antes de generar la carta.", _
               vbExclamation, "Datos Incompletos"
        VerificarDatosPedido = False
    Else
        VerificarDatosPedido = True
    End If
End Function