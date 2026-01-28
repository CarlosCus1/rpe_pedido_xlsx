# RPE Pedido XLSX - Macros VBA para Gestión de Pedidos

Sistema de macros VBA para Excel diseñado para automatizar la generación de documentos XLSX a partir de datos del sistema RPE.

## 📋 Características

- **Generación de Carta de Cotización**: Crea documentos XLSX profesionales con tabla de productos, condiciones comerciales y mensajes personalizables.
- **Generación de Pedido Técnico**: Produce archivos XLSX con formato técnico, tabla de 12 columnas, cálculos automáticos y estado de stock.
- **Sin PDF Automático**: Elimina la complejidad de generación automática de PDF. El usuario puede imprimir o guardar como PDF manualmente desde Excel.
- **Preservación de Códigos**: La columna de códigos (CÓDIGO) mantiene formato texto para preservar ceros a la izquierda (ej: "02182").
- **Configuración Flexible**: Hoja CONFIG con datos de empresa, vendedor, condiciones comerciales y mensajes personalizables.
- **Utilidades de Limpieza**: Macros para preparar la hoja PEDIDOS antes de nuevos pedidos.

## 🚀 Instalación

1. Abrir el archivo Excel donde deseas instalar las macros
2. Presionar `Alt + F11` para abrir el editor de VBA
3. Crear un nuevo módulo y copiar el código de los archivos `.bas`
4. Crear la hoja CONFIG ejecutando la macro `CrearHojaDeConfiguracion`
5. Configurar los datos de la empresa en la hoja CONFIG

## 📁 Archivos del Proyecto

| Archivo | Descripción |
|---------|-------------|
| `GenerarXLSXCarta_v4.3.bas` | Macro para generar carta de cotización XLSX |
| `GenerarXLSXPedido_v4.3.bas` | Macro para generar pedido técnico XLSX |
| `Utilidades_Pedidos.bas` | Utilidades para limpieza y preparación de hoja PEDIDOS |
| `setup_config_sheet_ES.vba` | Macro para crear la hoja de configuración CONFIG |
| `GUIA_USO_MACRO_MEJORADA_v4.3.md` | Guía completa de uso |
| `INSTALACION_PASO_A_PASO_v4.3.txt` | Instrucciones de instalación |

## 📊 Uso

### Preparación Inicial
1. Ejecutar `CrearHojaDeConfiguracion` para crear la hoja CONFIG
2. Llenar los datos de la empresa en CONFIG!B6, B7, B10, B25, B26
3. Configurar datos del vendedor en CONFIG!B15, B16, B17
4. Definir condiciones comerciales en CONFIG!B20-B23, B28
5. (Opcional) Personalizar textos de introducción/despedida en CONFIG!B31, B32

### Flujo de Trabajo
1. Ejecutar `LimpiarHojaPedidos` o `PrepararNuevoPedido` para preparar la hoja
2. Pegar datos del sistema RPE en la hoja PEDIDOS (desde fila 5, columna C)
3. Ejecutar `GenerarXLSXCarta_v4_3` para generar la carta de cotización
4. Ejecutar `GenerarXLSXPedido_v4_3` para generar el pedido técnico

## 📝 Estructura de Datos

### Hoja PEDIDOS
| Celda | Contenido |
|-------|-----------|
| D2 | Nombre del cliente |
| D3 | Número de pedido |
| C5+ | CÓDIGO del producto |
| D5+ | DESCRIPCIÓN del producto |
| E5+ | CANTIDAD |
| F5+ | STOCK |
| G5+ | U/M (unidad de medida) |
| H5+ | PRECIO |
| I5+ | DESC1 (descuento 1) |
| J5+ | DESC2 (descuento 2) |

### Hoja CONFIG
| Celda | Contenido |
|-------|-----------|
| B6 | Nombre de la empresa |
| B7 | Dirección |
| B10 | Website |
| B15 | Nombre del vendedor |
| B16 | Teléfono del vendedor |
| B17 | Email del vendedor |
| B20 | Validez de cotización |
| B21 | Forma de pago |
| B22 | Plazo de entrega |
| B23 | Garantía |
| B25 | Datos de cuenta bancaria |
| B26 | Símbolo de moneda |
| B28 | Medios de pago |
| B31 | Texto de introducción (opcional) |
| B32 | Texto de despedida (opcional) |

## ⚙️ Requisitos

- Microsoft Excel 2016 o superior
- Habilitar macros de Excel
- Conocimientos básicos de VBA (opcional)

## 📦 Versión

**v4.3** (Enero 2026)
- Unificación de versiones de Pedido y Carta
- Eliminación de generación automática de PDF
- Configuración de columna C como formato texto
- Mensajes personalizables en CONFIG

## 📄 Licencia

Este proyecto es de uso interno. Consulte con el administrador para permisos de modificación.

## 👤 Autor

Desarrollado para uso empresarial. Contacto: ccusi@outlook.com

---

**Nota**: Los archivos .xlsm no están incluidos en el repositorio. Debe crear su propia plantilla de Excel y copiar las macros VBA.
