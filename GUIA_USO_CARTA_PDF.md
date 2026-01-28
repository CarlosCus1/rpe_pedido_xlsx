# GUÍA DE USO - Macro GenerarCartaPDF v2.5

## 📋 Descripción General

La macro `GenerarCartaPDF` versión 2.5 genera cartas de cotización en formato PDF profesionales utilizando únicamente Excel, sin dependencia de Microsoft Word. Crea una hoja de plantilla temporal, la llena con datos de la empresa, cliente, productos y condiciones comerciales, y la exporta a PDF usando el método nativo `ExportAsFixedFormat` de Excel. Utiliza una paleta de colores corporativa tradicional ideal para contratos y documentos formales, con un encabezado minimalista que incluye solo el logo y el nombre de la empresa. Incluye un campo flexible "Medios de Pago" que permite incluir todos los tipos de cuentas y métodos de pago aceptados, y un mensaje personalizable breve.

## 🎯 Funcionalidades Principales

- ✅ Generación automática de cartas PDF profesionales
- ✅ **100% Excel nativo** - Sin dependencia de Microsoft Word
- ✅ Integración con logotipo y datos de empresa desde CONFIG
- ✅ Datos del cliente desde hoja PEDIDOS
- ✅ Tabla de productos con cálculos automáticos
- ✅ Cálculo de totales (subtotal, IGV, total con IGV)
- ✅ Condiciones comerciales personalizables
- ✅ Datos del vendedor en la carta
- ✅ Pie de página con información de la empresa
- ✅ Formato profesional con colores corporativos
- ✅ Guardado automático en carpeta "CartasPDF"
- ✅ **Soporte para tamaño A4 o Carta**
- ✅ **Hoja temporal eliminada automáticamente**

## 🔄 Cambios en la Versión 2.0

### Mejoras Principales
- **Eliminación de dependencia de Word**: Ya no requiere Microsoft Word instalado
- **Excel nativo**: Utiliza `ExportAsFixedFormat` de Excel para generar PDF
- **Hoja temporal**: Crea una hoja temporal que se elimina automáticamente
- **Mejor rendimiento**: Procesamiento más rápido sin automatización externa
- **Configuración de página**: Soporte para tamaño A4 o Carta
- **Menos errores**: Reduce problemas de compatibilidad entre versiones

### Cambios Técnicos
- Eliminados todos los objetos de Word (`Word.Application`, `Word.Document`)
- Nueva función `CrearHojaCarta()` para crear hoja temporal
- Nueva función `CopiarLogo()` para copiar logo entre hojas
- Nueva función `ConfigurarPagina()` para configurar impresión
- Mejor manejo de errores y limpieza de recursos

## 📊 Estructura de Datos de Entrada

### Hoja CONFIG
Debe contener los siguientes datos:

| Celda | Descripción | Ejemplo |
|-------|-------------|---------|
| A1:B3 | Logotipo (forma nombrada "logo_empresa") | Imagen del logo |
| B6 | Nombre de la Empresa | "CIP COMERCIAL" |
| B7 | Dirección | "Calle Principal 123, Lima, Perú" |
| B8 | Teléfono Empresa | "+51 1 2345678" |
| B9 | Email Empresa | "contacto@cipcomercial.com" |
| B10 | Website | "www.cipcomercial.com" |
| B15 | Nombre del Vendedor | "Juan Pérez García" |
| B16 | Teléfono del Vendedor | "999 123 456" |
| B17 | Email del Vendedor | "juan.perez@cipcomercial.com" |
| B20 | Validez de Cotización | "30 días a partir de la fecha" |
| B21 | Tipo de Pago | "Contado / Crédito 30 días" |
| B22 | Plazo de Entrega | "3-5 días hábiles" |
| B23 | Condición Especial 1 | "Garantía 12 meses" |
| B24 | Condición Especial 2 | "Transporte incluido" |
| B25 | Pie de Página | "CIP COMERCIAL S.A.C. - RUC: 20123456789" |
| B28 | Medios de Pago | "BCP Soles 191-12345678-0-00 | CCI: 002-191-001234567890-00 | Yape: 999 123 456" |
| B29 | Mensaje Personalizable (breve) | "Gracias por su preferencia. Esperamos poder servirle pronto." |

### Hoja PEDIDOS
Debe contener los siguientes datos:

| Celda | Descripción | Ejemplo |
|-------|-------------|---------|
| D2 | Nombre del Cliente | "Empresa Cliente S.A.C." |
| D3 | Número de Pedido | "COT-2024-001" |

**Productos (desde fila 5):**

| Columna | Descripción | Ejemplo |
|---------|-------------|---------|
| C | Artículo (Código) | "ART001" |
| D | Descripción | "Producto de ejemplo" |
| E | Cantidad | 10 |
| F | Stock | 50 |
| G | Unidad de medida | "UND" |
| H | Valor unitario | 100.00 |
| I | Descuento 1 (%) | 5.00 |
| J | Descuento 2 (%) | 2.00 |

## 🚀 Cómo Usar la Macro

### Paso 1: Preparar Datos
1. Abrir el libro Excel con las hojas CONFIG y PEDIDOS
2. En CONFIG, asegurar que:
   - Existe el logotipo como forma "logo_empresa" en A1:B3
   - Todos los datos de empresa están completos (B6-B10)
   - Datos del vendedor están completos (B15-B17)
   - Condiciones comerciales están definidas (B20-B24)
   - Pie de página está configurado (B25)
3. En PEDIDOS, colocar:
   - Nombre del cliente en D2
   - Número de pedido en D3
   - Productos desde la fila 5 (columnas C-J)

### Paso 2: Ejecutar la Macro
1. Presionar `Alt + F8` para abrir el ejecutor de macros
2. Seleccionar `GenerarCartaPDF`
3. Hacer clic en "Ejecutar"

### Paso 3: Resultado
- Se genera automáticamente un archivo PDF en la carpeta "CartasPDF"
- Nombre del archivo: `Cotizacion_[NúmeroPedido]_[Cliente].pdf`
- El PDF se abre automáticamente después de la generación
- La hoja temporal se elimina automáticamente
- Mensaje de confirmación con la ruta del archivo generado

## 📄 Estructura de la Carta PDF Generada

### 1. Encabezado Minimalista
- **Logotipo**: Imagen de la empresa (izquierda, en A1)
- **Nombre de la empresa**: En formato grande, en negrita, alineado a la derecha
- **Línea separadora**: Línea horizontal gris debajo del logo y nombre

**Nota:** El RUC y la página web se muestran en el pie de página (footer) del documento, no en el encabezado.

### 2. Fecha y Referencia
- **Fecha**: Fecha actual en formato "dd de MMMM de yyyy"
- **Cotización N°**: Número de pedido/cotización

### 3. Cliente
- **Señor(es):**: Etiqueta de destinatario
- **Nombre del cliente**: Nombre del cliente en negrita

### 4. Presentación
- **Saludo**: "De nuestra mayor consideración:"
- **Introducción**: Texto de presentación de la empresa y calidad de productos
- **Transición**: "A continuación, detallamos los productos solicitados:"

### 5. Tabla de Productos
La tabla incluye las siguientes columnas:

| Columna | Descripción | Formato |
|---------|-------------|---------|
| N° | Número de línea | Centrado |
| Código | Código del artículo | Izquierda |
| Producto | Descripción del producto | Izquierda |
| Cantidad | Cantidad solicitada | Centrado |
| U/M | Unidad de medida | Centrado |
| Precio Unit. | Precio unitario con descuentos | Derecha |
| Total | Total de línea (cantidad × precio) | Derecha |

**Totales:**
- **Subtotal sin IGV**: Suma de todos los totales de línea
- **IGV (18%)**: Impuesto General a las Ventas
- **TOTAL CON IGV**: Subtotal + IGV

### 6. Mensaje Personalizable (Opcional)
- Mensaje personalizado desde CONFIG!B29
- Aparece después de la tabla de productos
- Permite agregar un mensaje específico para cada cliente
- Debe ser breve (1-2 líneas máximo)
- Si está vacío, no se muestra

### 7. Agradecimiento
- Texto de agradecimiento por la preferencia
- Disposición para consultas
- "Atentamente,"

### 8. Datos del Vendedor
- **Nombre**: Nombre del vendedor en negrita
- **Teléfono**: Número de contacto del vendedor
- **Email**: Correo electrónico del vendedor

### 9. Condiciones Comerciales
Lista de condiciones con viñetas:
- Validez de la cotización
- Forma de pago
- Plazo de entrega
- Condición especial 1 (opcional)
- Condición especial 2 (opcional)
- Medios de pago (opcional, desde CONFIG!B28)

### 10. Pie de Página
- Información de la empresa (RUC, dirección)
- Website de la empresa
- Centrado y en negrita

## 🎨 Estilo y Formato

### Paleta de Colores Corporativa Tradicional (v2.3)
La macro utiliza una paleta de colores profesional y tradicional, ideal para contratos y documentos formales:

| Elemento | Color | Hex | RGB | Uso |
|----------|-------|-----|-----|-----|
| **Navy Blue** | Azul Marino | #001F3F | (0, 31, 63) | Encabezados de tabla |
| **Ivory** | Marfil | #F8F8F2 | (248, 248, 242) | Filas alternas de tabla |
| **Charcoal** | Carbón | #333333 | (51, 51, 51) | Sección de totales |
| **Blanco** | White | #FFFFFF | (255, 255, 255) | Texto sobre fondos oscuros |

**Detalles de uso:**
- **Encabezados de tabla**: Navy Blue (#001F3F) con texto blanco
- **Filas alternas**: Ivory (#F8F8F2) para facilitar lectura
- **Totales**: Charcoal (#333333) con texto blanco
- **Texto general**: Negro (#000000) sobre fondo blanco

**Ventajas de esta paleta:**
- ✅ Profesional y tradicional para contratos
- ✅ Alto contraste para mejor legibilidad
- ✅ Adecuado para impresión en blanco y negro
- ✅ Transmite seriedad y confianza
- ✅ Cumple con estándares corporativos

### Tipografía
- **Fuente**: Calibri (profesional)
- **Tamaños**:
  - Título empresa: 16 pt
  - Encabezados: 12 pt
  - Texto normal: 11 pt
  - Texto pequeño: 10 pt

### Alineación
- **Centrado**: Logotipo, nombre empresa, encabezados de tabla
- **Izquierda**: Texto del cuerpo, descripciones
- **Derecha**: Valores numéricos, totales

## 📐 Configuración de Página

### Tamaño de Papel
La macro está configurada para **tamaño A4** por defecto. Para cambiar a **Carta**, modificar la línea:

```vba
.PaperSize = xlPaperA4  ' Cambiar a xlPaperLetter para Carta
```

En el procedimiento `ConfigurarPagina()`.

### Márgenes
- **Izquierdo**: 0.5 pulgadas (36 puntos)
- **Derecho**: 0.5 pulgadas (36 puntos)
- **Superior**: 0.5 pulgadas (36 puntos)
- **Inferior**: 0.5 pulgadas (36 puntos)
- **Encabezado**: 0.25 pulgadas (18 puntos)
- **Pie de página**: 0.25 pulgadas (18 puntos)

### Orientación
- **Vertical** (Portrait) - Predeterminado

### Ajuste de Página
- **Ajustar a 1 página de ancho**
- **Zoom automático** para contenido

## ⚠️ Consideraciones Importantes

### Requisitos del Sistema
- ✅ Microsoft Excel 2010 o superior
- ✅ Habilitar macros en Excel
- ✅ Permisos de escritura en la carpeta del proyecto
- ❌ **NO requiere Microsoft Word** (novedad v2.0)

### Validaciones Automáticas
- ✅ Verificación de existencia de hoja CONFIG
- ✅ Verificación de existencia de hoja PEDIDOS
- ✅ Validación de nombre de empresa en CONFIG!B6
- ✅ Validación de nombre de vendedor en CONFIG!B15
- ✅ Validación de cliente en PEDIDOS!D2
- ✅ Validación de número de pedido en PEDIDOS!D3
- ✅ Verificación de productos en PEDIDOS

### Ventajas de la Versión 2.0
- ✅ **Sin dependencia de Word**: Funciona solo con Excel
- ✅ **Más rápido**: Procesamiento nativo de Excel
- ✅ **Menos errores**: Reduce problemas de compatibilidad
- ✅ **Más ligero**: No requiere automatización externa
- ✅ **Hoja temporal**: Se elimina automáticamente
- ✅ **Configuración flexible**: Fácil cambiar tamaño de papel

### Limitaciones
- ❌ Logotipo debe estar como forma nombrada "logo_empresa"
- ❌ Datos deben seguir formato específico
- ❌ No compatible con Excel versiones anteriores a 2010

### Manejo de Errores
- Mensajes descriptivos para cada tipo de error
- Recuperación automática en caso de fallos
- Limpieza de hoja temporal al finalizar
- Restauración de configuración de Excel

## 🔧 Solución de Problemas

### Logotipo no aparece
- Verificar que existe forma "logo_empresa" en CONFIG
- Asegurar que la forma no esté oculta
- Verificar que el logo tenga un nombre correcto

### Error "No se encontró hoja CONFIG"
- Verificar que la hoja se llame exactamente "CONFIG"
- Revisar que no haya espacios adicionales en el nombre

### Error "No se encontró hoja PEDIDOS"
- Verificar que la hoja se llame exactamente "PEDIDOS"
- Revisar que no haya espacios adicionales en el nombre

### Error "Falta el nombre de la empresa"
- Completar el campo CONFIG!B6 con el nombre de la empresa
- Asegurar que no esté vacío

### Error "Falta el nombre del vendedor"
- Completar el campo CONFIG!B15 con el nombre del vendedor
- Asegurar que no esté vacío

### Error "No hay productos en la hoja PEDIDOS"
- Verificar que haya productos desde la fila 5
- Asegurar que la columna C tenga datos de artículos

### Archivo PDF no se guarda
- Verificar permisos de escritura en la carpeta del proyecto
- Cerrar archivos PDF con nombres similares
- Verificar que la carpeta "CartasPDF" no esté bloqueada

### Hoja temporal no se elimina
- Verificar que no haya otra hoja con el mismo nombre
- Cerrar otros archivos de Excel que puedan estar bloqueando
- Reiniciar Excel si persiste el problema

### PDF no se abre automáticamente
- Verificar que el visor de PDF predeterminado esté configurado
- Revisar que `OpenAfterPublish:=True` esté activo en el código
- Abrir manualmente el archivo desde la carpeta "CartasPDF"

### Totales incorrectos
- Verificar que los valores numéricos sean válidos
- Revisar que las cantidades sean números positivos
- Verificar que los descuentos estén en porcentaje (0-100)

### Tamaño de papel incorrecto
- Modificar `.PaperSize` en `ConfigurarPagina()`
- Usar `xlPaperA4` para A4 o `xlPaperLetter` para Carta
- Guardar y volver a ejecutar la macro

## 📁 Estructura de Archivos

### Archivos Generados
```
[Proyecto]/
├── CartasPDF/                    ← Carpeta creada automáticamente
│   ├── Cotizacion_COT-2024-001_EmpresaCliente.pdf
│   ├── Cotizacion_COT-2024-002_OtroCliente.pdf
│   └── ...
├── GenerarCartaPDF.bas          ← Macro principal v2.0
├── CONFIG                        ← Hoja de configuración
└── PEDIDOS                       ← Hoja de pedidos
```

### Hoja Temporal
- **Nombre**: `Carta_Temporal_hhmmss` (ej: `Carta_Temporal_143025`)
- **Ubicación**: Se crea al final del libro
- **Duración**: Existe solo durante la ejecución
- **Eliminación**: Se elimina automáticamente después de exportar PDF

## 📞 Soporte

Para soporte técnico o reportes de bugs, proporcionar:
- Versión de Excel
- Descripción del error
- Captura de pantalla si es posible
- Datos de ejemplo que causan el problema

## 🔄 Actualizaciones

### Historial de Versiones

#### v2.0 (Actual)
- **Refactorización completa**: Eliminación de dependencia de Word
- **Excel nativo**: Uso de `ExportAsFixedFormat`
- **Hoja temporal**: Creación y eliminación automática
- **Configuración de página**: Soporte para A4/Carta
- **Mejor rendimiento**: Procesamiento más rápido
- **Menos errores**: Reducción de problemas de compatibilidad

#### v1.0
- Generación de cartas PDF con Word automation
- Integración con CONFIG y PEDIDOS
- Cálculo automático de totales
- Formato profesional

### Compatibilidad
- ✅ Compatible con estructura de datos existente
- ✅ Mantiene funcionalidad de otras macros
- ✅ No interfiere con otros procesos
- ✅ Requiere solo Excel (no Word)

## 💡 Consejos de Uso

### Mejores Prácticas
1. **Mantener datos actualizados**: Actualizar CONFIG cuando cambie información de la empresa
2. **Usar nombres descriptivos**: Para clientes y pedidos facilitar la organización
3. **Verificar datos antes de generar**: Revisar que todos los campos estén completos
4. **Organizar PDFs**: La carpeta "CartasPDF" se crea automáticamente para mantener orden
5. **Backup de PDFs**: Considerar hacer copias de seguridad de las cartas generadas
6. **Verificar hoja temporal**: Si algo falla, verificar que no queden hojas temporales

### Personalización

#### Cambiar Tamaño de Papel
En el procedimiento `ConfigurarPagina()`:
```vba
.PaperSize = xlPaperA4      ' Para A4
.PaperSize = xlPaperLetter  ' Para Carta
```

#### Cambiar Fuente
Modificar las constantes al inicio del módulo:
```vba
Private Const FONT_NAME As String = "Arial"  ' Cambiar a Arial
Private Const FONT_SIZE_NORMAL As Integer = 12  ' Cambiar tamaño
```

#### Cambiar Colores
Modificar las constantes de colores:
```vba
Private Const COLOR_HEADER_BG As Long = 255      ' Rojo
Private Const COLOR_HEADER_TEXT As Long = 0      ' Negro
Private Const COLOR_ROW_ALT As Long = 16777215   ' Blanco
```

#### Cambiar Porcentaje de IGV
Modificar la constante:
```vba
Private Const IGV_RATE As Double = 0.18  ' 18%
```

### Integración con Otras Macros
Esta macro puede integrarse con:
- `Procesar_Pedido_Sistema`: Para generar carta automáticamente después de procesar
- `GenerarXLSXPedido_v4.2`: Para generar ambos formatos (XLSX y PDF)
- Macros personalizadas: Para flujos de trabajo específicos

## 📝 Ejemplo de Uso Completo

### Escenario: Generar cotización para cliente nuevo

1. **Configurar empresa (una sola vez)**
   - Abrir hoja CONFIG
   - Insertar logo en A1:B3 y nombrarlo "logo_empresa"
   - Completar B6-B10 con datos de la empresa
   - Completar B15-B17 con datos del vendedor
   - Completar B20-B24 con condiciones comerciales
   - Completar B25 con pie de página
   - Completar B28 con medios de pago (opcional)
   - Completar B29 con mensaje personalizable (opcional)

2. **Preparar pedido**
   - Abrir hoja PEDIDOS
   - Colocar "Empresa Cliente S.A.C." en D2
   - Colocar "COT-2024-001" en D3
   - Pegar productos desde fila 5:
     - C: ART001, D: Producto A, E: 10, F: 50, G: UND, H: 100.00, I: 5.00, J: 2.00
     - C: ART002, D: Producto B, E: 5, F: 30, G: UND, H: 200.00, I: 0.00, J: 0.00

3. **Generar carta PDF**
   - Presionar Alt + F8
   - Seleccionar `GenerarCartaPDF`
   - Hacer clic en "Ejecutar"

4. **Resultado**
   - Se crea hoja temporal `Carta_Temporal_143025`
   - Se genera archivo: `Cotizacion_COT-2024-001_EmpresaClienteSAC.pdf`
   - El PDF se abre automáticamente
   - La hoja temporal se elimina
   - Mensaje de confirmación muestra la ruta completa

## 🎓 Referencias Técnicas

### Constantes Configurables
```vba
Private Const IGV_RATE As Double = 0.18        ' Porcentaje de IGV
Private Const FONT_NAME As String = "Calibri"  ' Fuente predeterminada
Private Const FONT_SIZE_NORMAL As Integer = 11 ' Tamaño de texto normal
Private Const FONT_SIZE_SMALL As Integer = 10  ' Tamaño de texto pequeño
Private Const FONT_SIZE_TITLE As Integer = 14  ' Tamaño de título
Private Const FONT_SIZE_HEADER As Integer = 12 ' Tamaño de encabezados
Private Const FONT_SIZE_LARGE As Integer = 16  ' Tamaño grande
```

### Constantes de Colores Corporativos (v2.3)
```vba
Private Const COLOR_HEADER_BG As Long = 4144959      ' Navy Blue (#001F3F)
Private Const COLOR_HEADER_TEXT As Long = 16777215   ' Blanco (#FFFFFF)
Private Const COLOR_ROW_ALT As Long = 16316670       ' Ivory (#F8F8F2)
Private Const COLOR_TOTAL_BG As Long = 3355443       ' Charcoal (#333333)
Private Const COLOR_TOTAL_TEXT As Long = 16777215    ' Blanco (#FFFFFF)
```

### Funciones Principales
- `GenerarCartaPDF()`: Procedimiento principal
- `CrearHojaCarta()`: Crea hoja temporal
- `LlenarCarta()`: Llena la hoja con datos
- `LeerProductos()`: Lee productos desde PEDIDOS
- `CalcularTotales()`: Calcula subtotal, IGV y total
- `CopiarLogo()`: Copia logo desde CONFIG
- `ConfigurarPagina()`: Configura página para impresión
- `LimpiarNombreArchivo()`: Limpia caracteres inválidos

### Métodos de Excel Utilizados
- `Sheets.Add()`: Crear nueva hoja
- `ExportAsFixedFormat()`: Exportar a PDF
- `Range.Merge()`: Fusionar celdas
- `Range.WrapText`: Ajustar texto
- `PageSetup`: Configuración de página

## 🔍 Comparación v1.0 vs v2.0

| Característica | v1.0 | v2.0 |
|---------------|------|------|
| Dependencia de Word | Sí | No |
| Método de exportación | Word automation | Excel nativo |
| Hoja temporal | No | Sí |
| Rendimiento | Medio | Alto |
| Errores de compatibilidad | Posibles | Reducidos |
| Configuración de página | Limitada | Flexible |
| Tamaño de papel | Fijo | A4/Carta |
| Requisitos | Excel + Word | Solo Excel |

## 📊 Flujo de Ejecución

```
1. Validar hojas (CONFIG, PEDIDOS)
   ↓
2. Leer datos de CONFIG (empresa, vendedor, condiciones)
   ↓
3. Leer datos de PEDIDOS (cliente, número, productos)
   ↓
4. Calcular totales (subtotal, IGV, total)
   ↓
5. Crear hoja temporal
   ↓
6. Copiar logo desde CONFIG
   ↓
7. Llenar hoja con todos los datos
   ↓
8. Configurar página (A4/Carta)
   ↓
9. Exportar a PDF
   ↓
10. Eliminar hoja temporal
   ↓
11. Mostrar mensaje de éxito
```

---

## 🆕 Novedades en la Versión 2.5

### Simplificación de Campos de Pago
- **Campo unificado B28**: "Medios de Pago" reemplaza los campos anteriores de cuenta bancaria y cuenta corriente
- **Mayor flexibilidad**: Permite incluir todos los tipos de cuentas y métodos de pago en un solo campo
- **Ejemplos de uso**: Cuentas bancarias (BCP, Interbank, etc.), Yape, Plin, transferencias, etc.
- **Formato libre**: Puede incluir múltiples líneas separadas por " | " o saltos de línea
- **Opcional**: Si está vacío, no se muestra en las condiciones comerciales

### Mensaje Personalizable Breve
- **Campo B29**: Mensaje personalizable que aparece en el cuerpo de la carta después de la tabla de productos
- **Optimización**: Debe ser breve (1-2 líneas máximo)
- **Mejor legibilidad**: Evita textos excesivamente largos en el cuerpo de la carta

### Mejoras Implementadas
- Simplificación de la estructura de CONFIG (de 3 campos de pago a 1)
- Mayor flexibilidad para incluir todos los medios de pago aceptados
- Reducción de redundancia en la configuración
- Mejor organización del contenido de la carta

### Campos de la Versión 2.5
- **B28 - Medios de Pago**: Se muestra en las condiciones comerciales (opcional)
- **B29 - Mensaje Personalizable (breve)**: Se muestra en el cuerpo de la carta después de la tabla de productos (opcional)
- Footer se repite automáticamente en todas las páginas del documento

### Encabezado Minimalista (v2.3)
- **Diseño simplificado**: Logo a la izquierda + Nombre de la empresa en negrita a la derecha
- **Línea separadora**: Línea horizontal gris debajo del logo y nombre
- **Eliminación de información redundante**: Dirección, teléfono y email eliminados del encabezado
- **Mayor limpieza visual**: Diseño más limpio y profesional
- **RUC y web en footer**: Información corporativa se muestra en el pie de página del documento

### Paleta de Colores Corporativa Tradicional (v2.2)
- **Navy Blue (#001F3F)**: Para encabezados de tabla - transmite profesionalismo y confianza
- **Ivory (#F8F8F2)**: Para filas alternas - facilita la lectura sin distraer
- **Charcoal (#333333)**: Para sección de totales - destaca información importante
- **Texto blanco en totales**: Mejor contraste y legibilidad

---

**Versión**: 2.5 (Simplificación de campos de pago a "Medios de Pago")
**Fecha**: 2024
**Autor**: Sistema de Gestión de Pedidos
**Licencia**: Uso interno