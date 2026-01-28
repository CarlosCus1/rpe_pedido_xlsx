# Guía de Uso - Sistema de Pedidos v4.3

## 📖 Introducción

Esta guía explica cómo utilizar el sistema de generación de documentos de pedidos en su versión 4.3, que incluye dos macros principales:

1. **GenerarXLSXPedido_v4.3** - Genera hoja técnica de pedido
2. **GenerarXLSXCarta_v4.3** - Genera carta de cotización en formato XLSX

## 🚀 Inicio Rápido

### Requisitos Previos

1. Archivo Excel habilitado para macros (`.xlsm`)
2. Hoja "CONFIG" configurada correctamente
3. Hoja "PEDIDOS" con datos del sistema RPE
4. Logotipo de empresa insertado en CONFIG con nombre "logo_empresa"

### Pasos Básicos

1. **Abrir el archivo** `Pedidos a Excel V4.x.xlsm`
2. **Verificar** que la hoja CONFIG tenga los datos de su empresa
3. **Pegar datos** del sistema RPE en la hoja PEDIDOS (fila 5 en adelante)
4. **Ejecutar** la macro deseada (Pedido o Carta)
5. **Abrir el archivo** generado en el escritorio

---

## 📊 Macro: GenerarXLSXPedido_v4.3

### Propósito
Genera un archivo XLSX técnico con formato de pedido, incluyendo:
- Logotipo de empresa
- Datos del cliente y pedido
- Tabla detallada con 12 columnas
- Estados de stock (con colores)
- Cálculos automáticos con fórmulas
- Totales superiores destacados

### Columnas Generadas

| Columna | Contenido | Formato |
|---------|-----------|---------|
| A | N° (índice) | Numérico |
| B | CANT. | Numérico |
| C | U/M | Texto |
| D | ARTICULO | Texto (preserva ceros) |
| E | DESCRIPCIÓN | Texto |
| F | STOCK | Estado con colores |
| G | VALOR VENTA UNITARIO | Moneda |
| H | DESC 1 | Porcentaje |
| I | DESC 2 | Porcentaje |
| J | VALOR VENTA | Moneda (fórmula) |
| K | PRECIO UNITARIO | Moneda (fórmula) |
| L | PRECIO VENTA | Moneda (fórmula) |

### Estados de Stock (Colores)

- 🔴 **Sin Stock** - Rojo oscuro
- 🟠 **Stock Insuficiente** - Rojo claro
- 🟡 **Stock Ajustado** - Amarillo
- 🟢 **Stock Disponible** - Verde

### Cómo Usar

1. Asegúrese de tener datos en la hoja PEDIDOS
2. Presione `Alt + F8` para abrir el diálogo de macros
3. Seleccione `GenerarXLSXPedido_v4_3`
4. Haga clic en "Ejecutar"
5. El archivo se guardará en el escritorio

---

## 📄 Macro: GenerarXLSXCarta_v4.3

### Propósito
Genera una carta de cotización profesional en formato XLSX, lista para:
- Imprimir directamente
- Guardar como PDF manualmente
- Enviar por correo electrónico

### Características

- ✅ **Sin PDF automático** - El usuario decide cuándo imprimir
- ✅ **Textos profesionales** - Introducción y despedida con formato
- ✅ **Tabla de Excel** - Con nombre "TablaProductos"
- ✅ **Fórmulas dinámicas** - Los totales se recalculan automáticamente
- ✅ **IGV visible** - En la sección de totales
- ✅ **Auto-ajuste** - Columnas B:G se ajustan automáticamente

### Estructura del Documento

```
[LOGO]                    [NOMBRE EMPRESA]
─────────────────────────────────────────────

COTIZACIÓN N°: XXX          Fecha: DD de Mes de AAAA

SEÑOR(ES): [Nombre Cliente]

Estimados:

Es un gusto saludarles. Les envío la propuesta comercial sobre los 
productos consultados.

Quedamos a su disposición para cualquier consulta:

┌──────┬────────┬─────────────┬───────┬─────┬──────────┬────────┐
│ ITEM │ CÓDIGO │ DESCRIPCIÓN │ CANT. │ U/M │ P. UNIT. │ TOTAL  │
├──────┼────────┼─────────────┼───────┼─────┼──────────┼────────┤
│  1   │  ...   │     ...     │   ... │ ... │   ...    │  ...   │
└──────┴────────┴─────────────┴───────┴─────┴──────────┴────────┘

                                    SUBTOTAL:    S/. X,XXX.XX
                                    IGV (18%):   S/. XXX.XX
                                    TOTAL:       S/. X,XXX.XX

CONDICIONES COMERCIALES
• Validez de la oferta: [días]
• Forma de pago: [condiciones]
• Plazo de entrega: [días]
• Garantía: [período]

MEDIOS DE PAGO
[Información de cuentas bancarias]

Agradecemos su interés y quedamos atentos a su aprobación.

Confiamos en que la calidad de nuestra marca sea de su agrado.

Atentamente,


[Nombre Vendedor]
[ cargo]
T: [Teléfono] | E: [Email]
```

### Cómo Usar

1. Verifique que CONFIG tenga los datos completos
2. Asegúrese de tener productos en PEDIDOS
3. Presione `Alt + F8`
4. Seleccione `GenerarXLSXCarta_v4_3`
5. Haga clic en "Ejecutar"
6. El archivo se guardará en el escritorio con nombre: `Cotizacion_[N°Pedido]_[Cliente].xlsx`

### Para Convertir a PDF

1. Abra el archivo XLSX generado
2. Vaya a **Archivo → Guardar como**
3. Seleccione formato **PDF**
4. Configure opciones de impresión si es necesario
5. Guarde el archivo

---

## ⚙️ Configuración (Hoja CONFIG)

### Datos Requeridos

| Celda | Contenido | Ejemplo |
|-------|-----------|---------|
| B6 | Nombre Empresa | "Mi Empresa S.A.C." |
| B7 | Dirección | "Av. Principal 123, Lima" |
| B10 | Sitio Web | "www.miempresa.com" |
| B15 | Nombre Vendedor | "Juan Pérez" |
| B16 | Teléfono | "999-888-777" |
| B17 | Email | "juan@miempresa.com" |
| B20 | Validez Cotización | "7 días" |
| B21 | Tipo de Pago | "Crédito 30 días" |
| B22 | Plazo Entrega | "Inmediata" |
| B23 | Garantía | "12 meses" |
| B25 | RUC | "20123456789" |
| B26 | Símbolo Moneda | "S/." o "$" |
| B28 | Medios de Pago | "BCP: 191-1234567-0-89\|CCI: 0021910123456789" |

### Logotipo

1. Inserte una imagen en la hoja CONFIG
2. Cambie el nombre a: `logo_empresa`
3. La imagen se copiará automáticamente a los documentos generados

---

## 🔧 Solución de Problemas

### Error: "La hoja 'CONFIG' no existe"

**Causa**: No se ha creado la hoja de configuración  
**Solución**: Ejecute la macro `CrearHojaDeConfiguracion`

### Error: "No se encontraron datos en la hoja PEDIDOS"

**Causa**: Los datos no están en la ubicación correcta  
**Solución**: Pegue los datos desde la fila 5, columna C

### Error: "Faltan datos del Cliente o N° de Pedido"

**Causa**: Celdas D2 o D3 de la hoja PEDIDOS están vacías  
**Solución**: Complete la información del cliente y número de pedido

### Los totales no se calculan

**Causa**: Excel tiene cálculo manual deshabilitado  
**Solución**: Presione `F9` para recalcular o cambie a cálculo automático

---

## 📞 Soporte

Si encuentra algún problema o tiene preguntas:

1. Revise que tiene la versión 4.3 de los macros
2. Verifique que la hoja CONFIG esté completa
3. Asegúrese de que los datos en PEDIDOS estén correctos

---

## 📋 Checklist Pre-Ejecución

- [ ] Archivo guardado como `.xlsm`
- [ ] Hoja CONFIG creada y completa
- [ ] Logotipo insertado con nombre "logo_empresa"
- [ ] Datos pegados en PEDIDOS desde fila 5
- [ ] Cliente y N° de Pedido completados (D2, D3)
- [ ] Macros habilitadas en Excel

---

**Versión:** 4.3  
**Fecha:** Enero 2026  
**Estado:** ✅ Documentación completa