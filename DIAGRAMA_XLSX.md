# Diagrama Visual - Archivo XLSX Generado v4.2

## 📄 Estructura del Archivo XLSX (Pedido)

```
╔══════════════════════════════════════════════════════════════════════════════════════════════════════════╗
║  [LOGO]                              CORPORACIÓN DE INDUSTRIAS PLÁSTICAS S.A.                           ║
║  (A1:A3)                             (C1:E1 - Negrita, Tamaño 16, Azul #003366)                        ║
╠══════════════════════════════════════════════════════════════════════════════════════════════════════════╣
║                                                                                                          ║
║   ┌─────────────────────────────────────────────────┐                                                    ║
║   │ CLIENTE:        CLIENTE                         │  [D2:E2]                                           ║
║   │ PEDIDO:         PEDIDO                          │  [D3:E3]                                           ║
║   └─────────────────────────────────────────────────┘                                                    ║
║                                                                                                          ║
║   ┌─────────────────────────────────────────────────┐                                                    ║
║   │ Total con Stock:        S/. 1,234.56          │  [K1:L1] - Verde #006600                           ║
║   │ Total General:          S/. 2,345.67          │  [K2:L2] - Azul #003366                            ║
║   │ Total Descuentos:       S/.    111.11         │  [K3:L3] - Rojo #660000                            ║
║   └─────────────────────────────────────────────────┘                                                    ║
║                                                                                                          ║
║   ┌────┬────────┬──────┬──────────┬─────────────────┬──────────┬────────────┬───────┬───────┬───────────┬────────────┬───────────┐
║   │ N° │ CANT.  │ U/M  │ ARTICULO │ DESCRIPCIÓN     │ STOCK    │ VALOR VENTA│ DESC  │ DESC  │ VALOR     │ PRECIO     │ PRECIO    │
║   │    │        │      │          │                 │          │ UNITARIO   │ 1     │ 2     │ VENTA     │ UNITARIO   │ VENTA     │
║   ├────┼────────┼──────┼──────────┼─────────────────┼──────────┼────────────┼───────┼───────┼───────────┼────────────┼───────────┤
║   │ 1  │ 10     │ UND  │ 01240    │ Producto A      │ Stock    │      95.00 │  5%   │  0%   │    902.50 │    112.10  │  1,064.95 │
║   │    │        │      │          │                 │ Disp.    │            │       │       │           │            │           │
║   ├────┼────────┼──────┼──────────┼─────────────────┼──────────┼────────────┼───────┼───────┼───────────┼────────────┼───────────┤
║   │ 2  │ 5      │ UND  │ 011019   │ Producto B      │ Stock    │     200.00 │ 10%   │  0%   │    900.00 │    212.40  │  1,062.00 │
║   │    │        │      │          │                 │ Ajustado │            │       │       │           │            │           │
║   ├────┼────────┼──────┼──────────┼─────────────────┼──────────┼────────────┼───────┼───────┼───────────┼────────────┼───────────┤
║   │ 3  │ 20     │ UND  │ 03521    │ Producto C      │ Stock    │      50.00 │  0%   │  0%   │  1,000.00 │     59.00  │  1,180.00 │
║   │    │        │      │          │                 │ Insuf.   │            │       │       │           │            │           │
║   └────┴────────┴──────┴──────────┴─────────────────┴──────────┴────────────┴───────┴───────┴───────────┴────────────┴───────────┘
║                                                                                                          ║
║   [Filas de datos continuas...]                                                                         ║
║                                                                                                          ║
║   ════════════════════════════════════════════════════════════════════════════════════════════════════════
║   FILA 5: Encabezados de tabla (congelados al visualizar)                                                ║
║   FILA 6+: Datos de productos                                                                            ║
╚══════════════════════════════════════════════════════════════════════════════════════════════════════════╝
```

## 📐 Estructura Detallada por Secciones

### 1. Encabezado (Logo y Empresa) - Filas 1-3

| Rango | Contenido | Fuente | Formato |
|-------|-----------|--------|---------|
| A1:A3 | Logo de la empresa (60 puntos = 2.10 cm) | CONFIG!logo_empresa | Imagen |
| C1:E1 | Nombre de la empresa | CONFIG!B6 | Negrita, Tamaño 16, Color #003366 |

### 2. Información de Cliente y Pedido - Filas 2-3

| Rango | Contenido | Fuente | Formato |
|-------|-----------|--------|---------|
| C2 | "CLIENTE" (etiqueta) | Fijo | Negrita, Tamaño 11, Fondo gris #EBEBEB |
| D2:E2 | Nombre del cliente | PEDIDOS!D2 | Texto, Fondo gris #EBEBEB |
| C3 | "PEDIDO" (etiqueta) | Fijo | Negrita, Tamaño 11, Fondo gris #EBEBEB |
| D3:E3 | Número de pedido | PEDIDOS!D3 | Texto, Fondo gris #EBEBEB |

### 3. Totales Superiores - Filas 1-3 (Columnas K-L)

| Celda | Contenido | Color | Formato |
|-------|-----------|-------|---------|
| K1 | "Total con Stock:" | Verde #006600 | Negrita, Derecha |
| L1 | Suma de productos con stock disponible | Verde #006600 | Soles, Borde grueso |
| K2 | "Total General:" | Azul #003366 | Negrita, Derecha |
| L2 | Suma de todos los productos | Azul #003366 | Soles, Borde doble |
| K3 | "Total Descuentos:" | Rojo #660000 | Negrita, Derecha |
| L3 | Total de descuentos aplicados | Rojo #660000 | Soles, Borde grueso |

### 4. Tabla de Productos - Filas 5 en adelante

#### Encabezados de Tabla (Fila 5)

| Columna | Encabezado | Alineación | Color Fondo |
|---------|------------|------------|-------------|
| A | N° | Centro | Azul #595959 |
| B | CANT. | Centro | Azul #595959 |
| C | U/M | Centro | Azul #595959 |
| D | ARTICULO | Centro | Azul #595959 |
| E | DESCRIPCIÓN | Centro | Azul #595959 |
| F | STOCK | Centro | Verde oscuro #64964A |
| G | VALOR VENTA UNITARIO | Derecha | Azul #595959 |
| H | DESC 1 | Centro | Azul #595959 |
| I | DESC 2 | Centro | Azul #595959 |
| J | VALOR VENTA | Centro | Gris oscuro #4B4B4B |
| K | PRECIO UNITARIO | Derecha | Gris claro #CFD5EA |
| L | PRECIO VENTA | Derecha | Gris claro #CFD5EA |

#### Formato Condicional - Columna STOCK (F)

| Valor | Color Fondo | Color Texto |
|-------|-------------|-------------|
| "Sin Stock" | Rojo claro #FF9696 | Rojo oscuro #640000, Negrita |
| "Stock Insuficiente" | Rojo muy claro #FFC8C8 | Rojo oscuro #963232, Negrita |
| "Stock Ajustado" | Amarillo #FFFFB4 | Amarillo oscuro #969600, Negrita |
| "Stock Disponible" | Verde claro #B4FFB4 | Verde oscuro #009600, Negrita |

#### Formato de Columnas

| Columna | Formato | Descripción |
|---------|---------|-------------|
| A | 0 | Número entero |
| B | 0 | Número entero |
| C | @ | Texto |
| D | @ | Texto (preserva ceros a la izquierda) |
| E | @ | Texto |
| F | @ | Texto |
| G | [$S/-409] #,##0.00 | Soles |
| H | 0.00% | Porcentaje |
| I | 0.00% | Porcentaje |
| J | [$S/-409] #,##0.00 | Soles |
| K | [$S/-409] #,##0.00 | Soles |
| L | [$S/-409] #,##0.00 | Soles |

## 🎨 Paleta de Colores

| Elemento | Color Hex | RGB | Uso |
|----------|-----------|-----|-----|
| Encabezados tabla | #595959 | 89, 85, 89 | Fondo azul grisáceo oscuro |
| Texto encabezados | #FFFFFF | 255, 255, 255 | Blanco |
| Zona superior | #EBEBEB | 235, 235, 235 | Fondo gris muy claro |
| Columnas calculadas | #CFD5EA | 207, 213, 234 | Fondo gris claro azulado |
| Total General | #BFDFFF | 191, 223, 255 | Azul medio claro |
| Stock Disponible | #B4FFB4 | 180, 255, 180 | Verde claro |
| Stock Ajustado | #FFFFB4 | 255, 255, 180 | Amarillo |
| Stock Insuficiente | #FFC8C8 | 255, 200, 200 | Rojo claro |
| Sin Stock | #FF9696 | 255, 150, 150 | Rojo oscuro claro |
| Índice N° | #D3D3D4 | 211, 211, 212 | Gris muy claro |
| Total Stock | #DCFFDC | 220, 255, 220 | Verde muy claro |
| Total General | #003366 | 0, 51, 102 | Azul oscuro |
| Total Descuentos | #660000 | 102, 0, 0 | Rojo oscuro |

## 📊 Flujo de Datos

```
┌─────────────────────────────────────────────────────────────────────────────────────────┐
│                           HOJA "PEDIDOS" (Origen)                                       │
├─────────────────────────────────────────────────────────────────────────────────────────┤
│ D2: CLIENTE          →  D2:E2 en XLSX                                                   │
│ D3: PEDIDO           →  D3:E3 en XLSX                                                   │
│ C5:Jn (datos)        →  Tabla en XLSX (12 columnas calculadas)                          │
└─────────────────────────────────────────────────────────────────────────────────────────┘
                                            ↓
┌─────────────────────────────────────────────────────────────────────────────────────────┐
│                           HOJA "CONFIG" (Referencia)                                    │
├─────────────────────────────────────────────────────────────────────────────────────────┤
│ B6: Nombre empresa   →  C1:E1 en XLSX                                                   │
│ logo_empresa         →  A1:A3 en XLSX (imagen)                                          │
└─────────────────────────────────────────────────────────────────────────────────────────┘
                                            ↓
┌─────────────────────────────────────────────────────────────────────────────────────────┐
│                           ARCHIVO XLSX GENERADO (Salida)                                │
├─────────────────────────────────────────────────────────────────────────────────────────┤
│ Filas 1-3: Encabezado con logo, empresa, cliente, pedido y totales                     │
│ Filas 5-n: Tabla de productos con cálculos y formato condicional                        │
│ Guardado en: Escritorio como "CLIENTE-PEDIDO.xlsx"                                      │
└─────────────────────────────────────────────────────────────────────────────────────────┘
```

## 🔄 Proceso de Generación

```
1. Leer datos de PEDIDOS (D2, D3, C5:Jn)
2. Leer logo y empresa de CONFIG
3. Crear nueva hoja "PEDIDO"
4. Insertar logo en A1:A3 (60 puntos altura)
5. Escribir nombre empresa en C1:E1
6. Copiar cliente y pedido a D2:E3
7. Crear tabla de productos (12 columnas)
8. Aplicar formato condicional a STOCK
9. Calcular totales superiores
10. Congelar paneles en fila 5
11. Guardar como XLSX en escritorio
12. Eliminar hoja temporal "PEDIDO"
```

## 📝 Columnas de la Tabla (Mapeo)

| # | Columna XLSX | Fuente PEDIDOS | Descripción |
|---|--------------|----------------|-------------|
| 1 | N° | Calculado (índice) | Número secuencial |
| 2 | CANT. | Columna E | Cantidad solicitada |
| 3 | U/M | Columna G | Unidad de medida |
| 4 | ARTICULO | Columna C | Código (texto) |
| 5 | DESCRIPCIÓN | Columna D | Descripción del producto |
| 6 | STOCK | Columna F (vs CANT) | Estado de stock |
| 7 | VALOR VENTA UNITARIO | Columna H | Precio sin IVA |
| 8 | DESC 1 | Columna I | Descuento 1 |
| 9 | DESC 2 | Columna J | Descuento 2 |
| 10 | VALOR VENTA | Calculado | Cant × Unit × (1-D1) × (1-D2) |
| 11 | PRECIO UNITARIO | Calculado | Unit × (1-D1) × (1-D2) × 1.18 |
| 12 | PRECIO VENTA | Calculado | Valor Venta × 1.18 |

## 💡 Características Principales

- **Logo**: Altura fija de 2.10 cm (60 puntos)
- **Totales**: 3 tipos (Stock, General, Descuentos) con colores distintivos
- **Tabla**: 12 columnas con formato condicional en STOCK
- **Formato**: Números en soles con símbolo S/.
- **Interactividad**: Congelar paneles en fila 5
- **Stock Informativo**: No modifica cantidades, solo indica disponibilidad
- **IVA**: 18% aplicado en columnas K y L

---

**Versión del diagrama**: 4.2
**Fecha**: Enero 2026
**Autor**: Sistema de Gestión de Pedidos
**Actualizaciones**: Totales superiores, formato condicional STOCK, logo desde CONFIG
