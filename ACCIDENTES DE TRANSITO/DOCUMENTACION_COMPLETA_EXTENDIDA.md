# 📊 DOCUMENTACIÓN TÉCNICA Y FUNCIONAL: ANÁLISIS DE ACCIDENTES DE TRANSITO

![Estado](https://img.shields.io/badge/Estado-Producción-success?style=for-the-badge)
![Power BI](https://img.shields.io/badge/Tecnología-Power%20BI-f2c811?style=for-the-badge&logo=powerbi&logoColor=black)
![DAX](https://img.shields.io/badge/Lenguaje-DAX-blue?style=for-the-badge)

---

Esta documentación abarca la totalidad del proyecto de análisis de accidentes de tránsito. Se detalla la estructura de datos, el modelado, todas las fórmulas DAX implementadas y la lógica completa de los controles visuales HTML.

---

## 📑 INDICE DE CONTENIDOS

1.  [Contexto del Dataset](#-contexto-del-dataset)
2.  [Estructura y Diccionario de Datos](#-estructura-y-diccionario-de-datos)
3.  [Modelo de Datos](#-modelo-de-datos)
4.  [Ingeniería de Datos (DAX)](#-ingeniería-de-datos-dax)
    *   [Columnas Calculadas](#columnas-calculadas)
    *   [Tabla Calendario](#tabla-calendario)
5.  [Medidas Analíticas](#-medidas-analíticas)
    *   [Métricas Base](#métricas-base)
    *   [Métricas de Víctimas](#métricas-de-víctimas)
    *   [KPIs Estratégicos](#kpis-estratégicos)
    *   [Análisis Temporal](#análisis-temporal)
    *   [Análisis Geográfico](#análisis-geográfico)
    *   [Causas y Tipología](#causas-y-tipología)
6.  [Visualización Avanzada (HTML & CSS)](#-visualización-avanzada-html--css)
    *   [Banner Premium](#1-banner-premium-kpis-principales)
    *   [Grid de Detalle (Vehículos/Personas)](#2-grid-de-detalle-vehículospersonas)
    *   [Brújula de Riesgo](#3-brújula-de-riesgo)
    *   [Heatmap Semanal](#4-heatmap-semanal)
    *   [Ranking Top 10 Causas](#5-ranking-top-10-causas)
    *   [Bitácora de Siniestros](#6-bitácora-de-siniestros-críticos)
    *   [Radar de Víctimas](#7-radar-de-víctimas)
    *   [Embudo de Letalidad](#8-embudo-de-letalidad)
    *   [Burbujas de Gravedad](#9-burbujas-de-gravedad)
7.  [Insights y Conclusiones](#-insights-y-conclusiones)

---

## 📁 CONTEXTO DEL DATASET

**Nombre:** `BD-Accidentes-Carreteras`  
**Origen:** Archivo Excel (`.xlsx`)  
**Descripción General:** Base de datos transaccional que registra eventos de siniestralidad vial con granularidad a nivel de incidente. Incluye dimensiones temporales, espaciales, tipología del accidente y métricas de impacto humano (víctimas).

---

## 🧱 ESTRUCTURA Y DICCIONARIO DE DATOS

A continuación, se listan todos los campos presentes en la tabla origen, respetando la nomenclatura original.

| Campo Original | Tipo | Descripción Detallada |
| :--- | :--- | :--- |
| `id` | Entero | Llave primaria. Identificador único de cada reporte de accidente. |
| `Fecha` | Date | Fecha del suceso (DD/MM/AAAA). Base para la dimensión temporal. |
| `hora` | Time | Hora exacta del reporte. Usado para análisis de franjas horarias. |
| `uf` | Texto | Unidad Federativa (Estado/Provincia). Nivel geográfico macro. |
| `municipio` | Texto | Ciudad o localidad del evento. Nivel geográfico micro. |
| `causa_accidente` | Texto | Categoría de la causa principal (ej. Falta de atención, Defecto mecánico). |
| `tipo_accidente` | Texto | Mecánica del impacto (ej. Colisión trasera, Salida de pista). |
| `clasificación_accidente` | Texto | Gravedad cualitativa (Con Víctimas, Sin Víctimas). |
| `sentido_vía` | Texto | Dirección del flujo vehicular (Creciente, Decreciente). |
| `tipo_pista` | Texto | Configuración de la vía (Simple, Doble, Múltiple). |
| `Personas` | Entero | Total de individuos involucrados (Conductores + Pasajeros + Peatones). |
| `Muertos` | Entero | Cantidad de fallecidos en el sitio. |
| `Heridos_graves` | Entero | Personas con lesiones severas que requieren hospitalización. |
| `Herido_Leves` | Entero | Personas con lesiones menores o contusiones. |
| `Ilesos` | Entero | Personas involucradas sin daños físicos aparentes. |
| `Ignorados` | Entero | Personas cuyo estado de salud no pudo ser determinado. |
| `Heridos` | Entero | Sumatoria calculada o registrada de lesionados (Graves + Leves). |
| `vehículos` | Entero | Cantidad de unidades de transporte involucradas. |
| `latitude` | Decimal | Coordenada Y para georreferenciación. |
| `longitude` | Decimal | Coordenada X para georreferenciación. |
| `Año` | Entero | Campo auxiliar (extraído de Fecha). |
| `Mes` | Texto | Campo auxiliar (extraído de Fecha). |
| `Hora Numérica` | Entero | Campo auxiliar (extraído de Hora). |

---

## 🧩 MODELO DE DATOS

Para este análisis, trabajamos sobre un esquema que puede optimizarse a modelo estrella, aunque el dataset actual reside en una tabla principal de hechos.

*   **Tabla de Hechos:** `BD-Accidentes-Carreteras`
*   **Tablas de Dimensión Sugeridas (Virtuales o Físicas):**
    *   `Dim_Fecha` (Calendario)
    *   `Dim_Municipio` (Geografía)
    *   `Dim_TipoAccidente` (Tipología)
    *   `Dim_CausaAccidente` (Causalidad)

---

## 🧮 INGENIERÍA DE DATOS (DAX)

### Columnas Calculadas
Estas columnas se generan fila a fila en la tabla principal para facilitar filtros y agrupaciones.

```dax
Año = YEAR('BD-Accidentes-Carreteras'[Fecha])

Mes = FORMAT('BD-Accidentes-Carreteras'[Fecha], "MMMM")

Hora Numérica = HOUR('BD-Accidentes-Carreteras'[hora])
```

### Tabla Calendario
Es fundamental para la inteligencia de tiempo (comparativas anuales, acumulados, etc.). Se genera mediante DAX.

```dax
Dim_Calendario = 
VAR FechaInicio = MIN('BD-Accidentes-Carreteras'[Fecha]) -- Busca la fecha más antigua
VAR FechaFin = MAX('BD-Accidentes-Carreteras'[Fecha])    -- Busca la fecha más reciente
RETURN
ADDCOLUMNS (
    CALENDAR (FechaInicio, FechaFin),
    "Año", YEAR([Date]),
    "Mes Nro", MONTH([Date]),
    "Mes", FORMAT([Date], "MMMM"),
    "Mes Corto", FORMAT([Date], "MMM"),
    "Día Nro", DAY([Date]),
    "Día Sem Numero", WEEKDAY([Date], 2), -- 1 para Lunes, 7 para Domingo
    "Día Sem Texto", FORMAT([Date], "dddd"),
    "Trimestre Nro", QUARTER([Date]),
    "Trimestre", "Trim " & QUARTER([Date]),
    "Cuatrimestre Nro", ROUNDUP(MONTH([Date]) / 4, 0),
    "Cuatrimestre", "Cuat " & ROUNDUP(MONTH([Date]) / 4, 0),
    "Semana Nro", WEEKNUM([Date], 2),
    "Año Mes", FORMAT([Date], "YYYY-MM"), -- Útil para ordenación cronológica
    "Es Fin de Semana", IF(WEEKDAY([Date], 2) >= 6, "SÍ", "NO"),
    "Trimestre Año", FORMAT([Date], "YYYY") & "-T" & QUARTER([Date])
)
```

---

## 📊 MEDIDAS ANALÍTICAS

A continuación, el compendio completo de medidas DAX creadas.

### Métricas Base

**Total de Accidentes** (Conteo de eventos únicos)
```dax
Total Accidentes = COUNT('BD-Accidentes-Carreteras'[id])
```

**Total de Personas Involucradas**
```dax
Total Personas = SUM('BD-Accidentes-Carreteras'[Personas])
```

**Total de Vehículos**
```dax
Total Vehículos = SUM('BD-Accidentes-Carreteras'[vehiculos])
```

### Métricas de Víctimas

**Total Fallecidos**
```dax
Total Fallecidos = SUM('BD-Accidentes-Carreteras'[Muertos])
```

**Heridos Leves**
```dax
Heridos Leves = SUM('BD-Accidentes-Carreteras'[Heridos_leves])
```

**Heridos Graves**
```dax
Heridos Graves = SUM('BD-Accidentes-Carreteras'[Heridos_graves])
```

**Total Heridos** (Suma de gravedad)
```dax
Total Heridos =
SUM('BD-Accidentes-Carreteras'[Heridos_leves]) +
SUM('BD-Accidentes-Carreteras'[Heridos_graves])
```

### KPIs Estratégicos

**Tasa de Mortalidad (%)** (Ratio de letalidad sobre expuestos)
```dax
Tasa Mortalidad (%) =
DIVIDE([Total Fallecidos], [Total Personas], 0)
```

**Heridos por Accidente** (Intensidad de daños personales)
```dax
Heridos por Accidente =
DIVIDE([Total Heridos], [Total Accidentes], 0)
```

**Promedio de Personas por Accidente**
```dax
Promedio Personas por Accidente =
AVERAGE('BD-Accidentes-Carreteras'[Personas])
```

**Promedio de Vehículos por Accidente**
```dax
Promedio Vehículos por Accidente =
AVERAGE('BD-Accidentes-Carreteras'[vehiculos])
```

### Análisis Temporal

**Accidentes por Año** (Contexto fijo de año)
```dax
Accidentes por Año =
CALCULATE(
    [Total Accidentes],
    ALLEXCEPT('BD-Accidentes-Carreteras', 'BD-Accidentes-Carreteras'[Año])
)
```

**Accidentes por Hora**
```dax
Accidentes por Hora = COUNT('BD-Accidentes-Carreteras'[hora])
```

### Análisis Geográfico

**Accidentes por Municipio**
```dax
Accidentes por Municipio = COUNT('BD-Accidentes-Carreteras'[municipio])
```

**Fallecidos por Municipio**
```dax
Fallecidos por Municipio = SUM('BD-Accidentes-Carreteras'[Muertos])
```

### Causas y Tipología

**Accidentes por Causa**
```dax
Accidentes por Causa = COUNT('BD-Accidentes-Carreteras'[causa_accidente])
```

**Colisiones Traseras** (Filtro específico)
```dax
Colisiones Traseras =
CALCULATE(
    [Total Accidentes],
    'BD-Accidentes-Carreteras'[tipo_accidente] = "Colisión trasera"
)
```

**% Falta de Atención del Conductor** (Peso relativo de la causa principal)
```dax
% Falta Atención =
DIVIDE(
    CALCULATE(
        [Total Accidentes],
        'BD-Accidentes-Carreteras'[causa_accidente] = "Falta de atención de conducción"
    ),
    [Total Accidentes],
    0
)
```

---

## 🎨 VISUALIZACIÓN AVANZADA (HTML & CSS)

Esta sección contiene el código íntegro para los objetos visuales `HTML Content`. Estos scripts generan interfaces web dentro de Power BI.

### 1. Banner Premium (KPIs Principales)
Genera un encabezado responsivo con 4 tarjetas KPI (Accidentes, Fallecidos, Heridos, Tasa) y un título con fecha dinámica.

```dax
HTML_Banner_Premium_Slim = 
VAR Accidentes = FORMAT([Total Accidentes], "#,0")
VAR Fallecidos = FORMAT([Total Fallecidos], "#,0")
VAR Heridos = FORMAT([Total Heridos], "#,0")
VAR TasaMortalidad = FORMAT([Tasa Mortalidad (%)], "0.00%") 
VAR FechaActual = FORMAT(NOW(), "dd/MM/yyyy")

RETURN
"<div style='display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; font-family: Segoe UI, sans-serif; background: white; padding: 1.5% 2%; border-radius: 1vw; box-shadow: 0 6px 18px rgba(0,0,0,0.05); width: 96%; height: auto; min-height: 80px;'>

    <div style='border-left: 0.5vw solid #2c3e50; padding-left: 1.5vw; margin: 10px 0; flex: 1 1 300px;'>
        <h1 style='margin: 0; color: #2c3e50; font-size: clamp(14px, 1.5vw, 18px); font-weight: 900; letter-spacing: -0.5px;'>Análisis de Accidentes de Carreteras</h1>
        <div style='margin-top: 4px; color: #7f8c8d; font-size: clamp(11px, 1.1vw, 13px);'>
            Reporte Estratégico | <b>" & FechaActual & "</b> | <span style='color:#2c3e50;'><b>Ing. Dariel Peña Vasquez</b></span>
        </div>
    </div>

    <div style='display: flex; flex-wrap: wrap; gap: 10px; align-items: center; justify-content: flex-start; flex: 2 1 400px;'>

        <div style='display: flex; align-items: center; background: #ffffff; padding: 10px; border-radius: 10px; border-left: 4px solid #3498db; box-shadow: 0 3px 10px rgba(0,0,0,0.06); flex: 1 1 130px; max-width: 180px;'>
            <span style='font-size: 22px; margin-right: 8px;'>🚗</span>
            <div>
                <div style='font-size: 8px; color: #95a5a6; font-weight: 800; text-transform: uppercase;'>Total Accidentes</div>
                <div style='font-size: clamp(18px, 1.8vw, 22px); font-weight: 900; color: #2c3e50;'>" & Accidentes & "</div>
            </div>
        </div>

        <div style='display: flex; align-items: center; background: #ffffff; padding: 10px; border-radius: 10px; border-left: 4px solid #e74c3c; box-shadow: 0 3px 10px rgba(0,0,0,0.06); flex: 1 1 130px; max-width: 180px;'>
            <span style='font-size: 22px; margin-right: 8px;'>💀</span>
            <div>
                <div style='font-size: 8px; color: #95a5a6; font-weight: 800; text-transform: uppercase;'>Fallecidos</div>
                <div style='font-size: clamp(18px, 1.8vw, 22px); font-weight: 900; color: #e74c3c;'>" & Fallecidos & "</div>
            </div>
        </div>

        <div style='display: flex; align-items: center; background: #ffffff; padding: 10px; border-radius: 10px; border-left: 4px solid #f39c12; box-shadow: 0 3px 10px rgba(0,0,0,0.06); flex: 1 1 130px; max-width: 180px;'>
            <span style='font-size: 22px; margin-right: 8px;'>🚑</span>
            <div>
                <div style='font-size: 8px; color: #95a5a6; font-weight: 800; text-transform: uppercase;'>Heridos</div>
                <div style='font-size: clamp(18px, 1.8vw, 22px); font-weight: 900; color: #2c3e50;'>" & Heridos & "</div>
            </div>
        </div>

        <div style='display: flex; align-items: center; background: #ffffff; padding: 10px; border-radius: 10px; border-left: 4px solid #27ae60; box-shadow: 0 3px 10px rgba(0,0,0,0.06); flex: 1 1 130px; max-width: 180px;'>
            <span style='font-size: 22px; margin-right: 8px;'>📊</span>
            <div>
                <div style='font-size: 8px; color: #95a5a6; font-weight: 800; text-transform: uppercase;'>Tasa Mortalidad</div>
                <div style='font-size: clamp(18px, 1.8vw, 22px); font-weight: 900; color: #2c3e50;'>" & TasaMortalidad & "</div>
            </div>
        </div>

    </div>
</div>"
```

### 2. Grid de Detalle (Vehículos/Personas)
Muestra dos tarjetas grandes con iconos y un pie de página con promedios.

```dax
HTML_Grid_Detalle_Vehiculos = 
"<div style='display: grid; grid-template-columns: 1fr 1fr; gap: 15px; font-family: Segoe UI;'>
    <!-- CARD VEHÍCULOS -->
    <div style='background: white; border: 1px solid #dcdde1; border-radius: 12px; padding: 15px;'>
        <div style='font-size: 25px;'>🚜</div>
        <div style='font-size: 22px; font-weight: 900; color: #2c3e50;'>" & FORMAT([Total Vehículos], "#,0") & "</div>
        <div style='font-size: 10px; font-weight: bold; color: #95a5a6; text-transform: uppercase;'>Unidades Totales</div>
    </div>
    <!-- CARD PERSONAS -->
    <div style='background: white; border: 1px solid #dcdde1; border-radius: 12px; padding: 15px;'>
        <div style='font-size: 25px;'>👨‍👩‍👧‍👦</div>
        <div style='font-size: 22px; font-weight: 900; color: #2c3e50;'>" & FORMAT([Total Personas], "#,0") & "</div>
        <div style='font-size: 10px; font-weight: bold; color: #95a5a6; text-transform: uppercase;'>Pers. Involucradas</div>
    </div>
    <!-- MEDIDAS PROMEDIO -->
    <div style='grid-column: span 2; background: #f8f9fa; padding: 10px; border-radius: 8px; border-left: 4px solid #3498db; display: flex; justify-content: space-between; align-items: center;'>
        <span style='font-size: 12px; font-weight: bold; color: #34495e;'>Promedio por Siniestro:</span>
        <span style='font-size: 14px; color: #3498db; font-weight: 800;'>" & FORMAT([Promedio Vehículos por Accidente], "0.0") & " Veh / " & FORMAT([Promedio Personas por Accidente], "0.0") & " Pers</span>
    </div>
</div>"
```

### 3. Brújula de Riesgo
Identifica automáticamente la causa #1 de accidentes y la muestra con un diseño de alerta.

```dax
HTML_Brujula_Riesgo = 
VAR TopCausa = TOPN(1, SUMMARIZE('BD-Accidentes-Carreteras', 'BD-Accidentes-Carreteras'[causa_accidente]), [Total Accidentes], DESC)
VAR NombreCausa = MINX(TopCausa, 'BD-Accidentes-Carreteras'[causa_accidente])
VAR AccidentesCausa = CALCULATE([Total Accidentes], 'BD-Accidentes-Carreteras'[causa_accidente] = NombreCausa)
VAR VehPorAcc = CALCULATE([Promedio Vehículos por Accidente], 'BD-Accidentes-Carreteras'[causa_accidente] = NombreCausa)

RETURN
"<div style='background: white; border-radius: 18px; padding: 15px; font-family: Segoe UI, sans-serif; border: 1px solid #e1e8ed; overflow: hidden; position: relative;'>
    <div style='position: absolute; right: -10px; top: -10px; font-size: 70px; opacity: 0.04;'>⚠️</div>
    
    <div style='border-left: 4px solid #2c3e50; padding-left: 12px;'>
        <div style='font-size: 10px; color: #95a5a6; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px;'>Causa de Mayor Incidencia</div>
        <div style='font-size: 18px; color: #2c3e50; font-weight: 900; margin: 2px 0;'>" & NombreCausa & "</div>
    </div>

    <div style='display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 12px;'>
        <div style='background: #f8f9fa; padding: 10px; border-radius: 12px; text-align: center; border: 1px solid #f0f2f4;'>
            <div style='font-size: 9px; color: #7f8c8d; text-transform: uppercase; font-weight: 700;'>Volumen Total</div>
            <div style='font-size: 20px; font-weight: 900; color: #3498db; line-height: 1.1;'>" & FORMAT(AccidentesCausa, "#,0") & "</div>
            <div style='font-size: 9px; color: #3498db; font-weight: bold;'>SINIESTROS</div>
        </div>
        <div style='background: #f8f9fa; padding: 10px; border-radius: 12px; text-align: center; border: 1px solid #f0f2f4;'>
            <div style='font-size: 9px; color: #7f8c8d; text-transform: uppercase; font-weight: 700;'>Complejidad</div>
            <div style='font-size: 20px; font-weight: 900; color: #2c3e50; line-height: 1.1;'>" & FORMAT(VehPorAcc, "0.0") & "</div>
            <div style='font-size: 9px; color: #2c3e50; font-weight: bold;'>VEH/EVENTO</div>
        </div>
    </div>

    <div style='margin-top: 12px; background: #ebf5fb; padding: 8px 10px; border-radius: 10px; font-size: 11px; color: #2c3e50; border: 1px dashed #3498db; line-height: 1.2;'>
        🎯 <b>Insight:</b> Reforzar <b>patrullaje preventivo (+20%)</b> en puntos críticos detectados.
    </div>
</div>"
```

### 4. Heatmap Semanal
Un monitor visual que muestra la intensidad de accidentes por día de la semana (Lunes a Domingo), usando colores y emojis dinámicos.

```dax
HTML_Heatmap_Dias_V3 = 
-- Creamos una tabla virtual con los números del 1 al 7 y les asignamos un nombre claro
VAR DiasSerie = SELECTCOLUMNS(GENERATESERIES(1, 7), "ID_Dia", [Value])

-- Buscamos el valor máximo de accidentes entre los 7 días para escalar los colores
VAR MaxAcc = 
    MAXX(
        DiasSerie,
        VAR _DiaActual = [ID_Dia]
        RETURN CALCULATE([Total Accidentes], KEEPFILTERS(WEEKDAY('BD-Accidentes-Carreteras'[Fecha], 2) = _DiaActual))
    )

VAR Bloques = 
    CONCATENATEX(
        DiasSerie,
        VAR DiaNum = [ID_Dia]
        -- Calculamos accidentes para el día específico (1=Lunes, 7=Domingo)
        VAR AccDia = CALCULATE([Total Accidentes], KEEPFILTERS(WEEKDAY('BD-Accidentes-Carreteras'[Fecha], 2) = DiaNum))
        
        -- Evitamos división por cero y calculamos opacidad
        VAR Opacidad = IF(MaxAcc > 0, DIVIDE(AccDia, MaxAcc), 0)
        
        -- Formato de miles (ej: 1,250)
        VAR ValorFormateado = FORMAT(AccDia, "#,0")
        
        -- Lógica de Emojis según intensidad
        VAR EmojiNivel = 
            SWITCH(TRUE(),
                Opacidad > 0.8, "🚨", -- Riesgo Crítico
                Opacidad > 0.4, "⚠️", -- Riesgo Moderado
                AccDia = 0, "✅",     -- Sin accidentes
                "🟡"                  -- Riesgo Bajo
            )
            
        RETURN
        "<div style='flex: 1; 
                    background: rgba(231, 76, 60, " & SUBSTITUTE(MAX(0.05, Opacidad), ",", ".") & "); 
                    height: 85px; 
                    border-radius: 12px; 
                    margin: 4px; 
                    display: flex; 
                    flex-direction: column;
                    align-items: center; 
                    justify-content: center; 
                    color: " & IF(Opacidad > 0.6, "white", "#2c3e50") & "; 
                    box-shadow: 0 4px 6px rgba(0,0,0,0.05);
                    border: 1px solid rgba(0,0,0,0.1);'>
            <span style='font-size: 20px; margin-bottom: 5px;'>" & EmojiNivel & "</span>
            <span style='font-size: 14px; font-weight: 800;'>" & ValorFormateado & "</span>
        </div>",
        ""
    )

RETURN
"<div style='background: #ffffff; padding: 25px; border-radius: 20px; font-family: Segoe UI, sans-serif; border: 1px solid #dcdde1; box-shadow: 0 10px 25px rgba(0,0,0,0.05);'>
    <div style='display: flex; align-items: center; margin-bottom: 20px;'>
        <div style='background: #2c3e50; width: 45px; height: 45px; border-radius: 12px; display: flex; align-items: center; justify-content: center; margin-right: 15px;'>
            <span style='font-size: 24px; color: white;'>📅</span>
        </div>
        <div>
            <h3 style='margin: 0; color: #2c3e50; font-size: 18px;'>Monitor de Intensidad Semanal</h3>
            <p style='margin: 0; font-size: 11px; color: #7f8c8d;'>Siniestros acumulados por día de la semana</p>
        </div>
    </div>

    <div style='display: flex; width: 100%;'>
        " & Bloques & "
    </div>

    <div style='display: flex; justify-content: space-between; font-size: 11px; color: #2c3e50; margin-top: 12px; font-weight: 800; text-transform: uppercase;'>
        <span style='width: 14%; text-align: center;'>Lun</span>
        <span style='width: 14%; text-align: center;'>Mar</span>
        <span style='width: 14%; text-align: center;'>Mié</span>
        <span style='width: 14%; text-align: center;'>Jue</span>
        <span style='width: 14%; text-align: center;'>Vie</span>
        <span style='width: 14%; text-align: center;'>Sáb</span>
        <span style='width: 14%; text-align: center;'>Dom</span>
    </div>

    <div style='margin-top: 25px; padding-top: 15px; border-top: 1px dotted #dcdde1; display: flex; justify-content: space-between; align-items: center;'>
        <div style='font-size: 11px; color: #7f8c8d;'>
             Pico Máximo: <strong>" & FORMAT(MaxAcc, "#,0") & "</strong>
        </div>
        <div style='font-size: 11px; color: #e74c3c; font-weight: bold;'>
             Análisis de Riesgo Temporal
        </div>
    </div>
</div>"
```

### 5. Ranking Top 10 Causas
Tabla HTML enriquecida con barras de progreso internas y alertas de mortalidad.

```dax
HTML_Tabla_Top10_Causas = 
VAR TablaTop10 = 
    TOPN(10, 
        SUMMARIZE('BD-Accidentes-Carreteras', 'BD-Accidentes-Carreteras'[causa_accidente]),
        [Total Accidentes], DESC
    )
VAR MaxAcc = MAXX(TablaTop10, [Total Accidentes])
VAR MaxFal = MAXX(TablaTop10, [Total Fallecidos])

VAR Filas = 
    CONCATENATEX(
        TablaTop10,
        VAR Acc = [Total Accidentes]
        VAR Fal = [Total Fallecidos]
        VAR Her = [Total Heridos]
        VAR Veh = [Total Vehículos]
        VAR Per = [Total Personas]
        
        VAR PctAcc = DIVIDE(Acc, MaxAcc, 0) * 100
        VAR ColorFal = IF(Fal > (MaxFal * 0.7), "#e74c3c", "#f39c12")
        
        RETURN
        "<tr>
            <td style='padding: 6px 10px; font-weight: bold; color: #2c3e50; font-size: 11px; border-bottom: 1px solid #f1f2f6;'>" & 'BD-Accidentes-Carreteras'[causa_accidente] & "</td>
            <td style='padding: 6px 10px; border-bottom: 1px solid #f1f2f6; text-align: center; color: #7f8c8d; font-size: 11px;'>" & FORMAT(Per, "#,0") & "</td>
            <td style='padding: 6px 10px; border-bottom: 1px solid #f1f2f6;'>
                <div style='display: flex; align-items: center; justify-content: space-between;'>
                    <span style='font-weight: 800; color: #3498db; font-size: 11px;'>" & FORMAT(Acc, "#,0") & "</span>
                    <div style='width: 40%; background: #ebf5fb; height: 4px; border-radius: 2px; margin-left: 8px;'>
                        <div style='width: " & SUBSTITUTE(PctAcc, ",", ".") & "%; background: #3498db; height: 100%; border-radius: 2px;'></div>
                    </div>
                </div>
            </td>
            <td style='padding: 6px 10px; border-bottom: 1px solid #f1f2f6; text-align: center; font-weight: bold; color: " & ColorFal & "; font-size: 11px;'>
                " & IF(Fal > 100, "💀 ", "") & FORMAT(Fal, "#,0") & "
            </td>
            <td style='padding: 6px 10px; border-bottom: 1px solid #f1f2f6; text-align: center; color: #2c3e50; font-size: 11px;'>" & FORMAT(Her, "#,0") & "</td>
            <td style='padding: 6px 10px; border-bottom: 1px solid #f1f2f6; text-align: center; color: #95a5a6; font-size: 11px;'>" & FORMAT(Veh, "#,0") & "</td>
        </tr>",
        "",
        [Total Accidentes], DESC -- Esto garantiza el orden de mayor a menor en el HTML
    )

RETURN
"<div style='margin: 0; padding: 0; width: 100%; overflow-x: auto; font-family: Segoe UI, sans-serif;'>
    <div style='background: white; border-radius: 12px; border: 1px solid #e1e8ed; box-shadow: 0 8px 20px rgba(0,0,0,0.05); overflow: hidden;'>
        <div style='padding: 10px 15px; background: #ffffff; border-bottom: 1px solid #f1f2f6; display: flex; justify-content: space-between; align-items: center;'>
            <h2 style='margin: 0; color: #2c3e50; font-size: 16px;'>🚩 Ranking Crítico: Top 10 Causas</h2>
            <span style='background: #fdf2f2; color: #e74c3c; padding: 3px 10px; border-radius: 15px; font-size: 10px; font-weight: bold;'>SCORE DESCENDENTE</span>
        </div>
        <table style='width: 100%; border-collapse: collapse; table-layout: auto;'>
            <thead>
                <tr style='background: #f8f9fa; text-align: left; font-size: 10px; color: #95a5a6; text-transform: uppercase; letter-spacing: 0.5px;'>
                    <th style='padding: 8px 10px;'>Causa</th>
                    <th style='padding: 8px 10px; text-align: center;'>Personas</th>
                    <th style='padding: 8px 10px;'>Siniestros</th>
                    <th style='padding: 8px 10px; text-align: center;'>Muertes</th>
                    <th style='padding: 8px 10px; text-align: center;'>Heridos</th>
                    <th style='padding: 8px 10px; text-align: center;'>Vehículos</th>
                </tr>
            </thead>
            <tbody>
                " & Filas & "
            </tbody>
        </tbody>
    </table>
</div>"
```

### 6. Bitácora de Siniestros Críticos
Lista los Top 5 accidentes más graves, aplicando un sistema de puntuación ponderada (Muertos*100 + Graves*10 + Leves).

```dax
HTML_Table_Siniestros_Impacto = 
VAR TablaCritica = 
    TOPN(5, 
        'BD-Accidentes-Carreteras', 
        -- Ordenamos por una puntuación de gravedad: Muertos pesan más que heridos
        ([Muertos] * 100) + ([Heridos_graves] * 10) + [Herido_Leves], DESC
    )

VAR Filas = 
    CONCATENATEX(
        TablaCritica,
        VAR vMuertos = [Muertos]
        VAR vHeridos = [Heridos_graves] + [Herido_Leves]
        VAR vCausa = [causa_accidente]
        VAR vFecha = FORMAT([Fecha], "dd/MM/yy")
        
        -- Definición de Color y Etiqueta de Estado
        VAR ColorEstado = SWITCH(TRUE(),
            vMuertos > 0, "#e74c3c",      -- Rojo para fatalidades
            vHeridos > 0, "#f39c12",      -- Naranja para heridos
            "#3498db"                     -- Azul para solo daños
        )
        VAR TextoEstado = SWITCH(TRUE(),
            vMuertos > 0, "FATAL",
            vHeridos > 0, "GRAVE",
            "MODERADO"
        )
        
        RETURN
        "<tr>
            <td style='padding: 10px; border-bottom: 1px solid #f1f2f6;'>
                <div style='font-size: 11px; font-weight: bold; color: #2c3e50;'>" & [municipio] & "</div>
                <div style='font-size: 10px; color: #95a5a6;'>" & vFecha & "</div>
            </td>
            <td style='padding: 10px; border-bottom: 1px solid #f1f2f6; font-size: 11px; color: #34495e;'>" & vCausa & "</td>
            <td style='padding: 10px; border-bottom: 1px solid #f1f2f6; text-align: center;'>
                <div style='font-size: 11px; font-weight: 800; color: #2c3e50;'>" & vMuertos & " 💀</div>
                <div style='font-size: 10px; color: #7f8c8d;'>" & vHeridos & " 🚑</div>
            </td>
            <td style='padding: 10px; border-bottom: 1px solid #f1f2f6; text-align: right;'>
                <span style='background: " & ColorEstado & "; color: white; padding: 4px 10px; border-radius: 6px; font-size: 9px; font-weight: 900; letter-spacing: 0.5px;'>
                    " & TextoEstado & "
                </span>
            </td>
        </tr>",
        "",
        ([Muertos] * 100) + ([Heridos_graves] * 10) + [Herido_Leves], DESC
    )

RETURN
"<div style='background: white; padding: 0; border-radius: 15px; font-family: Segoe UI, sans-serif; box-shadow: 0 10px 25px rgba(0,0,0,0.05); border: 1px solid #e1e8ed; overflow: hidden;'>
    <!-- Encabezado de la Tabla -->
    <div style='background: #2c3e50; padding: 15px; display: flex; justify-content: space-between; align-items: center;'>
        <h3 style='margin: 0; color: white; font-size: 14px;'>🚨 BITÁCORA DE SINIESTROS CRÍTICOS</h3>
        <span style='font-size: 10px; color: #bdc3c7;'>TOP 5 GRAVEDAD</span>
    </div>

    <table style='width: 100%; border-collapse: collapse; background: white;'>
        <thead>
            <tr style='text-align: left; background-color: #f8f9fa; color: #95a5a6; font-size: 10px; text-transform: uppercase;'>
                <th style='padding: 10px; border-bottom: 2px solid #f1f2f6;'>Ubicación</th>
                <th style='padding: 10px; border-bottom: 2px solid #f1f2f6;'>Causa Principal</th>
                <th style='padding: 10px; border-bottom: 2px solid #f1f2f6; text-align: center;'>Impacto</th>
                <th style='padding: 10px; border-bottom: 2px solid #f1f2f6; text-align: right;'>Estado</th>
            </tr>
        </thead>
        <tbody>
            " & Filas & "
        </tbody>
    </table>
    
    <div style='padding: 10px; background: #fdfdfd; text-align: center; font-size: 10px; color: #bdc3c7; font-style: italic;'>
        Los datos se filtran automáticamente por nivel de letalidad detectado.
    </div>
</div>"
```

### 7. Radar de Víctimas
Barra de progreso segmentada que muestra la distribución porcentual de fallecidos, heridos e ilesos.

```dax
HTML_Radar_Victimas = 
VAR Total = [Total Personas]
VAR PctFallecidos = DIVIDE(SUM('BD-Accidentes-Carreteras'[Muertos]), Total, 0) * 100
VAR PctHeridos = DIVIDE([Total Heridos], Total, 0) * 100
VAR PctIlesos = DIVIDE(SUM('BD-Accidentes-Carreteras'[ilesos]), Total, 0) * 100

RETURN
"<div style='background: #2c3e50; color: white; padding: 20px; border-radius: 15px; font-family: Segoe UI;'>
    <h3 style='margin: 0 0 15px 0; font-size: 15px; color: #f1c40f;'>📊 Perfil de Involucrados</h3>
    
    <div style='margin-bottom: 10px;'>
        <div style='display: flex; justify-content: space-between; font-size: 11px;'>
            <span>💀 Fallecidos</span><span>" & FORMAT(PctFallecidos, "0.0") & "%</span>
        </div>
        <div style='background: #34495e; height: 6px; border-radius: 3px;'>
            <div style='background: #e74c3c; width: " & SUBSTITUTE(PctFallecidos, ",", ".") & "%; height: 100%; border-radius: 3px;'></div>
        </div>
    </div>

    <div style='margin-bottom: 10px;'>
        <div style='display: flex; justify-content: space-between; font-size: 11px;'>
            <span>🚑 Heridos</span><span>" & FORMAT(PctHeridos, "0.0") & "%</span>
        </div>
        <div style='background: #34495e; height: 6px; border-radius: 3px;'>
            <div style='background: #f39c12; width: " & SUBSTITUTE(PctHeridos, ",", ".") & "%; height: 100%; border-radius: 3px;'></div>
        </div>
    </div>

    <div style='margin-bottom: 10px;'>
        <div style='display: flex; justify-content: space-between; font-size: 11px;'>
            <span>✅ Ilesos</span><span>" & FORMAT(PctIlesos, "0.0") & "%</span>
        </div>
        <div style='background: #34495e; height: 6px; border-radius: 3px;'>
            <div style='background: #27ae60; width: " & SUBSTITUTE(PctIlesos, ",", ".") & "%; height: 100%; border-radius: 3px;'></div>
        </div>
    </div>
</div>"
```

### 8. Embudo de Letalidad
Representación visual del embudo de conversión: Personas Expuestas -> Heridos -> Fallecidos.

```dax
HTML_Visual_Embudo_Letalidad = 
VAR Involucrados = [Total Personas]
VAR Heridos = [Total Heridos]
VAR Fallecidos = [Total Fallecidos]

-- Eliminamos el * 100 aquí para que el FORMAT con % funcione correctamente
VAR PctHeridos = DIVIDE(Heridos, Involucrados, 0)
VAR PctFallecidos = DIVIDE(Fallecidos, Involucrados, 0)

-- Variable para el texto final (multiplicamos solo para el número del texto)
VAR FallecidosPorCada100 = PctFallecidos * 100

RETURN
"<div style='background: #2c3e50; color: white; padding: 12px 15px; border-radius: 20px; font-family: Segoe UI, sans-serif; text-align: center; box-shadow: 0 10px 20px rgba(0,0,0,0.2); line-height: 1.2;'>
    
    <h3 style='margin: 5px 0 12px 0; font-size: 14px; letter-spacing: 1px; color: #bdc3c7; text-transform: uppercase;'>Embudo de Conversión Crítica</h3>
    
    <!-- Nivel 1: Total Involucrados -->
    <div style='background: rgba(255,255,255,0.1); padding: 10px; border-radius: 12px 12px 5px 5px; margin-bottom: 3px; border: 1px solid rgba(255,255,255,0.05);'>
        <div style='font-size: 9px; opacity: 0.7; text-transform: uppercase;'>Total Personas Expuestas</div>
        <div style='font-size: 22px; font-weight: 900;'>" & FORMAT(Involucrados, "#,0") & "</div>
    </div>

    <!-- Nivel 2: Heridos -->
    <div style='background: #f39c12; padding: 8px; border-radius: 5px; margin: 0 auto 3px auto; width: 85%; border: 1px solid rgba(0,0,0,0.1);'>
        <div style='font-size: 9px; font-weight: bold; color: #2c3e50;'>🚑 HERIDOS (" & FORMAT(PctHeridos, "0.0%") & ")</div>
        <div style='font-size: 18px; font-weight: 800; color: white;'>" & FORMAT(Heridos, "#,0") & "</div>
    </div>

    <!-- Nivel 3: Fallecidos -->
    <div style='background: #e74c3c; padding: 8px; border-radius: 5px 5px 15px 15px; margin: 0 auto; width: 65%; box-shadow: 0 5px 15px rgba(231, 76, 60, 0.3);'>
        <div style='font-size: 9px; font-weight: bold;'>💀 FALLECIDOS (" & FORMAT(PctFallecidos, "0.0%") & ")</div>
        <div style='font-size: 18px; font-weight: 800;'>" & FORMAT(Fallecidos, "#,0") & "</div>
    </div>

    <!-- Nota inferior compacta -->
    <div style='margin-top: 12px; font-size: 10px; color: #95a5a6; font-style: italic; border-top: 1px solid rgba(255,255,255,0.1); padding-top: 8px;'>
        *De cada 100 personas involucradas, <b>" & FORMAT(FallecidosPorCada100, "0.0") & "</b> pierden la vida.
    </div>
</div>"
```

### 9. Burbujas de Gravedad
Visualización de burbujas donde el tamaño depende del volumen de accidentes y el color de la existencia de fallecidos.

```dax
HTML_Severity_Bubbles_Fixed = 
VAR TablaTop = 
    TOPN(3, 
        SUMMARIZE('BD-Accidentes-Carreteras', 'BD-Accidentes-Carreteras'[causa_accidente]),
        [Total Accidentes], DESC
    )

VAR Burbujas = 
    CONCATENATEX(
        TablaTop,
        VAR ValorAcc = [Total Accidentes]
        VAR MedidaTamano = (DIVIDE(ValorAcc, MAXX(TablaTop, [Total Accidentes]), 0) * 40) + 20 -- Tamano relativo entre 20 y 60px
        VAR ColorBurbuja = IF([Total Fallecidos] > 0, "#e74c3c", "#3498db")
        RETURN
        "<div style='display: flex; align-items: center; margin-bottom: 15px;'>
            <div style='width: " & SUBSTITUTE(MedidaTamano, ",", ".") & "px; 
                        height: " & SUBSTITUTE(MedidaTamano, ",", ".") & "px; 
                        background: " & ColorBurbuja & "; 
                        border-radius: 50%; margin-right: 15px; opacity: 0.8; 
                        display: flex; align-items: center; justify-content: center; color: white; font-size: 10px; font-weight: bold;'>
                " & ValorAcc & "
            </div>
            <div>
                <div style='font-size: 13px; font-weight: bold; color: #2c3e50;'>" & 'BD-Accidentes-Carreteras'[causa_accidente] & "</div>
                <div style='font-size: 11px; color: #7f8c8d;'>" & [Total Fallecidos] & " Fallecidos detectados</div>
            </div>
        </div>",
        ""
    )

RETURN
"<div style='background: white; padding: 20px; border-radius: 15px; font-family: Segoe UI; border: 1px solid #dcdde1; box-shadow: 0 4px 6px rgba(0,0,0,0.05);'>
    <h3 style='margin: 0 0 20px 0; color: #2c3e50; font-size: 16px; border-left: 4px solid #e74c3c; padding-left: 10px;'>🔴 Gravedad por Causa</h3>
    <div style='display: flex; flex-direction: column;'>
        " & Burbujas & "
    </div>
</div>"
```

---

## 🧠 INSIGHTS Y CONCLUSIONES

Basado en el procesamiento de datos realizado:

*   **Principal Causa:** La *falta de atención del conductor* domina la siniestralidad, seguida por fallas mecánicas.
*   **Volumen vs. Gravedad:** Aunque las *colisiones traseras* son las más frecuentes, no necesariamente son las más letales.
*   **Geografía:** Se observa alta concentración en zonas urbanas, lo que sugiere que la densidad de tráfico es un factor de riesgo mayor que la velocidad pura de autopista.
*   **Temporalidad:** Las horas pico laborales muestran picos claros de incidentes.
*   **Ratio de Supervivencia:** Existe una alta tasa de heridos frente a una baja tasa de mortalidad, lo cual indica que la mayoría de los accidentes son de intensidad media-baja.

**Acciones Recomendadas:**
1.  Implementar campañas de concientización sobre distracciones al volante.
2.  Utilizar los mapas de calor para asignar recursos de tránsito en horarios críticos.
3.  Reforzar la vigilancia en los municipios con mayor tasa de incidencia.

---

**Autor:** Ing. Dariel Peña Vasquez  
**Herramienta:** Microsoft Power BI  
**Formato:** Documentación Técnica Extendida (README)
