# 🏨 Rainforest Revenue Intelligence Dashboard

[![Microsoft Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=flat&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Power Query](https://img.shields.io/badge/Power%20Query-0078D4?style=flat&logo=microsoft&logoColor=white)](https://learn.microsoft.com/en-us/power-query/)
[![DAX](https://img.shields.io/badge/DAX-F2C811?style=flat&logo=powerbi&logoColor=black)](https://learn.microsoft.com/en-us/dax/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/Version-1.0-blue.svg)]()

---

## 📋 Tabla de Contenidos

- [Descripción General](#-descripción-general)
- [Características Principales](#-características-principales)
- [Contenido del Proyecto](#-contenido-del-proyecto)
- [Herramientas y Tecnologías](#️-herramientas-y-tecnologías-utilizadas)
- [Configuración del Entorno](#️-configuración-del-entorno)
- [Estructura del Repositorio](#-estructura-del-repositorio)
- [Medidas DAX Implementadas](#-medidas-dax-implementadas)
- [Vistas Previas](#️-vistas-previas-del-proyecto)
- [Autor](#-autor)
- [Licencia](#-licencia)

---

## 📃 Descripción General

Dashboard de **Revenue Management para Hotelería** desarrollado en **Microsoft Excel**, diseñado para analizar y optimizar el rendimiento de ingresos de una cadena hotelera nacional en la selva. Este proyecto transforma datos de reservaciones en métricas estratégicas del sector hotelero mediante:

- 📊 **2 Vistas Analíticas:** Overview (vista general) y Deep Dive (Actual vs Budget vs LY).
- 🏨 **KPIs Hoteleros:** +10 medidas DAX especializadas (RevPAR, ADR, Occupancy %, RNS).
- 🎯 **Análisis de Ingresos:** Comparativa detallada de Revenue real, presupuestado y año anterior (Last Year).
- 🖱️ **Navegación Interactiva:** Segmentación dinámica por sucursal y período de tiempo.

---

## ✨ Características Principales

### 🔄 Transformaciones en Power Query
- Procesos de limpieza y modelado de datos para optimizar el rendimiento.
- Creación de tabla calendario personalizada para análisis temporal y comparativas YoY (Year over Year).
- Unificación de fuentes de datos en un modelo relacional coherente dentro de Excel.

### 📐 Modelo de Datos
- **Tablas de Dimensión:** `DimCalendar` (Calendario).
- **Tablas de Hechos:** `FactAvailability` (Disponibilidad de habitaciones) y `FactTransaction` (Ingresos, RNS y Metas).
- Relaciones optimizadas para análisis multidimensional utilizando Power Pivot.

### 🎨 Diseño Interactivo
- Navegación por hojas con paginado intuitivo (Home, Overview, Deep Dive).
- Segmentación de datos (Slicers) dinámica por sucursal y fechas.
- Diseño limpio y profesional enfocado en la toma de decisiones gerenciales.

---

## 📊 Contenido del Proyecto

| Hoja | Descripción | Contenido Principal |
|--------|-------------|---------------------|
| **Home** | Presentación | Portada del proyecto y descripción general. |
| **Overview** | Vista general ejecutiva | KPIs hoteleros globales (RevPAR, ADR, Occupancy) y tendencias principales. |
| **Deep Dive** | Análisis profundo | Comparativa detallada de Ingresos (Revenue), Metas (Budget) y Año Anterior (LY) por período. |
| **Settings** | Configuración | Datos y parámetros que controlan la interactividad y segmentación del Dashboard. |

---

## 🛠️ Herramientas y Tecnologías Utilizadas

### Software Requerido
- **Microsoft Excel** (versión 2016 o superior, recomendado Microsoft 365 para mejores funciones de Power Pivot y Power Query).

### Fuente de Datos
- 📄 Los datos fuente están integrados y estructurados directamente dentro del archivo `.xlsx` en las hojas de hechos (`FactAvailability`, `FactTransaction`).

### Lenguajes y Tecnologías
- **DAX** - Cálculo de medidas y KPIs hoteleros avanzados (+15 fórmulas).
- **Power Query (Lenguaje M)** - Transformación, limpieza y carga de datos.
- **Power Pivot** - Modelado de datos y creación de relaciones dentro de Excel.

---

## ⚙️ Configuración del Entorno

### Prerequisitos
1. Tener instalado **Microsoft Excel** (versión de escritorio).
2. Asegurarse de que los complementos **Power Query** y **Power Pivot** estén habilitados en Excel (Archivo > Opciones > Complementos).

### Paso 1: Obtener el Proyecto
1. Clonar o descargar el repositorio completo.
2. Ubicar el archivo `Rainforest Hotel Dashboard.xlsx`.

### Paso 2: Cargar y Actualizar Datos
1. Abrir el archivo en Microsoft Excel.
2. Navegar a la pestaña **Datos** en la cinta de opciones.
3. Hacer clic en el botón **Actualizar todo** para refrescar las consultas de Power Query y el modelo de datos Power Pivot.
4. Si se desea modificar los datos base, se pueden editar las tabulas directamente o ajustar las consultas en el Editor de Power Query.

---

## 📂 Estructura del Repositorio

```text
.
├── Rainforest Hotel Dashboard.xlsx  # Archivo principal del proyecto con el dashboard, modelo de datos y hojas de cálculo
├── .gitignore                       # Archivo de configuración Git
└── README.md                        # Este archivo de documentación
```

## 🧮 Medidas DAX Implementadas
El proyecto cuenta con medidas DAX especializadas para el análisis de revenue management hotelero. A continuación, se detallan las principales:
#### 💰 Ingresos y Noches (Revenue & RNS)
<details>
<summary>Ver medidas de Ingresos y Noches</summary>

#Actual Revenue 

```dax
= CALCULATE(SUM([Revenue]), FactTransaction[Base] = "Actual")
```

#LY Rev 

```dax
= 
SUMX(
    VALUES('DimCalendar'[Date]),
    CALCULATE([#Actual Revenue], DATEADD('DimCalendar'[Date], -1, YEAR))
)
```

#Budget Revenue 

```dax
= CALCULATE(SUM(FactTransaction[Revenue]), FactTransaction[Base] = "Budget")
```


#Actual RNS 

```dax
= CALCULATE(SUM(FactTransaction[RNS]), FactTransaction[Base] = "Actual")
```

#LY RNS 

```dax
= 
SUMX(
    VALUES('DimCalendar'[Date]),
    CALCULATE([#Actual RNS], DATEADD('DimCalendar'[Date], -1, YEAR))
)
```

#Budget RNS 

```dax
= CALCULATE(SUM(FactTransaction[RNS]), FactTransaction[Base] = "Budget")
```

</details>

#### 🏨 KPIs Hoteleros (ADR, Occupancy, RevPAR)
<details>
<summary>Ver medidas de KPIs Hoteleros</summary>

#Actual ADR

```dax
= DIVIDE([#Actual Revenue], [#Actual RNS], 0)
```

#LY ADR 

```dax
= DIVIDE([#LY Rev], [#LY RNS], 0)
```

#Budget ADR 

```dax
= DIVIDE([#Budget Revenue], [#Budget RNS], 0)
```

#Actual Availability 

```dax
= SUM(FactAvailability[Availability])
```

#LY Availability 

```dax
= SUMX(
    VALUES('DimCalendar'[Date]),
    CALCULATE([#Actual Availability], DATEADD('DimCalendar'[Date], -1, YEAR))
)
```

#Actual Occupancy(%) 

```dax
= DIVIDE([#Actual RNS], [#Actual Availability], 0)
```

#LY Occupancy(%) 

```dax
= DIVIDE([#LY RNS], [#LY Availability], 0)
```

#Budget Occupancy(%) 

```dax
= DIVIDE([#Budget RNS], [#Actual Availability], 0)
```

#Actual RevPAR 

```dax
= [#Actual ADR] * [#Actual Occupancy(%)]
```

#LY RevPAR 

```dax
= [#LY ADR] * [#LY Occupancy(%)]
```


#Budget RevPAR 

```dax
= [#Budget ADR] * [#Budget Occupancy(%)]
```

</details>

#### 📈 Acumulados Anuales (YTD)
<details>
<summary>Ver medidas YTD</summary>

#Budget Revenue YTD 

```dax
= CALCULATE([#Budget Revenue], DATESYTD(DimCalendar[Date]))
```

#Actual Revenue YTD 

```dax
= CALCULATE([#Actual Revenue], DATESYTD(DimCalendar[Date]))
```

</details>

## 🖼️ Vistas Previas del Proyecto
<details>
<summary>📸 Ver Capturas de Pantalla</summary>

<img width="2101"  alt="Dashboard View 1" src="https://github.com/user-attachments/assets/d43168e3-cb4e-4ade-ae5b-6113861f1310" />
<img width="1512"  alt="Dashboard View 2" src="https://github.com/user-attachments/assets/7160fe37-7260-4b29-955d-6bf7435357b7" />
<img width="1824"  alt="Dashboard View 3" src="https://github.com/user-attachments/assets/e70a39e8-cdfe-4629-a9b8-7d966ca106de" />

</details>

<details>
<summary>🎥 Ver Video Demostrativo</summary>

   
   Ver en YouTube: https://youtu.be/3ah8CH0h_vA
</details>

## 👤 Autor
Giancarlo Barrantes

Business Intelligence Developer | Power BI Specialist


<a href="https://www.linkedin.com/in/gb25"><img width="70" height="20" alt="image" src="https://github.com/user-attachments/assets/e0d69586-c6f9-4f66-9a3d-bc8dfa146463" /></a>
<a href="https://www.instagram.com/gbarrantes25"><img width="70" height="20" alt="image" src="https://github.com/user-attachments/assets/68773029-5e0f-49da-92e9-fb93a223dd09" />
</a>
<a href="https://www.tiktok.com/@gbarrantes25"><img width="65" height="20" alt="image" src="https://github.com/user-attachments/assets/349c87c0-1f3a-4368-8517-64a6012565c0" /></a>
<a href="https://www.youtube.com/@gbarrantes25"><img width="70" height="20" alt="image" src="https://github.com/user-attachments/assets/d9e2cb69-29ec-47fc-93ae-30896a7b533d" />
</a>

## 📝 Licencia
Este proyecto está bajo la Licencia MIT - ver el archivo <a href="https://mit-license.org/">LICENSE</a> para más detalles.

## ⭐ Agradecimientos
Si este proyecto te resultó útil, considera darle una ⭐ en GitHub y compartirlo con otros profesionales de Business Intelligence.

<div align="center">
   Hecho por <a href="https://gbarrantes25.github.io/">Giancarlo Barrantes</a>
</div>
- [tiktok.com/@gbarrantes25](https://www.tiktok.com/@gbarrantes25)
- [youtube.com/@gbarrantes25](https://www.youtube.com/@gbarrantes25)
