# Rainforest Revenue Intelligence Dashboard

## 📃 Descripción General
Dashboard de Revenue Management para Hotelería desarrollado en Excel, diseñado para analizar y optimizar el rendimiento de ingresos de una cadena hotelera nacional en la selva. Este proyecto transforma datos de reservaciones en métricas estratégicas del sector hotelero mediante:


- 📊 2 Vistas Analíticas: Overview (vista general) y Deep Dive (Actual vs Budget vs LY).
- 🏨 KPIs Hoteleros: +10 medidas DAX especializadas (RevPAR, ADR, Occupancy %).
- 🎯 Análisis de Ingresos: Revenue y Metas.
- 🖱️ Navegación Interactiva: Segmentación por sucursal y período de tiempo.


## 📊 Contenido del proyecto
- Hoja "Home": Contiene la presentación del proyecto.
- Hoja "FactAvailability": Contiene la tabla de disponibilidad de habitaciones.
- Hoja "FactTransaction": Contiene la tabla de ingresos y metas.
- Hoja "Overview": Contiene la vista general del dashboard.
- Hoja "Deep Dive": Contiene la comparación de ingresos, metas y LY.
- Hoja "Settings": Contiene los datos de la interactividad del Dashboard.


## 🛠️ Herramientas y Tecnologías Utilizadas
- Visualización: Microsoft Excel.
- Fuente de Datos: Está incluido dentro del archivo .xlsx
- Lenguajes: DAX para las medidas calculadas y Power Query (Lenguaje M) para la transformación de datos.


## ⚙️ Configuración del Entorno
- Software Necesario: Microsoft Excel.
- Instalación:
  - Descargar [Rainforest Hotel Dashboard.xlsx](https://github.com/Gbarrantes25/RainforestRevenueIntelligence-Dashboard-Excel/blob/main/Rainforest%20Hotel%20Dashboard.xlsx) con Microsoft Excel.
  - Entrar a Inicio y darle click a "Actualizar".


## 📂 Estructura del Repositorio
<code>.
  ├── Rainforest Hotel Dashboard.xlsx  # Contiene el archivo del proyecto en formato .xlsx            
  └── README.md                        # Este archivo.
</code>


## ✅ Características Principales
- Transformaciones en Power Query: Se realizaron procesos de limpieza y modelado de datos para optimizar el rendimiento.
- Creación de tabla calendario.
- Medidas DAX:
  <details>
  <summary>Click para expandir medidas</summary>

    
  - **#Actual Revenue:** `=CALCULATE(SUM([Revenue]),FactTransaction[Base]="Actual")`
  - **#LY Rev:**
    ```dax
    =SUMX(
    VALUES('DimCalendar'[Date]),
    CALCULATE(
        [#Actual Revenue],
        DATEADD('DimCalendar'[Date], -1, YEAR)
    ))
    ```
  - **#Budget Revenue:** `=CALCULATE(SUM(FactTransaction[Revenue]),FactTransaction[Base]="Budget")`
  - **#Actual RNS:** `=CALCULATE(SUM(FactTransaction[RNS]),FactTransaction[Base]="Actual")`
  - **#LY RNS:**
    ```dax
    =SUMX(
    VALUES('DimCalendar'[Date]),
    CALCULATE(
        [#Actual RNS],
        DATEADD('DimCalendar'[Date], -1, YEAR)
    ))
    ```
  - **#Budget RNS:** `=CALCULATE(SUM(FactTransaction[RNS]),FactTransaction[Base]="Budget")`
  - **#Actual ADR:** `=DIVIDE([#Actual Revenue],[#Actual RNS],0)`
  - **#LY ADR:** `=DIVIDE([#LY Rev],[#LY RNS],0)`
  - **#Budget ADR:** `=DIVIDE([#Budget Revenue],[#Budget RNS],0)`
  - **#Actual Availability:** `=SUM(FactAvailability[Availability])`
  - **#LY Availability:**
    ```dax
    =SUMX(
    VALUES('DimCalendar'[Date]),
    CALCULATE(
        [#Actual Availability],
        DATEADD('DimCalendar'[Date], -1, YEAR)
    ))
    ```
  - **#Actual Occupancy(%):** `=DIVIDE([#Actual RNS],[#Actual Availability],0)`
  - **#LY Occupancy(%):** `=DIVIDE([#LY RNS],[#LY Availability],0)`
  - **#Budget Occupancy(%):** `=DIVIDE([#Budget RNS],[#Actual Availability],0)`
  - **#Actual Revpar:** `=[#Actual ADR]*[#Actual Occupancy(%)]`
  - **#LY Revpar:** `=[#LY ADR]*[#LY Occupancy(%)]`
  - **#Budget Revpar:** `=[#Budget ADR]*[#Budget Occupancy(%)]`
  - **#Budget Revenue YTD:** `=CALCULATE([#Budget Revenue],DATESYTD(DimCalendar[Date]))`
  - **#Actual Revenue YTD:** `=CALCULATE([#Actual Revenue],DATESYTD(DimCalendar[Date]))`
  
  </details>
- Diseño Interactivo: Uso de paginado para navegación y segmentación de datos.

## 🖼️ Vistas Previas del proyecto
<details>
  <summary>Capturas</summary>
    <img width="2101" height="808" alt="image" src="https://github.com/user-attachments/assets/d43168e3-cb4e-4ade-ae5b-6113861f1310" />
    <img width="1512" height="979" alt="image" src="https://github.com/user-attachments/assets/7160fe37-7260-4b29-955d-6bf7435357b7" />
    <img width="1824" height="971" alt="image" src="https://github.com/user-attachments/assets/e70a39e8-cdfe-4629-a9b8-7d966ca106de" />
</details>

<details>
  <summary>Video</summary>
  https://youtu.be/3ah8CH0h_vA
</details>



## 👤 Autor
- Giancarlo Barrantes
- Lima, Perú
- [Linkedin](https://www.linkedin.com/in/gb25/)
