# Generador de Dashboards Power BI (PBIP)

Script en Python que genera dashboards completos para Power BI Desktop en formato PBIP (carpetas JSON). Solo necesitas tus datos y describir que quieres — el script genera todo: modelo semantico, layout del reporte, tema visual y datos limpios.

## Como funciona

1. Pones tus datos fuente (CSV o Excel) en la carpeta `source/`
2. Describes el dashboard que quieres usando la plantilla de `prompt.md`
3. Claude Code modifica `generate_dashboard.py` segun lo que pediste
4. Ejecutas el script y se generan todos los archivos PBIP
5. Abres `PortfolioDashboard.pbip` en Power BI Desktop

## Estructura

```
├── generate_dashboard.py    # Script generador — este es el unico archivo que se edita
├── prompt.md                # Plantilla de prompt para pedir nuevos dashboards
├── README.md
├── .gitignore
└── source/                  # Aqui van los datos fuente de cada proyecto
```

Los archivos PBIP (`PortfolioDashboard.pbip`, carpetas `Report/` y `SemanticModel/`, `portfolio_theme.json`) y los datos procesados (`source/Portfolio_Data.xlsx`) se generan automaticamente al correr el script y no se suben a git.

## Requisitos

- **Python 3.10+** con `pandas` y `openpyxl`:
  ```
  pip install pandas openpyxl
  ```
- **Power BI Desktop** (marzo 2025 o posterior)

## Uso

### 1. Configurar la ruta base

En `generate_dashboard.py`, cambiar `BASE_DIR`:

```python
BASE_DIR = "/ruta/a/tu/clon/PowerBI_Kevinsito"
```

### 2. Ejecutar

```bash
python generate_dashboard.py
```

### 3. Abrir en Power BI Desktop

1. Abrir `PortfolioDashboard.pbip`
2. Configurar el parametro `ExcelFilePath` con la ruta a `source/Portfolio_Data.xlsx`:
   - **WSL**: `\\wsl.localhost\Ubuntu-24.04\ruta\al\repo\source\Portfolio_Data.xlsx`
   - **Windows**: `C:\ruta\al\repo\source\Portfolio_Data.xlsx`
3. Clic en **Refresh**

### 4. Importante

- **"Upgrade to TMDL"** → Rechazar (clic en "Not now")
- **"Upgrade to PBIR"** → Rechazar (clic en "Not now")

Aceptar cualquiera rompe la compatibilidad con el generador.
