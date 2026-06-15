# RPA — Recepción de Facturas | SIPP Petroplazas

<p align="center">
  <img src="Logo_Petroil.png" alt="Petroil Logo" width="100"/>
</p>

<p align="center">
  Automatización de captura de datos en SIPP + comparativa de sucursales y cálculo de distribución de facturas.
</p>

---

## ¿Qué hace?

El proyecto tiene dos componentes principales:

### 🤖 RPA (Robotic Process Automation)
Navega automáticamente el portal **SIPP** (`sipp.petroil.com.mx`) y por cada folio en el archivo Excel:
1. Busca el folio en *Recepción de Facturas*
2. Abre el modal *Visualizar Detalle* → sección *Servicios*
3. Abre el visor de documento OC
4. Extrae: **CC OC**, **Observaciones OC**, **Subtotal**, **Descuento**, **IVA**, **Gastos de Envío**, **Total**
5. Escribe los resultados de vuelta al Excel (columnas AF–AL)

Soporta **sesiones paralelas** (1×, 2×, 4×) para procesar el trabajo varias veces más rápido.

### 📊 Comparativa de Sucursales
Analiza la columna *Observaciones OC* del CSV exportado y clasifica cada registro:

| Resultado | Descripción |
|-----------|-------------|
| ✅ **MATCH** | Una sucursal detectada que coincide con *Grupo Centro de Costo* |
| ❌ **MISMATCH** | Una sucursal detectada que NO coincide |
| 🔵 **DISTRIBUCIÓN** | Múltiples sucursales o una zona detectada — el monto se distribuye entre estaciones |
| ⬜ **Sin sucursal** | No se detectó ninguna sucursal en las observaciones |

Genera un Excel con **5 hojas**: Comparación, Resumen, Por Sucursal, Distribución y **Distribución Calculada** (montos por estación según catálogos de porcentajes).

---

## Instalación

### Windows

```bat
# 1. Clona el repositorio
git clone https://github.com/Anibal-F/RPA_SIPP-Petroplazas_Distribucion.git
cd RPA_SIPP-Petroplazas_Distribucion

# 2. Ejecuta el instalador (doble clic o desde CMD)
setup.bat
```

El script instala dependencias, descarga Chromium, genera el ícono `.ico` y crea un **acceso directo en el escritorio** listo para usar.

### macOS / Linux

```bash
git clone https://github.com/Anibal-F/RPA_SIPP-Petroplazas_Distribucion.git
cd RPA_SIPP-Petroplazas_Distribucion
bash setup.sh
```

### Instalación manual

```bash
pip install -r requirements.txt
playwright install chromium
python main.py          # macOS/Linux
python main.py          # Windows (o doble clic en el acceso directo)
```

---

## Uso

### Interfaz principal

```
python main.py
```

| Control | Función |
|---------|---------|
| **Examinar** | Seleccionar el archivo Excel con los folios (columna D, desde fila 9) |
| **Sesiones paralelas** | `1×` `2×` `4×` — divide el trabajo entre múltiples ventanas de navegador |
| **▶ EJECUTAR RPA** | Inicia la extracción automática en SIPP |
| **⏹ CANCELAR** | Detiene el proceso al terminar el folio actual |
| **📊 COMPARAR CSV** | Abre un CSV ya procesado y genera el Excel de comparativa |
| **📋 CATÁLOGOS** | Editor visual de los catálogos de distribución por zona |
| **🔄** *(header)* | Comprueba manualmente si hay actualizaciones en GitHub |

### Modo CLI (limitar registros para pruebas)

```bash
python main.py --max 10   # procesa solo los primeros 10 folios
```

---

## Estructura del proyecto

```
├── main.py                   # GUI (customtkinter) + orquestación
├── compare_sucursales.py     # Comparativa y cálculo de distribución
├── requirements.txt          # Dependencias Python
├── setup.bat                 # Instalador Windows (con acceso directo)
├── setup.sh                  # Instalador macOS/Linux
├── Logo_Petroil.png          # Logo de la aplicación
│
├── rpa/
│   ├── automation.py         # Playwright: navegación y extracción en SIPP
│   └── excel_handler.py      # Lectura de folios y escritura de resultados
│
└── Distribucion/
    ├── Mazatlan_General.csv  # 20 estaciones · 5% c/u
    ├── Corporativo.csv       # ~45 estaciones · 2.22% c/u
    └── Zonas.csv             # Zonas (MAZATLAN 1–4, CULIACAN, NORTE, etc.)
```

---

## Catálogos de distribución

Los archivos en `Distribucion/` definen cómo se reparte el monto de una factura cuando el registro es **DISTRIBUCIÓN**. Se pueden editar directamente desde la interfaz con el botón **📋 CATÁLOGOS**.

| Catálogo | Uso |
|----------|-----|
| `Mazatlan_General.csv` | Distribución general Mazatlán (20 est.) |
| `Corporativo.csv` | Distribución corporativa (~45 est.) |
| `Zonas.csv` | Zonas específicas: MAZATLAN 1–4, CULIACAN, NORTE, CENTRO, SUR, JALISCO, SONORA, BCS |

**Lógica de distribución:**
- Si se detecta `ZONA MAZATLAN 2` → se busca en `Zonas.csv` y se aplican los % de esa zona
- Si se detectan estaciones individuales → split equitativo
- `ZONA CLN` → alias de `ZONA CULIACAN`
- `ZONA MAZATLAN` (sin número) → `MAZATLAN GRAL` en `Mazatlan_General.csv`

---

## Actualizaciones automáticas

Al iniciar la app, se ejecuta `git fetch` en segundo plano. Si hay commits nuevos en GitHub, aparece un banner verde con el botón **⬇ Actualizar y reiniciar** que ejecuta `git pull` y reinicia la app automáticamente.

> **Requisito:** tener [Git](https://git-scm.com) instalado y en el PATH.

---

## Requisitos del sistema

| Componente | Versión mínima |
|------------|----------------|
| Python | 3.10+ |
| Playwright | 1.44+ |
| customtkinter | 5.2+ |
| openpyxl | 3.1+ |
| Pillow | 10.0+ |
| Git | cualquiera (para auto-update) |

---

## Notas

- Los archivos de datos (`Recepcion_Facturas/`) están excluidos del repositorio por contener información sensible del negocio.
- Las credenciales de SIPP **no se almacenan** — se ingresan en la interfaz cada sesión.
- En Windows, la app se lanza con `pythonw.exe` (sin ventana de consola) cuando se usa el acceso directo del escritorio.
