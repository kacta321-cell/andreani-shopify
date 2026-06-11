---
name: andreani-shopify
description: "Convierte pedidos exportados de Shopify al formato Excel de envío masivo de Andreani. Maneja parsing de teléfonos argentinos, matching de 29k localidades, DNI en Billing Company, totales con apostrofe, números de calle con texto. Keywords: andreani, shopify, envios, pedidos, xlsx, csv, logistica, argentina"
compatibility: Claude Code
metadata:
  author: kacta321-cell
  repo: https://github.com/kacta321-cell/andreani-shopify
---

# Skill: Shopify → Andreani

## Tu rol
Convertís pedidos exportados de Shopify al Excel que pide Andreani para envío masivo, sin ningún error de validación.

## Setup — primera vez (automático)

Cuando el usuario diga "instalá el skill de Andreani" o similar, corré:

```powershell
git clone --depth=1 https://github.com/kacta321-cell/andreani-shopify.git "$env:USERPROFILE\Downloads\andreani-shopify-skill"
```

El repo ya incluye el JSON de localidades y la plantilla Excel, así que no necesita descargar nada más.

**Si Python no está instalado**, instalalo automáticamente (Windows):
```powershell
winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
```
En Mac: `brew install python` · En Linux: ya viene. El script instala `openpyxl` solo si falta.

> Para usuarios sin Claude Code (o que no quieren instalar nada): existe una web app que hace lo mismo en el navegador → https://kacta321-cell.github.io/andreani-shopify/

## Flujo de conversión

### Paso 1 — El usuario exporta el CSV desde Shopify
Shopify Admin → Pedidos → Exportar → "Pedidos coinciden con tu búsqueda" → CSV para Excel → guardar en Downloads como `orders_export.csv`

### Paso 2 — Correr el script
```powershell
python "$env:USERPROFILE\Downloads\andreani-shopify-skill\convertir_andreani.py"
```
O si el CSV está en otra ruta:
```powershell
python "$env:USERPROFILE\Downloads\andreani-shopify-skill\convertir_andreani.py" "C:\ruta\al\orders_export.csv"
```

### Paso 3 — Subir a Andreani
Subir el archivo `Andreani_envio.xlsx` generado al portal de Andreani: Envío Masivo → A domicilio.

## Configuración del paquete

Editá estas variables al inicio del script si cambian las medidas:

```python
PESO  = 500   # gramos
ALTO  = 15    # cm
ANCHO = 20    # cm
PROF  = 10    # cm
```

## Mapeo de campos

| Campo Andreani        | Campo Shopify CSV         | Lógica especial |
|-----------------------|--------------------------|-----------------|
| Peso/Alto/Ancho/Prof  | —                        | Hardcodeado en config |
| Valor declarado       | `Total`                  | Strip `'`, convertir a `int` |
| Numero Interno        | `Name`                   | Ej: #1041 |
| Nombre / Apellido     | `Shipping Name`          | Split en primer espacio |
| DNI                   | `Billing Company`        | Shopify guarda el DNI acá |
| Email                 | `Email`                  | Default si vacío |
| Celular código / número | `Phone`                | Strip +54, split código/número |
| Calle                 | `Shipping Address1`      | Nombre de calle |
| Número                | `Shipping Address2`      | Solo dígitos, max 20 chars |
| Provincia/Localidad/CP | Lookup por CP           | 29k localidades, fallback por ciudad |
| Observaciones         | `Notes`                  | |

## Errores comunes de Andreani y solución

| Error Andreani               | Causa                          | Solución en el script |
|------------------------------|--------------------------------|----------------------|
| "No es número"               | Apóstrofe en Total             | `replace("'","")` + `int()` |
| "Lista desplegable inválida" | CP no matcheó ninguna localidad | Revisar CP manualmente |
| "Caracteres inválidos en Número" | Texto en número de calle  | Solo dígitos con `filter(str.isdigit)` |
| "Campo obligatorio - Celular" | Pedido sin teléfono           | Default `11 / 00000000` |
