# andreani-shopify

Convierte pedidos exportados de Shopify al Excel de envío masivo de Andreani. Sin errores de validación, sin configuración manual.

## Dos formas de usarlo

### 🌐 Opción A — Web app (sin instalar nada)

**→ [kacta321-cell.github.io/andreani-shopify](https://kacta321-cell.github.io/andreani-shopify/)**

Entrás, arrastrás el CSV, descargás el Excel. Todo corre en tu navegador — ningún dato sale de tu compu. Funciona hasta desde el celular.

### 🤖 Opción B — Skill de Claude Code (un solo prompt)

Decile esto a Claude:

> "Instalá el skill andreani-shopify desde https://github.com/kacta321-cell/andreani-shopify"

Claude clona el repo y deja todo listo.

## Requisitos

- Python 3.x instalado
- `openpyxl` (el script lo instala solo si no está)

## Uso

1. Exportá los pedidos desde Shopify Admin → Pedidos → Exportar → CSV para Excel
2. Decile a Claude: **"convertí mis pedidos de Shopify a Andreani"**
3. Claude corre el script y genera `Andreani_envio.xlsx`
4. Subí ese archivo al portal de Andreani → Envío Masivo → A domicilio

## Qué resuelve

- Teléfonos argentinos con `+54`, `54`, `9`, código de área variable
- CPs alfanuméricos (ej: `B1603ADA`) → extrae los 4 dígitos numéricos
- Fallback por nombre de ciudad si el CP no matchea
- Totales con apóstrofe (`'68990`) que Andreani rechaza como texto
- Números de calle con texto (ej: `"123 PB"`) → solo dígitos
- DNI guardado en `Billing Company` (convención de tiendas AR)
- Email vacío → default configurable

## Créditos

Desarrollado y testeado para el mercado argentino por [@kacta321-cell](https://github.com/kacta321-cell).
