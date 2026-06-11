# andreani-shopify

Skill para Claude Code que convierte pedidos exportados de Shopify al Excel de envío masivo de Andreani. Sin errores de validación, sin configuración manual.

## Instalación (un solo prompt)

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
