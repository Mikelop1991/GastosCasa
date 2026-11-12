# Gastos Casa

Libreta de gastos de casa: apuntar un pago en tres toques y sacar el mismo Excel de siempre (`GASTOS_CASA_<año>.xlsx`).

- **App**: `index.html` (PWA, sin dependencias). Se publica con GitHub Pages.
- **Datos**: una hoja de Google Sheets con la pestaña `Gastos` (id, fecha, categoria, importe, nota, creado, modificado).
- **Back**: `Code.gs`, pegado en Extensiones → Apps Script de esa hoja e implementado como aplicación web.
- **Excel**: `plantilla.xlsx` es el Excel original vacío. La app lo rellena en el propio aparato (un apunte por fila, fecha y nota en la columna B) y lo descarga.

La app se configura sola al abrir el enlace de instalación `…/GastosCasa/#u=<URL /exec>&k=<clave>`. La clave no está en este repositorio.
