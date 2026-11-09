# screener-datos

Datos diarios para el screener de entradas y salidas (190 valores de EE. UU. y 20 ETF).

- `scripts/descargar.py`: descarga las velas diarias de Yahoo Finance (respaldo: Nasdaq), calcula MA 50, MA 200, RSI 14, cambio y máximo de 52 semanas, y guarda las últimas 60 sesiones de cada valor, más objetivo, consenso, votos y fecha de resultados.
- `.github/workflows/datos.yml`: lo ejecuta varias veces cada día laborable después del cierre de EE. UU.
- `datos/diario.json`: el resultado. Solo incluye valores con la sesión del día; los que faltan aparecen en `cobertura.faltan`.
- `universo.json`: la lista de valores y ETF.

No contiene posiciones ni datos personales.

## Screener Buffett (S&P 500)

- `buffett/screener_buffett.py`: puntúa cada día las ~500 empresas del S&P 500 con los 4 factores y los 5 pasos de Warren Buffett (10–12 años de estados financieros de la SEC; precios y analistas de Yahoo Finance) y calcula valor intrínseco, precio de compra y señal.
- `.github/workflows/buffett.yml`: precarga lo lento a las 17:05 UTC y calcula tras el cierre (20:35 o 21:35 UTC, según el horario de verano de EE. UU.), separado del workflow de datos. Los datos quedan listos antes de las 23:30 de Madrid. Necesita el secreto `SEC_USER_AGENT` («Nombre Apellido email»), que la SEC exige para identificar a quien descarga.
- `buffett/revisiones.csv`: revisiones cualitativas del agente (círculo, foso, dirección, ajuste). Al editarlo, el screener se recalcula.
- `datos/buffett/`: `screener_buffett_agente.csv` (versión compacta para el agente), `screener_buffett.csv` (todas las columnas), `mercado.json` (resumen del día), `cambios.csv` (cambios de señal) e `historico/`.
