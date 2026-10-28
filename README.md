# screener-datos

Datos diarios para el screener de entradas y salidas (190 valores de EE. UU. y 20 ETF).

- `scripts/descargar.py`: descarga las velas diarias de Yahoo Finance (respaldo: Nasdaq), calcula MA 50, MA 200, RSI 14, cambio y máximo de 52 semanas, y guarda las últimas 60 sesiones de cada valor, más objetivo, consenso, votos y fecha de resultados.
- `.github/workflows/datos.yml`: lo ejecuta varias veces cada día laborable después del cierre de EE. UU.
- `datos/diario.json`: el resultado. Solo incluye valores con la sesión del día; los que faltan aparecen en `cobertura.faltan`.
- `universo.json`: la lista de valores y ETF.

No contiene posiciones ni datos personales.
