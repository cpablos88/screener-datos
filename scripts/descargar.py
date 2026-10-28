#!/usr/bin/env python3
"""Descarga diaria de datos para el screener de entradas y salidas.

Se ejecuta en GitHub Actions después del cierre de EE. UU. y escribe datos/diario.json.

- Precios (OHLCV diario, 14 meses): Yahoo Finance con yfinance. Si un valor falla o no trae
  la sesión del día, se pide a Nasdaq.
- Calcula MA 50, MA 200, RSI 14 (Wilder), cambio a 52 semanas, máximo y mínimo de 52 semanas,
  volumen medio de 20 sesiones y guarda las últimas 60 sesiones OHLCV de cada valor.
- Analistas (objetivo medio, consenso, votos por mes, fecha de resultados, beta): Yahoo Finance,
  una vez por sesión.
- Nunca publica precios de otra sesión: un valor sin la sesión del día va a «faltan» sin precio.
"""
import datetime as dt
import json
import os
import sys
import time
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import yfinance as yf

NY = ZoneInfo('America/New_York')
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'datos', 'diario.json')
MSG = os.path.join(ROOT, 'datos', 'mensaje.txt')
UNIV = json.load(open(os.path.join(ROOT, 'universo.json'), encoding='utf-8'))
FILAS = 60                      # sesiones OHLCV guardadas por valor (la fase 2 usa 50)
CIERRE_OK = dt.time(16, 20)     # hora de Nueva York desde la que la vela del día cuenta como cerrada
UA = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/124.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Origin': 'https://www.nasdaq.com',
    'Referer': 'https://www.nasdaq.com/',
}
CONSENSO = {'strong_buy': 'Strong Buy', 'buy': 'Buy', 'hold': 'Hold', 'underperform': 'Sell',
            'sell': 'Sell', 'strong_sell': 'Strong Sell'}


def log(*a):
    print(*a, flush=True)


def ahora_ny():
    return dt.datetime.now(NY)


def r2(v):
    return None if v is None else round(float(v), 2)


def limpiar(barras):
    """Quita la vela de hoy si el mercado aún no ha cerrado. barras: [(date, o, h, l, c, v)] antigua → reciente."""
    if not barras:
        return []
    n = ahora_ny()
    if barras[-1][0] == n.date() and n.time() < CIERRE_OK:
        return barras[:-1]
    return barras


# ---------------------------------------------------------------- precios
def yahoo_lote(tickers):
    out = {}
    for intento in range(3):
        faltan = [t for t in tickers if t not in out]
        if not faltan:
            break
        try:
            df = yf.download(faltan, period='14mo', interval='1d', group_by='ticker', auto_adjust=False,
                             actions=False, threads=True, progress=False)
        except Exception as e:  # noqa: BLE001
            log('yfinance: error en la descarga por lotes:', e)
            time.sleep(20)
            continue
        for t in faltan:
            try:
                sub = df[t] if isinstance(df.columns, pd.MultiIndex) else df
                sub = sub[['Open', 'High', 'Low', 'Close', 'Volume']].dropna(subset=['Open', 'High', 'Low', 'Close'])
                b = [(i.date(), float(r['Open']), float(r['High']), float(r['Low']), float(r['Close']),
                      0.0 if pd.isna(r['Volume']) else float(r['Volume'])) for i, r in sub.iterrows()]
                if len(b) >= 60:
                    out[t] = b
            except Exception:  # noqa: BLE001
                pass
        log(f'yfinance intento {intento + 1}: {len(out)}/{len(tickers)}')
        time.sleep(5)
    return out


def _num(s):
    if s is None:
        return None
    s = str(s).replace('$', '').replace(',', '').strip()
    try:
        return float(s)
    except ValueError:
        return None


def nasdaq(t, clase):
    """Histórico diario de Nasdaq (respaldo). clase: 'stocks' o 'etf'."""
    hoy = ahora_ny().date()
    params = {'assetclass': clase, 'fromdate': (hoy - dt.timedelta(days=430)).isoformat(),
              'todate': hoy.isoformat(), 'limit': '9999'}
    simbolos = [t] if '-' not in t else [t.replace('-', '.'), t.replace('-', '%25sl%25'), t]
    for s in simbolos:
        try:
            r = requests.get(f'https://api.nasdaq.com/api/quote/{s}/historical', params=params, headers=UA, timeout=25)
            filas = ((r.json().get('data') or {}).get('tradesTable') or {}).get('rows') or []
        except Exception:  # noqa: BLE001
            filas = []
        b = []
        for x in filas:
            try:
                f = dt.datetime.strptime(x['date'], '%m/%d/%Y').date()
            except Exception:  # noqa: BLE001
                continue
            o, h, lo, c, v = (_num(x.get(k)) for k in ('open', 'high', 'low', 'close', 'volume'))
            if None in (o, h, lo, c):
                continue
            b.append((f, o, h, lo, c, v or 0.0))
        if len(b) >= 60:
            return sorted(b)
        time.sleep(1)
    return None


# ---------------------------------------------------------------- indicadores
def rsi_wilder(c, n=14):
    if len(c) < n + 1:
        return None
    d = [c[i] - c[i - 1] for i in range(1, len(c))]
    g = [max(x, 0.0) for x in d]
    p = [max(-x, 0.0) for x in d]
    ag, ap = sum(g[:n]) / n, sum(p[:n]) / n
    for i in range(n, len(d)):
        ag = (ag * (n - 1) + g[i]) / n
        ap = (ap * (n - 1) + p[i]) / n
    return 100.0 if ap == 0 else 100 - 100 / (1 + ag / ap)


def metricas(b):
    f = b[-1][0]
    c = [x[4] for x in b]
    hace1a = f - dt.timedelta(days=365)
    antes = [x for x in b if x[0] <= hace1a]
    ano = [x for x in b if x[0] > hace1a]
    return dict(
        fecha=f.isoformat(),
        p=r2(c[-1]),
        ma50=r2(sum(c[-50:]) / 50) if len(c) >= 50 else None,
        ma200=r2(sum(c[-200:]) / 200) if len(c) >= 200 else None,
        rsi=r2(rsi_wilder(c)),
        ch52=r2((c[-1] / antes[-1][4] - 1) * 100) if antes else None,
        hi52=r2(max(x[2] for x in ano)),
        lo52=r2(min(x[3] for x in ano)),
        vol20=round(sum(x[5] for x in b[-20:]) / 20),
        ohlcv=[[x[0].isoformat(), r2(x[1]), r2(x[2]), r2(x[3]), r2(x[4]), int(x[5])] for x in reversed(b[-FILAS:])],
    )


# ---------------------------------------------------------------- analistas
def analistas(t, sesion):
    tk = yf.Ticker(t)
    info = tk.get_info() or {}
    a = dict(pt=r2(info.get('targetMeanPrice')), pt_lo=r2(info.get('targetLowPrice')),
             pt_hi=r2(info.get('targetHighPrice')), nan=info.get('numberOfAnalystOpinions'),
             cons=CONSENSO.get(info.get('recommendationKey')), beta=r2(info.get('beta')))
    earn = None
    try:
        cal = tk.calendar or {}
        fechas = cal.get('Earnings Date') if isinstance(cal, dict) else None
        if fechas:
            ses = dt.date.fromisoformat(sesion)
            futuras = [d for d in fechas if d >= ses]
            earn = min(futuras or fechas).isoformat()
    except Exception:  # noqa: BLE001
        pass
    if not earn:
        ts = info.get('earningsTimestampStart') or info.get('earningsTimestamp')
        if ts:
            earn = dt.datetime.fromtimestamp(ts, NY).date().isoformat()
    a['earn'] = earn
    try:
        rt = tk.recommendations
        votos = {}
        for _, r in rt.iterrows():
            votos[str(r['period'])] = [int(r.get(k, 0) or 0) for k in ('strongBuy', 'buy', 'hold', 'sell', 'strongSell')]
        a['votos'] = votos or None
    except Exception:  # noqa: BLE001
        a['votos'] = None
    if a['pt'] is None and a['cons'] is None and a['votos'] is None:
        raise RuntimeError('sin datos de analistas')
    return a


# ---------------------------------------------------------------- principal
def main():
    previo = json.load(open(OUT, encoding='utf-8')) if os.path.exists(OUT) else {}
    acc, etf = UNIV['acciones'], UNIV['etf']

    spy = limpiar(yahoo_lote(['SPY']).get('SPY') or nasdaq('SPY', 'etf') or [])
    if not spy:
        log('Sin datos de SPY en Yahoo ni en Nasdaq: no se guarda nada.')
        sys.exit(1)
    sesion = spy[-1][0].isoformat()
    log('Sesión de referencia (último cierre de SPY):', sesion)

    cob = previo.get('cobertura', {})
    completo = (previo.get('sesion') == sesion and not cob.get('faltan') and not cob.get('etf_faltan')
                and previo.get('analistas_sesion') == sesion)
    if completo and os.environ.get('FORZAR') != '1':
        log('La sesión', sesion, 'ya está completa. Nada que hacer.')
        return

    todos = acc + etf
    yb = yahoo_lote(todos)
    barras, fuente = {}, {}
    for t in todos:
        b = limpiar(yb.get(t) or [])
        if b and b[-1][0].isoformat() == sesion:
            barras[t], fuente[t] = b, 'Yahoo'
            continue
        nb = limpiar(nasdaq(t, 'etf' if t in etf else 'stocks') or [])
        if nb and nb[-1][0].isoformat() == sesion:
            barras[t], fuente[t] = nb, 'Nasdaq'
        else:
            log('Sin la sesión', sesion, 'para', t)
        time.sleep(0.5)

    prevS = previo.get('S', {})
    hacer_an = previo.get('analistas_sesion') != sesion
    an_ok = 0
    S = {}
    for t in acc:
        x = metricas(barras[t]) if t in barras else {'fecha': None}
        x['fuente'] = fuente.get(t)
        an = {k: prevS.get(t, {}).get(k) for k in ('pt', 'pt_lo', 'pt_hi', 'nan', 'cons', 'beta', 'earn', 'votos', 'an_fecha')}
        if hacer_an:
            for _ in range(3):
                try:
                    a = analistas(t, sesion)
                    an.update({k: v for k, v in a.items() if v is not None})
                    an['an_fecha'] = sesion
                    an_ok += 1
                    break
                except Exception as e:  # noqa: BLE001
                    log('Analistas', t, ':', e)
                    time.sleep(8)
            time.sleep(0.4)
        x.update(an)
        S[t] = x

    E = {}
    for t in etf:
        if t in barras:
            b = barras[t]
            m = metricas(b)
            E[t] = dict(fecha=m['fecha'], c=[r2(x[4]) for x in reversed(b[-FILAS:])],
                        d=[x[0].isoformat() for x in reversed(b[-FILAS:])],
                        ma50=m['ma50'], ma200=m['ma200'], ch52=m['ch52'], hi52=m['hi52'], fuente=fuente[t])
        else:
            E[t] = dict(fecha=None)

    faltan = [t for t in acc if S[t].get('fecha') != sesion]
    efaltan = [t for t in etf if E[t].get('fecha') != sesion]
    an_hecho = (not hacer_an) or an_ok >= 0.9 * len(acc)
    out = dict(
        version=1,
        generado=dt.datetime.now(dt.timezone.utc).isoformat(timespec='seconds'),
        sesion=sesion,
        analistas_sesion=sesion if an_hecho else previo.get('analistas_sesion'),
        fuentes=dict(precios='Yahoo Finance (yfinance); respaldo Nasdaq', analistas='Yahoo Finance'),
        cobertura=dict(acciones_ok=len(acc) - len(faltan), acciones_total=len(acc), faltan=faltan,
                       etf_ok=len(etf) - len(efaltan), etf_total=len(etf), etf_faltan=efaltan,
                       analistas_ok=an_ok if hacer_an else None),
        excluidas=UNIV.get('excluidas', {}),
        etf=E,
        S=S,
    )
    if previo.get('sesion') == sesion and previo.get('cobertura', {}).get('acciones_ok', 0) > out['cobertura']['acciones_ok']:
        log('Esta descarga cubre menos valores que la ya guardada para la misma sesión: no se guarda.')
        return
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
    msg = (f"datos {sesion}: {out['cobertura']['acciones_ok']}/{len(acc)} valores, "
           f"{out['cobertura']['etf_ok']}/{len(etf)} ETF" + ('' if an_hecho else ', analistas incompletos'))
    with open(MSG, 'w', encoding='utf-8') as f:
        f.write(msg)
    log(msg)
    if faltan or efaltan:
        log('Faltan:', faltan, efaltan)


if __name__ == '__main__':
    main()
