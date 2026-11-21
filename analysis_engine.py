"""Deterministic IV measurement processing. No AI-based modification of measured values."""
import io
import re
import numpy as np
import pandas as pd

VOLTAGE_ALIASES = ('V(V)', 'voltage', 'voltage(v)', 'voltage(vdc)', 'volt', 'v', 'u(v)', 'pv voltage')
CURRENT_ALIASES = ('I(A)', 'current', 'current(a)', 'amp', 'amps', 'i', 'i(a)', 'pv current')

def clean(s):
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

def identify_columns(df):
    cols = {clean(c): c for c in df.columns}
    def choose(names):
        for n in names:
            if clean(n) in cols:
                return cols[clean(n)]
        for c in df.columns:
            cc = clean(c)
            if any(cc.startswith(clean(n)) for n in names if len(clean(n)) > 3):
                return c
        return None
    return choose(VOLTAGE_ALIASES), choose(CURRENT_ALIASES)

def parse_csv(blob):
    errors = []
    for enc in ('utf-8-sig', 'utf-16', 'latin-1'):
        for sep in (None, ',', ';', '\t'):
            try:
                raw = pd.read_csv(io.BytesIO(blob), encoding=enc, sep=sep, engine='python', on_bad_lines='skip')
                vcol, icol = identify_columns(raw)
                if vcol is None or icol is None:
                    # Some instruments prepend metadata rows; scan first 30 lines for header.
                    for skip in range(1, 31):
                        test = pd.read_csv(io.BytesIO(blob), encoding=enc, sep=sep, engine='python', skiprows=skip, on_bad_lines='skip')
                        vcol, icol = identify_columns(test)
                        if vcol is not None and icol is not None:
                            raw = test
                            break
                if vcol is None or icol is None:
                    continue
                d = pd.DataFrame({'V': pd.to_numeric(raw[vcol], errors='coerce'), 'I': pd.to_numeric(raw[icol], errors='coerce')})
                d = d.replace([np.inf, -np.inf], np.nan).dropna()
                d = d[(d.V >= 0) & (d.I >= 0)].sort_values('V').drop_duplicates('V', keep='first')
                if len(d) < 3:
                    raise ValueError('Fewer than three valid nonnegative V/I points')
                return d.reset_index(drop=True)
            except Exception as exc:
                errors.append(str(exc))
    raise ValueError('Could not locate voltage/current columns. Expected V(V) and I(A) (also supports Voltage (V) and Current (A)).')

def string_number(filename):
    """Extract the full string number from b1-inv1-s12(1).csv, etc."""
    from pathlib import Path
    stem = Path(str(filename)).name.rsplit('.', 1)[0].lower().strip()
    match = re.search(r'(?:^|[-_\s])(?:string|str|s)[-_\s]*(\d+)(?=$|[^0-9])', stem)
    if match:
        return int(match.group(1))
    return None


def analyze(df):
    v = df.V.to_numpy(dtype=float)
    i = df.I.to_numpy(dtype=float)
    p = v * i
    idx = int(np.argmax(p))
    # Endpoints are estimated only if the measured sweep does not include the axes.
    isc = float(i[0]) if v[0] <= 0.01 else float(i[0] + (0-v[0])*(i[1]-i[0])/(v[1]-v[0]))
    voc = float(v[-1]) if i[-1] <= 0.01 else float(v[-1] + (0-i[-1])*(v[-1]-v[-2])/(i[-1]-i[-2])) if i[-1] != i[-2] else float(v[-1])
    isc = max(0.0, isc)
    voc = max(0.0, voc)
    pm = float(p[idx]); vm = float(v[idx]); im = float(i[idx]); ff = 100 * pm/(voc*isc) if voc*isc > 0 else float('nan')
    if not np.isfinite(ff):
        status = 'Review'
    elif ff < 75:
        status = 'Abnormal'
    else:
        status = 'Normal'
    warnings = []
    if v[0] > 0.01: warnings.append('Isc extrapolated (sweep does not reach V=0)')
    if i[-1] > 0.01: warnings.append('Voc extrapolated (sweep does not reach I=0)')
    if ff > 100 or ff < 0: status = 'Review'; warnings.append('Nonphysical FF; inspect CSV units/curve')
    return {'Isc (A)':isc,'Voc (V)':voc,'Pmax (W)':pm,'Vm (V)':vm,'Im (A)':im,'FF (%)':ff,'Result':status,'Warnings':'; '.join(warnings)}
