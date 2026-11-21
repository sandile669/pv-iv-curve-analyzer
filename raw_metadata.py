"""Read measured/STC values from common tester summary layouts without fabricating data."""
import io
import re
from pathlib import Path
import pandas as pd
from analysis_engine import string_number

FIELDS = {
 'Pmax (W)': ['pmax','maximum power','max power','pm','pmaxw'],
 'Voc (V)': ['voc','open circuit voltage'],
 'Isc (A)': ['isc','short circuit current'],
 'Vm (V)': ['vm','vmp','vmax','voltage at maximum power','vpm'],
 'Im (A)': ['im','imp','imax','current at maximum power','ipm'],
 'Irradiance (W/m²)': ['irradiance','radiation','solar irradiance','irr'],
 'PV Temp (°C)': ['pv temp','pv temperature','module temperature','cell temperature','tmod','tpv'],
 'En Temp (°C)': ['en temp','environment temperature','ambient temperature','air temperature','tamb','ten'],
 'Incline': ['incline','inclination','tilt'],
 'FF (%)': ['ff','fill factor'],
}
def norm(s): return re.sub(r'[^a-z0-9]','',str(s).lower())
def _metric(col):
    c=norm(col)
    if not c:return None
    for field,aliases in FIELDS.items():
        if any(c==norm(a) or c.startswith(norm(a)) for a in aliases):return field
    return None

def _read_sheets(blob, filename):
    if Path(filename).suffix.lower() in ('.xlsx','.xls'):
        return pd.read_excel(io.BytesIO(blob),sheet_name=None,header=None)
    for enc in ('utf-8-sig','utf-16','latin-1'):
        try:return {'CSV':pd.read_csv(io.BytesIO(blob),header=None,encoding=enc,sep=None,engine='python')}
        except Exception:pass
    raise ValueError('Could not read Raw Data Files. Upload CSV or XLSX.')

def _header(frame):
    for ix in range(min(50,len(frame))):
        values=[str(v).strip() if pd.notna(v) else '' for v in frame.iloc[ix]]
        if any(norm(v) in ('filename','file','filenamecsv','filenamexls') for v in values) and sum(_metric(v) is not None for v in values)>=2:
            df=frame.iloc[ix+1:].copy();df.columns=values;return df
    return None

def _number(v):
    if pd.isna(v):return None
    try:
        n=float(str(v).strip().replace(',','').replace('%',''))
        return n if pd.notna(n) else None
    except (ValueError,TypeError):return None

def _string_id(filename):
    name=str(filename).strip()
    s=string_number(name)
    if s is not None:return s
    m=re.search(r'(?:string|str)[\s_-]*(\d+)(?:\D|$)',name,re.I)
    return int(m.group(1)) if m else None

def _extract(df):
    cols=list(df.columns)
    name_col=next((c for c in cols if norm(c) in ('filename','filenamecsv','file','filenamexls')),None)
    if name_col is None:return {}
    type_col=next((c for c in cols if norm(c) in ('type','condition','measurementtype','testtype','mode','testcondition','standard','item')),None)
    mapping={'measured':{},'stc':{}}
    # Supports 'STC Pmax', 'Pmax STC', 'Pmax(STC)', 'Pmax_stc', 'Pmax corrected'.
    for c in cols:
        n=norm(c)
        is_stc=('stc' in n or 'standardtestcondition' in n or 'corrected' in n)
        stripped=re.sub(r'stc|standard\s*test\s*conditions?|corrected', '',str(c),flags=re.I)
        metric=_metric(stripped)
        if metric:mapping['stc' if is_stc else 'measured'].setdefault(metric,c)
    out={}
    for _,row in df.iterrows():
        filename=str(row[name_col]);s=_string_id(filename)
        if s is None:continue
        condition=str(row[type_col]).strip().lower() if type_col else ''
        # The tester exports three rows per string: Value, STC, and module specification.
        # Never mistake module nameplate specifications for measured or STC test results.
        if type_col and condition not in ('value','stc','measured','measurement','actual','corrected','standard test conditions','standard test condition'):
            continue
        row_is_stc=('stc' in condition or 'standard test' in condition or 'corrected' in condition or bool(re.search(r'\bstc\b',filename,re.I)))
        entry=out.setdefault(s,{'measured':{},'stc':{}})
        for metric,col in mapping['measured'].items():
            val=_number(row[col])
            if val is not None:entry['stc' if row_is_stc else 'measured'][metric]=val
        for metric,col in mapping['stc'].items():
            val=_number(row[col])
            if val is not None:entry['stc'][metric]=val
    return out

def parse_raw_metadata(blob,filename):
    result={}
    for _,frame in _read_sheets(blob,filename).items():
        df=_header(frame)
        if df is None:continue
        for s,data in _extract(df).items():
            entry=result.setdefault(s,{'measured':{},'stc':{}})
            for kind in ('measured','stc'):entry[kind].update(data[kind])
    if not result:raise ValueError('No recognized string summary data found. Please upload your actual Raw Data Files export so its columns can be mapped.')
    return result
