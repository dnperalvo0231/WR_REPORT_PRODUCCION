"""Lector especializado del formato WOR Gente Oil.

La hoja principal del archivo original puede tener ~100 MB de XML interno por
formato. Por eso se leen únicamente las filas semanales necesarias, sin cargar
el libro completo en openpyxl. Ningún dato se cambia dentro del Excel.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from io import BytesIO
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

import pandas as pd

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PACKAGE_REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
BASE_DATE = datetime(1899, 12, 30)


@dataclass
class WORData:
    production: pd.DataFrame
    losses: pd.DataFrame
    water: pd.DataFrame
    gas: pd.DataFrame
    forecast: pd.DataFrame
    rld: pd.DataFrame
    warnings: list[str] = field(default_factory=list)
    injection_records: pd.DataFrame = field(default_factory=pd.DataFrame)

    @property
    def start_date(self) -> date:
        return self.production['fecha'].min()

    @property
    def end_date(self) -> date:
        return self.production['fecha'].max()


def excel_date(value) -> date | None:
    if value is None or value == '':
        return None
    if isinstance(value, (date, datetime)):
        return value.date() if isinstance(value, datetime) else value
    if isinstance(value, (float, int)) or (isinstance(value, str) and re.fullmatch(r'\d+(\.\d+)?', value.strip())):
        v = float(value)
        if 30000 <= v < 90000:
            return (BASE_DATE + timedelta(days=v)).date()
        return None
    s = str(value).strip()
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', s):
        try:
            return datetime.strptime(s, '%Y-%m-%d').date()
        except ValueError:
            return None
    try:
        return pd.to_datetime(value, dayfirst=True, errors='raise').date()
    except (ValueError, TypeError, OverflowError):
        return None


def number(value, default=None):
    if value is None or value == '':
        return default
    if isinstance(value, (int, float)):
        return float(value)
    value = str(value).strip().replace(' ', '')
    # Se respetan decimales europeos y estadounidenses, sin convertir 4.089 a 4089
    # arbitrariamente: los valores numéricos originales de Excel ya tienen su escala.
    if ',' in value and '.' in value:
        value = value.replace('.', '').replace(',', '.') if value.rfind(',') > value.rfind('.') else value.replace(',', '')
    elif ',' in value:
        value = value.replace(',', '.')
    try:
        return float(value)
    except ValueError:
        return default


def well_key(text) -> str:
    """SB-2 'UI', SB2 UI, B2 -> B2. Conserva ST1: B18ST1 != B18."""
    if text is None:
        return ''
    s = str(text).upper().replace('SINGUE', '').replace('"', '').replace("'", '').strip()
    m = re.search(r'(?:^|\b)S?([ABN])\s*[- ]?\s*(\d+)\s*(ST\d+|RE|H)?', s)
    if not m:
        return re.sub(r'[^A-Z0-9]', '', s)
    return f'{m.group(1)}{m.group(2)}{m.group(3) or ""}'


def _col(ref: str) -> str:
    return re.match(r'[A-Z]+', ref).group(0)


def _shared_strings(z) -> list[str]:
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    with z.open('xl/sharedStrings.xml') as f:
        root = ET.parse(f).getroot()
    return [''.join(t.text or '' for t in si.iter(NS+'t')) for si in root.findall(NS+'si')]


def _sheet_paths(z) -> dict[str, str]:
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    paths = {r.get('Id'): r.get('Target') for r in rels.findall(PACKAGE_REL+'Relationship')}
    result = {}
    for sheet in wb.findall('.//'+NS+'sheet'):
        target = paths.get(sheet.get(REL+'id'), '')
        path = target.lstrip('/') if target.startswith('/') else 'xl/' + target.lstrip('/')
        if path in z.namelist():
            result[sheet.get('name')] = path
    return result


def _rows(z, path: str, shared: list[str], max_row: int | None = None):
    """Iterparse de una hoja. max_row corta la lectura del XML enorme."""
    with z.open(path) as stream:
        for _, row in ET.iterparse(stream, events=('end',)):
            if row.tag != NS+'row':
                continue
            row_i = int(row.get('r', '0'))
            if max_row is not None and row_i > max_row:
                break
            cells = {}
            for c in row.findall(NS+'c'):
                v = c.find(NS+'v')
                inline = c.find(NS+'is')
                value = v.text if v is not None else None
                if c.get('t') == 's' and value not in (None, ''):
                    i = int(value)
                    value = shared[i] if 0 <= i < len(shared) else None
                elif c.get('t') == 'inlineStr' and inline is not None:
                    value = ''.join(t.text or '' for t in inline.iter(NS+'t'))
                if value is not None:
                    cells[_col(c.get('r'))] = value
            yield row_i, cells
            row.clear()


def _extract(rows, start_row: int, key_col: str, mapping: dict[str, str], *, end_row=None) -> pd.DataFrame:
    result = []
    for idx, values in rows:
        if idx < start_row:
            continue
        if end_row is not None and idx > end_row:
            break
        date_value = excel_date(values.get(key_col))
        if date_value is None:
            continue
        out = {'fecha': date_value}
        out.update({dest: number(values.get(col)) for dest, col in mapping.items()})
        out['_excel_row'] = idx
        result.append(out)
    return pd.DataFrame(result, columns=['fecha', *mapping.keys(), '_excel_row'])


def load_wor(file_bytes: bytes) -> WORData:
    if len(file_bytes) > 55 * 1024 * 1024:
        raise ValueError('El archivo supera 55 MB; verifica que sea el WOR en formato .xlsx.')
    try:
        z = zipfile.ZipFile(BytesIO(file_bytes))
    except (zipfile.BadZipFile, ValueError) as exc:
        raise ValueError('El archivo no es un Excel .xlsx válido.') from exc
    with z:
        paths = _sheet_paths(z)
        required = ['Production and Pumping', 'Produccion Losses', 'BSW', 'GAS']
        missing = [name for name in required if name not in paths]
        if missing:
            raise ValueError('Faltan hojas obligatorias del formato WOR: '+', '.join(missing))
        ss = _shared_strings(z)
        p = _extract(_rows(z, paths['Production and Pumping'], ss, max_row=16), 5, 'A', {
            'petroleo': 'B', 'agua': 'C', 'fluido': 'D', 'gas': 'E',
            'corte_agua': 'F', 'entrega': 'G', 'bsw_entrega': 'H'
        }, end_row=16)
        if p.empty:
            raise ValueError('No se identificaron fechas semanales en Production and Pumping (filas 5-11).')
        # Solo la primera semana (no los bloques de producción mensual).
        p = p.sort_values('fecha').drop_duplicates('fecha').reset_index(drop=True)
        if len(p) > 14:
            raise ValueError('La sección semanal contiene más de 14 días; revisa el formato.')

        losses_data = []
        for idx, c in _rows(z, paths['Produccion Losses'], ss):
            if idx < 5:
                continue
            d = excel_date(c.get('B'))
            well = str(c.get('C', '')).strip()
            if not d or not well:
                continue
            r, n = number(c.get('D'), 0.0), number(c.get('E'), 0.0)
            if r < 0 or n < 0:
                continue
            losses_data.append({'fecha': d, 'pozo': well, 'pozo_id': well_key(well),
                                'recuperable': r, 'no_recuperable': n,
                                'comentario': str(c.get('F', '')).strip(), '_excel_row': idx})
        losses = pd.DataFrame(losses_data, columns=['fecha', 'pozo', 'pozo_id', 'recuperable', 'no_recuperable', 'comentario', '_excel_row'])

        water = _extract(_rows(z, paths['BSW'], ss, max_row=14), 5, 'B', {
            'a10_agua': 'C', 'a10_presion': 'D', 'a10_sst': 'E', 'a10_aceite': 'F',
            'a11_agua': 'G', 'a11_presion': 'H', 'a11_sst': 'I', 'a11_aceite': 'J'
        }, end_row=14)
        gas = _extract(_rows(z, paths['GAS'], ss, max_row=13), 4, 'B', {
            'ptg': 'C', 'micro': 'D', 'combustible': 'E', 'producido': 'F', 'quemado': 'G'
        }, end_row=13)
        forecast = pd.DataFrame(columns=['fecha', 'real', 'pronostico', '_excel_row'])
        if 'FORECAST' in paths:
            forecast = _extract(_rows(z, paths['FORECAST'], ss, max_row=120), 4, 'A', {
                'real': 'B', 'pronostico': 'C'
            }, end_row=120)
            if not forecast.empty:
                forecast = forecast[forecast['fecha'].apply(lambda d: d.day == 1)]
        rld_data = []
        if 'RLD' in paths:
            for idx, c in _rows(z, paths['RLD'], ss, max_row=28):
                if idx < 4 or not c.get('B'):
                    continue
                days = number(c.get('F'))
                if days is None:
                    continue
                rld_data.append({'pozo': str(c['B']).strip(), 'pozo_id': well_key(c['B']),
                                 'rld_dias': days, 'objetivo_dias': number(c.get('G')),
                                 'fin_rld': excel_date(c.get('E'))})
        rld = pd.DataFrame(rld_data, columns=['pozo','pozo_id','rld_dias','objetivo_dias','fin_rld'])
    issues = []
    for name, df in [('agua', water), ('gas', gas)]:
        if df.empty:
            issues.append('No se encontraron datos diarios en la hoja de '+name+'.')
        elif set(df['fecha']) != set(p['fecha']):
            issues.append(f'Las fechas de {name} no coinciden exactamente con las de producción.')
    if p['petroleo'].isna().any() or p['agua'].isna().any():
        issues.append('Faltan tasas de producción en uno o más días de la semana.')
    if not rld.empty and rld.fin_rld.notna().any() and rld.fin_rld.max() > p['fecha'].max():
        issues.append('La hoja RLD contiene fechas posteriores al último día del reporte; revise su corte antes de presentarla como indicador semanal.')
    return WORData(p, losses, water, gas, forecast, rld, issues)


def load_daily_wells(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Acepta CSV o XLSX con fecha, pozo y petroleo_bopd, sin imputar ceros."""
    if filename.lower().endswith('.csv'):
        for enc in ['utf-8-sig', 'latin-1']:
            try:
                df = pd.read_csv(BytesIO(file_bytes), sep=None, engine='python', encoding=enc)
                break
            except UnicodeDecodeError:
                continue
    elif filename.lower().endswith('.xlsx'):
        df = pd.read_excel(BytesIO(file_bytes), engine='openpyxl')
    else:
        raise ValueError('Sube un archivo CSV o XLSX.')
    def norm(x):
        s = unicodedata.normalize('NFKD', str(x)).encode('ascii', 'ignore').decode().lower().strip()
        return re.sub('[^a-z0-9]', '', s)
    names = {norm(c): c for c in df.columns}
    wanted = [('fecha', ['fecha','date','dia']), ('pozo', ['pozo','well','wellname']),
              ('bopd', ['petroleobopd','bopd','oilbopd','oilrate','produccionbopd'])]
    cols = {}
    for dest, alts in wanted:
        original = next((names[n] for n in alts if n in names), None)
        if original is None:
            raise ValueError('La tabla por pozo debe contener: fecha, pozo, petroleo_bopd.')
        cols[dest] = original
    out = df[[cols[k] for k in ['fecha','pozo','bopd']]].copy()
    out.columns = ['fecha','pozo','bopd']
    # pd.to_datetime: distinguir serial Excel, dd/mm/aaaa e ISO.
    out['fecha'] = out['fecha'].map(excel_date)
    out['pozo_id'] = out['pozo'].map(well_key)
    out['bopd'] = pd.to_numeric(out['bopd'], errors='coerce')
    out = out.dropna(subset=['fecha','bopd'])
    out = out[(out['pozo_id'] != '') & (out['bopd'] >= 0)]
    if out.duplicated(['fecha', 'pozo_id']).any():
        raise ValueError('Hay registros duplicados de producción para el mismo pozo y fecha. Corrígelos antes de comparar.')
    return out.reset_index(drop=True)


def match_loss_to_daily(loss: pd.Series, daily: pd.DataFrame):
    """Devuelve producción del MISMO pozo en D-1 y D; nunca usa el total del campo."""
    if daily is None or daily.empty:
        return None, None
    well = loss['pozo_id']
    previous = loss['fecha'] - timedelta(days=1)
    prev_rows = daily[(daily['pozo_id'] == well) & (daily['fecha'] == previous)]
    curr_rows = daily[(daily['pozo_id'] == well) & (daily['fecha'] == loss['fecha'])]
    prev = float(prev_rows.iloc[0]['bopd']) if len(prev_rows) else None
    curr = float(curr_rows.iloc[0]['bopd']) if len(curr_rows) else None
    return prev, curr
