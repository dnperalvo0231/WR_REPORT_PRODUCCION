"""Lectura del formato modelo editable, compatible con el libro WOR original.

Los encabezados deben ser idénticos a los de la plantilla. Los ejemplos de
'GUÍA' no se importan. No se presume que una pérdida equivalga a la variación
observada de producción entre días.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from io import BytesIO
import zipfile
import re
import pandas as pd

from wor_parser import WORData, excel_date, number, well_key, _sheet_paths

MODEL_TABS = ["Weekly Summary", "Daily Production", "Reported Losses", "Injection Water", "Gas Distribution", "Forecast"]


@dataclass
class ModelResult:
    wor: WORData | None
    daily: pd.DataFrame
    injection: pd.DataFrame
    gas: pd.DataFrame
    losses: pd.DataFrame


def is_model_xlsx(raw: bytes) -> bool:
    with zipfile.ZipFile(BytesIO(raw)) as z:
        names = set(_sheet_paths(z))
        return "Weekly Summary" in names and "Daily Production" in names


def _sheet(raw: bytes, name: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(BytesIO(raw), sheet_name=name, engine='openpyxl', dtype={'Well':str, 'Reservoir':str})
    except ValueError as exc:
        raise ValueError(f'No existe la hoja «{name}» en la plantilla.') from exc
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how='all')


def _check(df: pd.DataFrame, cols: list[str], name: str) -> None:
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f'En «{name}» faltan columnas: {", ".join(missing)}. No cambies los encabezados del modelo.')


def _dates(df: pd.DataFrame, name: str, col: str='Date') -> pd.DataFrame:
    result = df.copy()
    result[col] = result[col].map(excel_date)
    if result[col].isna().any():
        raise ValueError(f'La hoja «{name}» contiene fecha(s) no válidas. Usa fechas reales de Excel.')
    return result


def _numeric(df: pd.DataFrame, columns: list[str], name: str, *, allow_empty=False) -> pd.DataFrame:
    result = df.copy()
    for col in columns:
        original = result[col]
        result[col] = original.map(lambda v: number(v))
        if not allow_empty and result[col].isna().any():
            raise ValueError(f'La columna «{col}» en «{name}» tiene valores vacíos/no numéricos.')
        if (result[col].dropna() < 0).any():
            raise ValueError(f'La columna «{col}» en «{name}» contiene valores negativos.')
    return result


def read_model(raw: bytes, *, require_weekly: bool=True) -> ModelResult:
    """Lee el modelo completo o sus hojas complementarias.

    Si require_weekly=False, «Weekly Summary» puede estar vacía y se devuelven
    solamente series auxiliares, sin reemplazar cifras del Excel WOR principal.
    """
    if len(raw) > 25*1024*1024:
        raise ValueError('El archivo modelo supera 25 MB.')
    with zipfile.ZipFile(BytesIO(raw)) as z:
        names=set(_sheet_paths(z))
    need = set(MODEL_TABS)
    if not need.issubset(names):
        raise ValueError('Al formato modelo le faltan hojas: '+', '.join(sorted(need-names)))

    raw_week = _sheet(raw, 'Weekly Summary')
    weekly_cols=['Date','Oil Rate BOPD','Water Rate BWPD','Total Fluid BPD','Gas MPCD',
                 'Water Cut %','Delivery BOPD','Delivery BSW %']
    _check(raw_week, weekly_cols, 'Weekly Summary')
    wor=None
    prod=pd.DataFrame(columns=['fecha','petroleo','agua','fluido','gas','corte_agua','entrega','bsw_entrega'])
    if not raw_week.empty:
        frame=_dates(raw_week,'Weekly Summary')
        frame=_numeric(frame,weekly_cols[1:],'Weekly Summary')
        if frame['Date'].duplicated().any():
            raise ValueError('Weekly Summary: hay fechas de campo duplicadas.')
        if len(frame)>14:
            raise ValueError('Weekly Summary: carga máximo 14 días; normalmente serán 7.')
        prod = frame.rename(columns={
            'Date':'fecha','Oil Rate BOPD':'petroleo','Water Rate BWPD':'agua',
            'Total Fluid BPD':'fluido','Gas MPCD':'gas','Water Cut %':'corte_agua',
            'Delivery BOPD':'entrega','Delivery BSW %':'bsw_entrega'
        })[['fecha','petroleo','agua','fluido','gas','corte_agua','entrega','bsw_entrega']]
        if (prod['corte_agua']>100).any() or (prod['bsw_entrega']>100).any():
            raise ValueError('Los campos de porcentaje deben estar entre 0 y 100 (por ejemplo, 0,36 para 0,36 %).')
        prod['bsw_entrega'] = prod['bsw_entrega']/100
        prod = prod.sort_values('fecha').reset_index(drop=True)
    elif require_weekly:
        raise ValueError('«Weekly Summary» está vacía: coloca las 7 filas diarias o sube tu WOR original como archivo principal.')

    raw_daily=_sheet(raw,'Daily Production')
    _check(raw_daily,['Date','Well','Oil Rate BOPD'],'Daily Production')
    daily=pd.DataFrame(columns=['fecha','pozo','bopd','pozo_id'])
    if not raw_daily.empty:
        df=_dates(raw_daily,'Daily Production')
        df=_numeric(df,['Oil Rate BOPD'],'Daily Production')
        df['Well']=df['Well'].astype(str).str.strip()
        if df['Well'].eq('').any() or df['Well'].eq('nan').any():
            raise ValueError('Daily Production: todos los registros necesitan Well.')
        daily=df.rename(columns={'Date':'fecha','Well':'pozo','Oil Rate BOPD':'bopd'})[['fecha','pozo','bopd']]
        daily['pozo_id']=daily['pozo'].map(well_key)
        if daily.duplicated(['fecha','pozo_id']).any():
            raise ValueError('Daily Production: fecha y pozo repetidos (incluye variantes SB-2/B2).')
        daily=daily.reset_index(drop=True)

    raw_loss=_sheet(raw,'Reported Losses')
    loss_cols=['Date','Well','Recoverable BOPD','Non-Recoverable BOPD','Comments']
    _check(raw_loss,loss_cols,'Reported Losses')
    losses=pd.DataFrame(columns=['fecha','pozo','pozo_id','recuperable','no_recuperable','comentario','_excel_row'])
    if not raw_loss.empty:
        df=_dates(raw_loss,'Reported Losses')
        df=_numeric(df,['Recoverable BOPD','Non-Recoverable BOPD'],'Reported Losses',allow_empty=True)
        if df[['Recoverable BOPD','Non-Recoverable BOPD']].isna().all(axis=1).any():
            raise ValueError('Reported Losses: indica al menos una de las dos pérdidas en cada incidente.')
        df[['Recoverable BOPD','Non-Recoverable BOPD']]=df[['Recoverable BOPD','Non-Recoverable BOPD']].fillna(0)
        df['Well']=df['Well'].fillna('').astype(str).str.strip()
        if df['Well'].eq('').any():
            raise ValueError('Reported Losses: indica el pozo en cada incidente.')
        losses=df.rename(columns={'Date':'fecha','Well':'pozo','Recoverable BOPD':'recuperable',
                   'Non-Recoverable BOPD':'no_recuperable','Comments':'comentario'})
        losses['pozo_id']=losses['pozo'].map(well_key)
        losses['comentario']=losses['comentario'].fillna('')
        losses['_excel_row']=range(2,2+len(losses))
        losses=losses[['fecha','pozo','pozo_id','recuperable','no_recuperable','comentario','_excel_row']].reset_index(drop=True)

    raw_water=_sheet(raw,'Injection Water')
    water_cols=['Date','Well','Reservoir','Water BWPD','Pressure psi','ppm SST','ppm oil','ppm Salinity']
    _check(raw_water,water_cols,'Injection Water')
    injection=pd.DataFrame(columns=['fecha','pozo','pozo_id','reservorio','agua','presion','sst','aceite','salinidad'])
    if not raw_water.empty:
        df=_dates(raw_water,'Injection Water')
        df=_numeric(df,['Water BWPD','Pressure psi','ppm SST','ppm oil','ppm Salinity'],
                    'Injection Water',allow_empty=True)
        df['Well']=df['Well'].fillna('').astype(str).str.strip()
        if df['Well'].eq('').any():
            raise ValueError('Injection Water: indica el pozo de inyección.')
        injection=df.rename(columns={'Date':'fecha','Well':'pozo','Reservoir':'reservorio',
                       'Water BWPD':'agua','Pressure psi':'presion','ppm SST':'sst',
                       'ppm oil':'aceite','ppm Salinity':'salinidad'})
        injection['pozo_id']=injection['pozo'].map(well_key)
        injection=injection[['fecha','pozo','pozo_id','reservorio','agua','presion','sst','aceite','salinidad']]
        if injection.duplicated(['fecha','pozo_id']).any():
            raise ValueError('Injection Water: no repitas pozo y fecha; usa una fila por pozo y día.')
        injection=injection.reset_index(drop=True)

    raw_gas=_sheet(raw,'Gas Distribution')
    gas_cols=['Date','Produced Gas Mscf','Energy Gas Mscf','Flared Gas Mscf']
    _check(raw_gas,gas_cols,'Gas Distribution')
    gas=pd.DataFrame(columns=['fecha','producido','combustible','quemado'])
    if not raw_gas.empty:
        df=_dates(raw_gas,'Gas Distribution')
        df=_numeric(df,gas_cols[1:],'Gas Distribution')
        if df['Date'].duplicated().any():
            raise ValueError('Gas Distribution: hay fechas repetidas; usa una fila por día.')
        gas=df.rename(columns={'Date':'fecha','Produced Gas Mscf':'producido',
                'Energy Gas Mscf':'combustible','Flared Gas Mscf':'quemado'})[['fecha','producido','combustible','quemado']]
        gas['ptg']=gas['combustible'];gas['micro']=0
        gas['_excel_row']=range(2,2+len(gas))
        gas=gas[['fecha','ptg','micro','combustible','producido','quemado','_excel_row']]

    raw_fc=_sheet(raw,'Forecast')
    _check(raw_fc,['Month','Actual BOPD','Forecast BOPD'],'Forecast')
    forecast=pd.DataFrame(columns=['fecha','real','pronostico','_excel_row'])
    if not raw_fc.empty:
        df=_dates(raw_fc,'Forecast','Month')
        df=_numeric(df,['Actual BOPD','Forecast BOPD'],'Forecast')
        forecast=df.rename(columns={'Month':'fecha','Actual BOPD':'real','Forecast BOPD':'pronostico'})[['fecha','real','pronostico']]
        forecast['fecha']=forecast['fecha'].map(lambda d:d.replace(day=1))
        forecast['_excel_row']=range(2,2+len(forecast))
        if forecast['fecha'].duplicated().any():
            raise ValueError('Forecast: un registro mensual duplicado.')
    if not prod.empty:
        rld=pd.DataFrame(columns=['pozo','pozo_id','rld_dias','objetivo_dias','fin_rld'])
        wor=WORData(prod,losses,pd.DataFrame(columns=['fecha','a10_agua','a11_agua']),
                    gas,forecast,rld,[],injection)
        if not gas.empty:
            discrepancy=(gas['combustible']+gas['quemado']-gas['producido']).abs()
            if (discrepancy>0.02).any():
                wor.warnings.append('Gas Distribution: Energy + Flared difiere de Produced en uno o más días; revisar balance.')
    return ModelResult(wor,daily,injection,gas,losses)
