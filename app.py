"""WOR Executive Dashboard: cargar Excel semanal -> dos láminas + PDF 2 páginas.

Ejecución: streamlit run app.py
"""
from __future__ import annotations
from datetime import timedelta
from hashlib import sha256
import hmac
from io import BytesIO
from pathlib import Path
from PIL import Image, ImageOps

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from wor_parser import load_wor, load_daily_wells, match_loss_to_daily
from wor_model import is_model_xlsx, read_model
from dataclasses import replace
from report_export import build_exports, fmt, injection_rows

st.set_page_config(page_title='WOR | Singue Operations',page_icon='🛢️',layout='wide',initial_sidebar_state='expanded')

C = {'oil':'#07806F','water':'#27B5D3','gas':'#E9AC26','loss':'#D7504F',
     'recovery':'#E58D32','delivery':'#075D55','ink':'#172E38','muted':'#687C85', 'gray':'#90A5AB'}

st.markdown('''
<style>
:root {color-scheme: light;}
[data-testid="stAppViewContainer"] {background:#F2F6F7;color:#152F3A;}
[data-testid="stSidebar"] {background:#EAF2F2;}
.block-container {padding-top:1.25rem;padding-bottom:1.8rem;max-width:1720px;}
[data-testid="stMetric"] {background:#fff;border:1px solid #e6eded;border-radius:13px;padding:14px 16px;min-height:110px;}
[data-testid="stMetricLabel"] {font-weight:650!important;color:#5F757D!important;}
[data-testid="stMetricValue"] {color:#1C333C!important;font-size:1.7rem!important;}
h1,h2,h3 {color:#17323B!important;letter-spacing:-.02em;}
div[data-baseweb="tab-list"] {gap:11px;}
button[data-baseweb="tab"] {font-weight:700;border-radius:12px 12px 0 0;}
[data-testid="stAlert"] {border-radius:12px;}
.smallnote {font-size:.82rem;color:#667E87;line-height:1.45}
.head {padding:15px 20px;background:#fff;border-left:6px solid #007A69;border-radius:12px;margin-bottom:18px;}
.head h2 {margin:0;font-size:1.8rem;}
.head p {margin:6px 0 0;color:#617782;}
</style>
''',unsafe_allow_html=True)


def require_password():
    """Protección básica OPCIONAL; no sustituye SSO o un servidor privado."""
    try:
        pwd=st.secrets.get('APP_PASSWORD','')
    except (FileNotFoundError,KeyError):
        pwd=''
    if not pwd:
        return True
    if st.session_state.get('authenticated'):
        return True
    st.title('Acceso al dashboard WOR')
    value=st.text_input('Contraseña de acceso',type='password')
    if st.button('Ingresar',type='primary'):
        if hmac.compare_digest(value,pwd):
            st.session_state.authenticated=True;st.rerun()
        else:
            st.error('Contraseña incorrecta.')
    st.stop()

require_password()


def style_fig(fig,height=288):
    fig.update_layout(height=height,margin=dict(l=12,r=16,t=22,b=7),
       paper_bgcolor='#FFFFFF',plot_bgcolor='#FFFFFF',font=dict(family='Arial',size=12,color='#38515B'),
       legend=dict(orientation='h',yanchor='bottom',y=1.02,x=0),
       hovermode='x unified')
    fig.update_xaxes(showgrid=False,showline=True,linecolor='#D5E1E4',tickformat='%d/%m')
    fig.update_yaxes(gridcolor='#E7F0F0',zeroline=False)
    return fig


def title_block(title,subtitle):
    st.markdown(f'<div class="head"><h2>{title}</h2><p>{subtitle}</p></div>',unsafe_allow_html=True)


def section(label,color='oil'):
    st.markdown(f'<h3 style="border-left:5px solid {C[color]};padding-left:11px;margin-bottom:6px">{label}</h3>',unsafe_allow_html=True)


def weekly_fig(frame):
    fig=go.Figure()
    for col,label,color in [('petroleo','Oil rate (BOPD)',C['oil']),('entrega','Oil delivery (BOPD)',C['delivery'])]:
        fig.add_trace(go.Scatter(x=frame.fecha,y=frame[col],mode='lines+markers',name=label,
          line=dict(color=color,width=3),marker=dict(size=7)))
    fig.update_yaxes(title='BOPD',tickformat=',.0f')
    return style_fig(fig)


def water_fig(frame,water):
    fig=go.Figure()
    fig.add_trace(go.Scatter(x=frame.fecha,y=frame.agua,mode='lines+markers',name='Produced water',
       line=dict(color=C['water'],width=3)))
    if not water.empty:
        merged=frame[['fecha']].merge(water,on='fecha',how='left')
        fig.add_trace(go.Scatter(x=merged.fecha,y=merged.a10_agua+merged.a11_agua,mode='lines+markers',
           name='Reinjected water',line=dict(color='#397E9B',width=3,dash='dot')))
    fig.update_yaxes(title='BWPD',tickformat=',.0f')
    return style_fig(fig)


def forecast_fig(wor):
    month=wor.end_date.replace(day=1)
    record=wor.forecast[wor.forecast.fecha==month]
    if record.empty or pd.isna(record.iloc[0]['real']) or pd.isna(record.iloc[0]['pronostico']):
        return None,None
    real=float(record.iloc[0]['real']);fore=float(record.iloc[0]['pronostico'])
    fig=go.Figure(go.Bar(x=['Actual','Forecast'],y=[real,fore],
        marker_color=[C['oil'],C['gray']],text=[fmt(real),fmt(fore)],textposition='outside'))
    fig.update_yaxes(title='BOPD',range=[0,max(real,fore)*1.27])
    fig.update_xaxes(tickformat=None)
    return style_fig(fig,260), (real/fore*100 if fore else None)


def gas_fig(gas):
    used=float(gas.combustible.sum());burn=float(gas.quemado.sum())
    if used+burn<=0:
        return None
    fig=go.Figure(go.Pie(labels=['Energy gas','Flared gas'],values=[used,burn],hole=0,
        sort=False,direction='clockwise',marker=dict(colors=[C['gas'],'#78939A']),
        texttemplate='%{percent:.2%}',textposition='outside',
        hovertemplate='%{label}: %{value:,.2f} Mscf / semana (%{percent:.2%})<extra></extra>'))
    fig.update_layout(showlegend=True,margin=dict(l=32,r=32,t=22,b=25))
    return style_fig(fig,290)


def loss_fig(loss, previous, current):
    """Dos columnas verticales: D-1 y D = petróleo real + pérdida roja.

    El rojo es una pérdida REPORTADA, no una tasa inferida del descenso
    observado. Si D no está disponible, el segmento rojo puede dibujarse,
    pero se advierte expresamente que falta el componente de producción real.
    """
    reported = float(loss.recuperable) + float(loss.no_recuperable)
    day_before = (loss.fecha-timedelta(days=1)).strftime('%d/%m')
    event_date = loss.fecha.strftime('%d/%m')
    labels = [f'D-1 · {day_before}', f'D · {event_date}']
    oil = [float(previous) if previous is not None else 0,
           float(current) if current is not None else 0]
    fig = go.Figure()
    fig.add_trace(go.Bar(name='Producción real',x=labels,y=oil,
        marker_color=C['oil'],width=0.42,
        text=[f'{fmt(previous,2)} BOPD' if previous is not None else 'N/D',
              f'{fmt(current,2)} BOPD' if current is not None else 'N/D'],
        textposition='inside',insidetextanchor='middle',textfont=dict(color='white'),
        hovertemplate='%{x}<br>Producción: %{y:,.2f} BOPD<extra></extra>'))
    fig.add_trace(go.Bar(name='Pérdida reportada',x=labels,y=[0,reported],
        marker_color=C['loss'],width=0.42,
        text=['',f'{fmt(reported,2)} BOPD'],textposition='inside',
        insidetextanchor='middle',textfont=dict(color='white',size=12),
        hovertemplate='%{x}<br>Pérdida reportada: %{y:,.2f} BOPD<extra></extra>'))
    fig.update_layout(barmode='stack',height=300,
        margin=dict(l=12,r=12,t=33,b=25),
        legend=dict(orientation='h',x=0,y=1.14),
        plot_bgcolor='white',paper_bgcolor='white',
        font=dict(family='Arial',size=12,color=C['ink']))
    fig.update_yaxes(title='BOPD',rangemode='tozero',gridcolor='#E7F0F0')
    fig.update_xaxes(showgrid=False,type='category')
    if previous is None:
        fig.add_annotation(x=labels[0],y=0,text='N/D',showarrow=False,yshift=14,font=dict(color=C['muted']))
    if current is None:
        fig.add_annotation(x=labels[1],y=reported,text='Petróleo D: N/D',showarrow=False,yshift=14,font=dict(color=C['muted']))
    return fig


@st.cache_data(show_spinner=False)
def parse_cached(raw: bytes):
    if is_model_xlsx(raw):
        result=read_model(raw, require_weekly=True)
        return result.wor, result.daily, 'Modelo WOR'
    return load_wor(raw), None, 'WOR original'


@st.cache_data(show_spinner=False)
def daily_cached(raw: bytes,filename: str):
    if filename.lower().endswith('.xlsx') and is_model_xlsx(raw):
        return read_model(raw, require_weekly=False)
    return load_daily_wells(raw,filename)


st.sidebar.markdown('## 🛢️ WOR • Singue')
st.sidebar.caption('Panel semanal de dos láminas para gerencia')
report_file=st.sidebar.file_uploader('1. Cargar Excel WOR o modelo completo',type=['xlsx'],help='Puedes usar el WOR original sin cambios o rellenar Weekly Summary en el modelo Excel incluido.')
if report_file is None:
    title_block('WEEKLY OPERATIONS DASHBOARD','Carga el Excel WOR desde la barra lateral para generar las dos láminas.')
    st.info('Acepta el WOR original o la plantilla Excel modelo; añade producción diaria por pozo para la comparación de pérdidas. Las cinco capturas se cargan en la barra lateral.')
    c1,c2=st.columns(2)
    with c1:
        st.markdown('#### 01 · Producción y cumplimiento')
        st.write('Indicadores, imágenes de petróleo, agua y Production vs Forecast, y tabla de reinyección.')
    with c2:
        st.markdown('#### 02 · Pérdidas y sistemas')
        st.write('Dos columnas por pozo (D-1 y D con pérdida roja), pastel de gas, imagen RLD y Delivery Sansahuari.')
    st.stop()

raw=report_file.getvalue()
try:
    wor,daily_main,input_type=parse_cached(raw)
except Exception as exc:
    st.error(f'No se pudo leer el formato WOR: {exc}')
    st.stop()

week_losses=wor.losses[(wor.losses.fecha>=wor.start_date)&(wor.losses.fecha<=wor.end_date)].copy()
filename_tag=f'{wor.start_date:%Y%m%d}_{wor.end_date:%Y%m%d}'
source_hash=sha256(raw).hexdigest()[:10]
st.sidebar.success(f'✓ {input_type}: {wor.start_date:%d/%m} - {wor.end_date:%d/%m/%Y}')
st.sidebar.caption(f'{len(wor.production)} días · {len(week_losses)} incidentes en la semana')

with st.sidebar.expander('2. Producción diaria por pozo / datos complementarios', expanded=True):
    st.caption('Si cargas el WOR original, añade el Excel modelo (con la hoja Daily Production llena) o un CSV/XLSX con fecha, pozo y tasa real.')
    daily_file=st.file_uploader('Modelo complementario o producción por pozo',type=['csv','xlsx'],key='daily_file')
    st.markdown('CSV simple: `fecha`, `pozo`, `petroleo_bopd`. Modelo: `Daily Production` con `Date`, `Well`, `Oil Rate BOPD`.')
    template_path=Path(__file__).parent/'plantillas'/'WOR_Modelo_Carga_Excel.xlsx'
    if template_path.exists():
        st.download_button('⬇ Descargar modelo Excel completo', data=template_path.read_bytes(),
            file_name='WOR_Modelo_Carga_Excel.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
    st.download_button('Descargar plantilla CSV simple',data='fecha,pozo,petroleo_bopd\n'.encode('utf-8'),
       file_name='plantilla_produccion_diaria_pozo.csv',mime='text/csv',use_container_width=True)

daily=daily_main
supplement=None
if daily_file is not None:
    try:
        result=daily_cached(daily_file.getvalue(),daily_file.name)
        if hasattr(result,'injection'):
            supplement=result
            if not result.daily.empty:
                daily=result.daily
            if not result.injection.empty:
                wor=replace(wor,injection_records=result.injection)
                st.sidebar.info('Injection Water usa las filas del modelo complementario (promedios por pozo).')
            if not result.losses.empty or not result.gas.empty:
                st.sidebar.caption('Las pérdidas y el gas del WOR original siguen siendo los datos principales; la plantilla complementaria no los reemplaza.')
        else:
            daily=result
        if daily is not None and not daily.empty:
            st.sidebar.success(f'✓ {len(daily)} registros de producción diaria por pozo')
        else:
            st.sidebar.warning('La hoja Daily Production todavía está vacía. Completa producción D-1 y D para comparar.')
    except Exception as exc:
        st.sidebar.error(f'No se pudo leer el archivo complementario: {exc}')


def uploaded_image(label, code):
    """Cada carga corresponde a una sección y a un Excel concreto; no persiste en GitHub."""
    f=st.file_uploader(label,type=['png','jpg','jpeg','webp'],
         key=f'img_{source_hash}_{code}',help='Imagen completa en pantalla y exportación. PNG/JPG/WebP; máximo 12 MB por imagen.')
    if f is None:
        return None
    raw_image=f.getvalue()
    if len(raw_image)>12*1024*1024:
        st.error(f'{label}: la imagen supera los 12 MB.');return None
    try:
        with Image.open(BytesIO(raw_image)) as im:
            if im.width*im.height>25_000_000:
                st.error(f'{label}: demasiados píxeles; reduce la resolución.');return None
            im.verify()
        return raw_image
    except (OSError,ValueError,Image.DecompressionBombError):
        st.error(f'{label}: formato de imagen no válido.');return None

with st.sidebar.expander('3. Imágenes para las láminas (5 espacios)',expanded=True):
    st.caption('Sube cada imagen en su lugar; se ajustan completas, sin recortes, y se incluyen en el PNG/PDF. No son obligatorias para generar el reporte.')
    image_specs=[
        ('oil','01 · Oil Production and Delivery'),
        ('water','02 · Water Produced'),
        ('forecast','03 · Production vs Forecast'),
        ('rld','04 · Artificial Lift / RLD'),
        ('delivery','05 · Delivery Sansahuari'),
    ]
    images={key:uploaded_image(label,key) for key,label in image_specs}
    n_images=sum(v is not None for v in images.values())
    st.caption(f'{n_images} de 5 imágenes cargadas')

with st.sidebar.expander('4. Datos de la tabla de reinyección',expanded=False):
    st.caption('Agua, presión, SST y aceite se obtienen del Excel BSW. Reservorio y salinidad no están en esta tabla del Excel: puedes completarlos aquí con los valores validados del Word.')
    injection_metadata={}
    for k,name in [('a10','Singue A10 WDW'),('a11','Singue A11 WDW')]:
        st.markdown(f'**{name}**')
        injection_metadata[f'{k}_reservoir']=st.text_input(f'Reservorio {name}',placeholder='Ej.: Tiyuyacu',key=f'reservoir_{source_hash}_{k}').strip()
        injection_metadata[f'{k}_salinity']=st.number_input(f'Salinidad {name} (ppm)',min_value=0.0,
           value=None,step=1.0,format='%.0f',placeholder='Pendiente',key=f'salinity_{source_hash}_{k}')

st.sidebar.markdown('---')
st.sidebar.caption('Confidencialidad: el Excel no se publica en GitHub; si alojas la app en un servidor de terceros, los archivos cargados se procesan allí. Configura acceso restringido antes de usar datos sensibles.')

select_loss=None;prev=None;current=None
if not week_losses.empty:
    opts=list(week_losses.index)
    def label_loss(i):
        r=week_losses.loc[i]
        return f"{r.fecha:%d/%m} · {r.pozo} · {fmt(r.recuperable+r.no_recuperable,2)} BOPD"
    st.sidebar.markdown('#### 5. Pérdida para comparar')
    row_idx=st.sidebar.selectbox('Seleccionar incidente',opts,format_func=label_loss)
    select_loss=week_losses.loc[row_idx]
    day_prev,day_curr=match_loss_to_daily(select_loss,daily)
    st.sidebar.caption(f"Comparación contra {select_loss.fecha-timedelta(days=1):%d/%m/%Y} (mismo pozo)")
    daily_hash=sha256(daily_file.getvalue()).hexdigest()[:8] if daily_file is not None else 'principal'
    key=f'{source_hash}_{daily_hash}_{select_loss._excel_row}'
    with st.sidebar.expander('Ingresar / revisar tasas por pozo',expanded=day_prev is None):
        st.caption('Datos reales; si no se conocen, dejar sin llenar. Se preservan durante esta sesión.')
        prev=st.number_input('Producción D-1 del pozo (BOPD)',min_value=0.0,value=day_prev,
            step=.01,format='%.2f',key='prev_'+key)
        current=st.number_input('Producción día D del pozo (BOPD)',min_value=0.0,value=day_curr,
            step=.01,format='%.2f',key='curr_'+key)
    if day_prev is not None and prev is not None and abs(day_prev-prev) > .001:
        st.sidebar.info('Se ha ajustado manualmente el valor D-1 para esta sesión.')

base_prod=wor.production
avg=base_prod[['petroleo','agua','gas','entrega','corte_agua','bsw_entrega']].mean(numeric_only=True)
injection_view=injection_rows(wor,injection_metadata)
water_vals=[r['agua'] for r in injection_view if r['agua'] is not None]
water_avg=sum(water_vals) if water_vals else None
use_fraction=(wor.gas.combustible.sum()/wor.gas.producido.sum()) if not wor.gas.empty and wor.gas.producido.sum()>0 else None

st.markdown('## 🛢️ Weekly Operations · Singue Field')
st.caption(f'Período del Excel: {wor.start_date:%d/%m/%Y} – {wor.end_date:%d/%m/%Y} · Valores reconstruidos desde el libro cargado, sin modificarlo.')
if wor.warnings:
    st.warning(' | '.join(wor.warnings))

tab1,tab2=st.tabs(['01  ·  Producción y cumplimiento','02  ·  Pérdidas y operaciones'])

with tab1:
    title_block('PRODUCTION & PERFORMANCE', 'Evolución semanal, reinyección y cumplimiento mensual — vista de campo')
    c=st.columns(6,gap='small')
    for col,label,val,unit in zip(c,['Petróleo','Agua producida','Gas','Entrega','Water cut','BSW entrega'],
                                   [fmt(avg.petroleo),fmt(avg.agua),fmt(avg.gas),fmt(avg.entrega),
                                    fmt(avg.corte_agua,2)+'%',fmt(avg.bsw_entrega*100,3)+'%'],
                                   ['BOPD','BWPD','MPCD','BOPD','Promedio semanal','Promedio semanal']):
        col.metric(label,val,help=unit);col.caption(unit)
    a,b=st.columns(2,gap='medium')
    with a:
        section('Oil Production and Delivery','oil')
        if images['oil']:
            st.image(images['oil'],caption='Imagen de producción / entregas cargada para este reporte',use_container_width=True)
        else:
            st.info('Espacio reservado: carga la imagen **01 · Oil Production and Delivery** desde la barra lateral.')
    with b:
        section('Water Produced','water')
        if images['water']:
            st.image(images['water'],caption='Imagen de producción de agua cargada para este reporte',use_container_width=True)
        else:
            st.info('Espacio reservado: carga la imagen **02 · Water Produced** desde la barra lateral.')
    a,b=st.columns([1,1],gap='medium')
    with a:
        section('Production vs Forecast','oil')
        if images['forecast']:
            st.image(images['forecast'],caption='Captura del pronóstico y producción real',use_container_width=True)
        else:
            st.info('Carga la imagen **03 · Production vs Forecast** desde la barra lateral.')
        ff,pct=forecast_fig(wor)
        if pct is not None:
            st.caption(f'Cumplimiento desde FORECAST: {fmt(pct,1)} % (compara las cifras con tu imagen adjunta).')
    with b:
        section('Injection Water · tabla del Word','water')
        injections=injection_rows(wor,injection_metadata)
        if injections:
            table=pd.DataFrame(injections).rename(columns={'pozo':'Well','reservorio':'Reservoir','agua':'Water (BWPD)',
                'presion':'Pressure (psi)','sst':'ppm SST','aceite':'ppm oil','salinidad':'ppm Salinity'})
            for k in ['Water (BWPD)','Pressure (psi)','ppm SST','ppm oil','ppm Salinity']:
                table[k]=table[k].apply(lambda x:fmt(x) if x is not None and pd.notna(x) else 'N/D')
            st.dataframe(table,hide_index=True,use_container_width=True)
            st.caption('Promedios semanales por pozo. Reservorio/salinidad provienen del modelo o se completan a mano.')
        else:
            st.info('No se encontró información de reinyección en BSW / Injection Water.')
    with st.expander('Detalle de producción diaria, gas y agua (conservar toda la información)',expanded=False):
        view=base_prod[['fecha','petroleo','agua','fluido','gas','corte_agua','entrega','bsw_entrega']].copy()
        view['bsw_entrega']=view['bsw_entrega']*100
        view.columns=['Date','Oil BOPD','Water BWPD','Total BPD','Gas MPCD','Water cut %','Delivery BOPD','Delivery BSW %']
        st.dataframe(view,hide_index=True,use_container_width=True)
        st.download_button('Descargar tabla diaria CSV',view.to_csv(index=False).encode('utf-8-sig'),
          f'produccion_{filename_tag}.csv',mime='text/csv')

with tab2:
    title_block('LOSSES & OPERATIONS', 'Incidentes, levantamiento artificial, reinyección y gas combustible')
    m=st.columns(4,gap='small')
    rec=week_losses.recuperable.sum();non=week_losses.no_recuperable.sum()
    m[0].metric('Pérdida recuperable',fmt(rec,2)+' BOPD')
    m[1].metric('Pérdida no recuperable',fmt(non,2)+' BOPD')
    m[2].metric('Agua reinyectada',fmt(water_avg)+' BWPD' if water_avg is not None else 'N/D')
    m[3].metric('Gas utilizado',fmt(use_fraction*100,2)+' %' if use_fraction is not None else 'N/D')
    a,b=st.columns([1.55,1],gap='medium')
    with a:
        section('Pérdida reportada vs producción del día anterior','loss')
        if select_loss is not None:
            st.plotly_chart(loss_fig(select_loss,prev,current),use_container_width=True,config={'displayModeBar':False})
            if prev is None or current is None:
                st.warning('Faltan tasas del mismo pozo: carga D-1 y D. La barra roja representa la pérdida reportada; la comparación NO está completa.')
            else:
                observed=prev-current
                st.caption(f'D-1: {fmt(prev,2)} BOPD · D: {fmt(current,2)} BOPD · pérdida reportada: {fmt(select_loss.recuperable+select_loss.no_recuperable,2)} BOPD. Variación observada: {fmt(observed,2)} BOPD; no necesariamente coincide con la pérdida.')
        else:
            st.success('No hay pérdidas registradas en las fechas de esta semana.')
    with b:
        section('Uso del gas producido','gas')
        if not wor.gas.empty:
            chart=gas_fig(wor.gas)
            if chart is not None:
                st.plotly_chart(chart,use_container_width=True,config={'displayModeBar':False})
                total_dist=wor.gas.combustible.sum()+wor.gas.quemado.sum()
                st.caption(f'Energía: {fmt(wor.gas.combustible.sum(),1)} Mscf ({fmt(wor.gas.combustible.sum()/total_dist*100,2)} %) · Quemado: {fmt(wor.gas.quemado.sum(),1)} Mscf ({fmt(wor.gas.quemado.sum()/total_dist*100,2)} %)')
            else:
                st.info('Sin distribución de gas positiva.')
        else:
            st.info('No hay datos de gas en la hoja GAS.')
    a,b=st.columns([1.45,1],gap='medium')
    with a:
        section('Artificial Lift / RLD · imagen adjunta','oil')
        if images['rld']:
            st.image(images['rld'],caption='Imagen RLD cargada por el usuario',use_container_width=True)
        else:
            st.info('Espacio reservado: carga la imagen **04 · Artificial Lift / RLD** desde la barra lateral.')
        if not wor.rld.empty:
            st.caption(f'RLD de la hoja Excel disponible para consulta; último registro {wor.rld.fin_rld.max():%d/%m/%Y}. No se infiere ON/OFF.')
    with b:
        section('Delivery Sansahuari · imagen adjunta','delivery')
        if images['delivery']:
            st.image(images['delivery'],caption='Imagen de entregas en Sansahuari',use_container_width=True)
        else:
            st.info('Carga la imagen **05 · Delivery Sansahuari** desde la barra lateral.')
        if not week_losses.empty:
            st.dataframe(week_losses[['fecha','pozo','recuperable','no_recuperable','comentario']],
                         hide_index=True,use_container_width=True)
    with st.expander('Detalle completo de pérdidas y reinyección',expanded=False):
        st.dataframe(week_losses[['fecha','pozo','recuperable','no_recuperable','comentario']],hide_index=True,use_container_width=True)
        if not wor.water.empty:
            st.dataframe(wor.water.drop(columns=['_excel_row'],errors='ignore'),hide_index=True,use_container_width=True)
        if not wor.gas.empty:
            st.dataframe(wor.gas.drop(columns=['_excel_row'],errors='ignore'),hide_index=True,use_container_width=True)
        st.download_button('Descargar pérdidas de la semana',week_losses.to_csv(index=False).encode('utf-8-sig'),
            f'perdidas_{filename_tag}.csv',mime='text/csv')
    with st.expander('Histórico de pérdidas de la hoja Produccion Losses',expanded=False):
        st.caption('El historial completo no se suma al reporte semanal; solamente sirve para consultas.')
        st.dataframe(wor.losses[['fecha','pozo','recuperable','no_recuperable','comentario']],
                     use_container_width=True,hide_index=True,height=230)

st.markdown('---')
st.markdown('### 📸 Exportar dos láminas listas para el Word')
st.caption('Las cinco capturas (petróleo, agua, pronóstico, RLD y Sansahuari), la tabla de inyección, el pastel de gas y las dos barras de pérdida se exportan en PNG y PDF de dos páginas.')
with st.spinner('Preparando láminas…'):
    png1,png2,pdf=build_exports(wor,week_losses,select_loss,prev,current,images,injection_metadata)
a,b,c=st.columns(3)
a.download_button('⬇ Descargar lámina 1 · PNG',png1,f'WOR_01_Produccion_{filename_tag}.png','image/png',use_container_width=True,type='primary')
b.download_button('⬇ Descargar lámina 2 · PNG',png2,f'WOR_02_Operaciones_{filename_tag}.png','image/png',use_container_width=True,type='primary')
c.download_button('⬇ Descargar reporte de 2 páginas · PDF',pdf,f'WOR_Dashboard_{filename_tag}.pdf','application/pdf',use_container_width=True,type='primary')
with st.expander('Vista previa de las láminas para insertar en Word',expanded=False):
    st.image(png1,caption='Lámina 1: producción y cumplimiento',use_container_width=True)
    st.image(png2,caption='Lámina 2: pérdidas y operación',use_container_width=True)
st.markdown('<div class="smallnote">Validación: el total de pérdidas es una suma de incidencias reportadas en BOPD, no un volumen físico semanal en barriles. No se imputan costos ni producción diaria individual. Los comentarios se mantienen en el documento Word original.</div>',unsafe_allow_html=True)
