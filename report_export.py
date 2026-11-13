"""Exportador de dos láminas 16:9 a PNG y PDF, listo para Word o gerencia.

Las láminas se renderizan con el Excel y capturas adjuntadas por el usuario.
Los cinco espacios son: oil, water, forecast, RLD y Delivery Sansahuari.
"""
from __future__ import annotations
from io import BytesIO
from datetime import timedelta
from math import isfinite
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageOps

W, H = 1600, 900
BG = '#F3F7F8'; WHITE = '#FFFFFF'; NAVY = '#172E38'
MUTED = '#61727B'; LINE = '#D7E2E5'; GREEN = '#07806F'
DARKGREEN = '#075D55'; BLUE = '#27B5D3'; YELLOW = '#E9AC26'
RED = '#D7504F'; ORANGE = '#E58D32'; SLATE = '#78939A'
FONT_PATHS = [
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf',
]
BOLD_PATHS = [
    '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
    '/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf',
]


def font(size, bold=False):
    for fp in (BOLD_PATHS if bold else FONT_PATHS):
        if Path(fp).exists():
            return ImageFont.truetype(fp, size)
    return ImageFont.load_default()


def fmt(value, digits=0):
    if value is None or pd.isna(value):
        return 'N/D'
    s = f'{float(value):,.{digits}f}'
    return s.replace(',', '@').replace('.', ',').replace('@', '.')


def text(draw, pos, value, sz=20, color=NAVY, bold=False, anchor=None):
    draw.text(pos, str(value), fill=color, font=font(sz, bold), anchor=anchor)


def rounded(draw, bbox, fill=WHITE, radius=22, outline=None, width=1):
    draw.rounded_rectangle(bbox, radius=radius, fill=fill, outline=outline, width=width)


def header(draw, number, title, subtitle, date_label):
    draw.rectangle((0,0,W,12), fill=GREEN)
    text(draw,(58,35), 'SINGUE  /  WEEKLY OPERATIONS', 23, GREEN, True)
    text(draw,(58,76), title, 40, NAVY, True)
    text(draw,(1540,53), date_label, 22, MUTED, anchor='ra')
    rounded(draw,(1356,83,1542,119), '#E4EEE9', radius=16)
    text(draw,(1449,101), f'PLATE  {number} / 2',17, DARKGREEN,True, anchor='mm')
    draw.line((58,867,1542,867), fill=LINE, width=2)
    text(draw,(58,878), 'Fuente: hojas del WOR Excel | Sin comentarios narrativos',14,MUTED)
    text(draw,(1540,878), 'OPERATIONS  •  INTERNAL',14,MUTED,anchor='ra')


def metric(draw, x, y, width, label, value, unit, color):
    rounded(draw,(x,y,x+width,y+121),WHITE,17)
    draw.rounded_rectangle((x+17,y+21,x+24,y+56), radius=3,fill=color)
    text(draw,(x+37,y+20), label.upper(),15,MUTED,True)
    text(draw,(x+20,y+53), value,33,NAVY,True)
    text(draw,(x+width-17,y+92),unit,15,color,True,anchor='ra')


def panel(draw,bbox,title,tag='',accent=GREEN):
    rounded(draw,bbox,WHITE,19)
    x1,y1,x2,y2=bbox
    draw.rounded_rectangle((x1+18,y1+19,x1+26,y1+46),radius=3,fill=accent)
    text(draw,(x1+40,y1+18),title,22,NAVY,True)
    if tag:
        text(draw,(x2-19,y1+23),tag,16,MUTED,anchor='ra')


def series_plot(draw,box,dates,data,colors,labels,axis_min=None,axis_max=None):
    x1,y1,x2,y2=box
    left=x1+75; right=x2-30; top=y1+70; bottom=y2-62
    vals=[float(x) for serie in data for x in serie if x is not None and isfinite(float(x))]
    if not vals:
        text(draw,(left,top+40),'NO DATA AVAILABLE',18,MUTED,True)
        return
    vmin=min(vals) if axis_min is None else axis_min
    vmax=max(vals) if axis_max is None else axis_max
    gap=max((vmax-vmin)*0.22,1)
    vmin-=gap; vmax+=gap
    for t in range(4):
        y=top+(bottom-top)*t/3
        draw.line((left,y,right,y),fill='#E6ECEE',width=2)
        text(draw,(left-9,y),fmt(vmax-(vmax-vmin)*t/3),14,MUTED,anchor='rm')
    for si, serie in enumerate(data):
        points=[]
        for i,v in enumerate(serie):
            if v is None or pd.isna(v):continue
            px=left+(right-left)*i/max(len(dates)-1,1)
            py=bottom-(float(v)-vmin)/(vmax-vmin)*(bottom-top)
            points.append((px,py))
        if len(points)>1:draw.line(points,fill=colors[si],width=5,joint='curve')
        for px,py in points:draw.ellipse((px-4,py-4,px+4,py+4),fill=colors[si])
    for i,d in enumerate(dates):
        text(draw,(left+(right-left)*i/max(len(dates)-1,1),bottom+10),d.strftime('%d/%m'),13,MUTED,anchor='mt')
    cur=left
    for label, col in zip(labels,colors):
        draw.rounded_rectangle((cur,y2-25,cur+17,y2-10),radius=3,fill=col)
        text(draw,(cur+24,y2-27),label,15,MUTED)
        cur+=max(150,len(label)*10+60)


def image_panel(draw, img_bytes, box, title, *, accent=SLATE, empty_label='ADJUNTAR CAPTURA EN STREAMLIT', tag='IMAGEN PERSONALIZADA'):
    """Inserta la imagen completa SIN RECORTAR. Etiquetas visibles incluso sin captura."""
    x1, y1, x2, y2 = box
    panel(draw, box, title, tag, accent)
    inset = (int(x1+19), int(y1+57), int(x2-19), int(y2-16))
    if not img_bytes:
        rounded(draw, inset, '#F0F6F7', 13, LINE)
        cx, cy = (inset[0]+inset[2])/2, (inset[1]+inset[3])/2
        text(draw, (cx,cy-14), 'IMAGEN PENDIENTE', 19, MUTED, True, anchor='mm')
        text(draw, (cx,cy+19), empty_label, 14, SLATE, anchor='mm')
        return
    try:
        with Image.open(BytesIO(img_bytes)) as im:
            pic = ImageOps.exif_transpose(im).convert('RGB')
        target = (inset[2]-inset[0],inset[3]-inset[1])
        pic.thumbnail(target, Image.Resampling.LANCZOS)
        dest_x = inset[0]+(target[0]-pic.width)//2
        dest_y = inset[1]+(target[1]-pic.height)//2
        draw._image.paste(pic,(dest_x,dest_y))
    except (OSError, ValueError, Image.DecompressionBombError):
        text(draw,(x1+33,y1+90),'No se pudo decodificar la imagen',18,RED)


def injection_rows(wor, metadata=None):
    """Promedio semanal por inyector; campos ausentes del Excel quedan pendientes.

    metadata: {a10_reservoir,a11_reservoir,a10_salinity,a11_salinity}.
    La salinidad y el reservorio NO aparecen en la lectura de la hoja BSW.
    """
    metadata = metadata or {}
    extras=getattr(wor,'injection_records',None)
    if extras is not None and not extras.empty:
        rows=[]
        for well_key_id, chunk in extras.groupby('pozo_id',sort=False):
            def avg(col):
                return float(chunk[col].mean()) if chunk[col].notna().any() else None
            reservoir=next((str(x) for x in chunk['reservorio'].dropna() if str(x).strip()),'N/D')
            rows.append({'pozo':chunk.iloc[0]['pozo'], 'reservorio':reservoir,
                'agua':avg('agua'),'presion':avg('presion'),'sst':avg('sst'),
                'aceite':avg('aceite'),'salinidad':avg('salinidad')})
        return rows
    if wor.water.empty:
        return []
    frame=wor.water
    rows=[]
    for alias,well in [('a10','Singue A10 WDW'),('a11','Singue A11 WDW')]:
        def mean(attr):
            col=f'{alias}_{attr}'
            return float(frame[col].mean()) if col in frame and frame[col].notna().any() else None
        rows.append({
            'pozo':well,
            'reservorio':metadata.get(f'{alias}_reservoir') or 'N/D',
            'agua':mean('agua'), 'presion':mean('presion'),
            'sst':mean('sst'), 'aceite':mean('aceite'),
            'salinidad':metadata.get(f'{alias}_salinity'),
        })
    return rows


def injection_table(draw, box, wor, metadata=None):
    panel(draw,box,'Injection water','PROMEDIO SEMANAL',BLUE)
    x1,y1,x2,y2=box
    rows=injection_rows(wor,metadata)
    if not rows:
        text(draw,(x1+28,y1+103),'No hay registros en BSW',18,MUTED)
        return
    # Diseño compacto: 7 columnas de la tabla original de Word.
    offsets=[20,140,233,325,420,503,575]
    columns=['WELL','RESERVOIR','WATER','PSI','SST','OIL','SALINITY']
    widths=[119,93,92,95,83,72,94]
    dx=x1
    maxwidth=x2-x1
    ratio=(maxwidth-40)/sum(widths)
    acc=x1+19
    starts=[]
    for wd in widths:
        starts.append(acc);acc+=wd*ratio
    hy=y1+79
    draw.rounded_rectangle((x1+14,hy-16,x2-14,hy+20),radius=8,fill='#EBF5F5')
    for name,xx in zip(columns,starts):
        text(draw,(xx,hy-6),name,13,DARKGREEN,True)
    for i,record in enumerate(rows):
        y=hy+37+i*46
        fields=[record['pozo'].replace('Singue ',''),str(record['reservorio']),
                fmt(record['agua']),fmt(record['presion']),fmt(record['sst']),
                fmt(record['aceite']),fmt(record['salinidad'])]
        for k,(value,xx) in enumerate(zip(fields,starts)):
            f=14 if k in (0,1) else 15
            text(draw,(xx,y),value,f,NAVY, k==0)
        draw.line((x1+16,y+31,x2-16,y+31),fill=LINE,width=1)
    text(draw,(x1+22,y2-27),'Promedios BSW • N/D = dato no proporcionado',13,MUTED)


def gas_distribution_pie(draw,box,gas):
    """Pastel 2 categorías con porcentajes exactos calculados del Excel."""
    x1,y1,x2,y2=box
    panel(draw,box,'Produced gas distribution','WEEKLY / Mscf',YELLOW)
    if gas.empty:
        text(draw,(x1+22,y1+120),'No hay datos en GAS',20,MUTED)
        return
    energy=float(gas.combustible.sum());flared=float(gas.quemado.sum())
    amount=energy+flared
    if amount<=0:
        text(draw,(x1+22,y1+120),'Sin volúmenes positivos',20,MUTED)
        return
    pct_energy=energy/amount*100;pct_flared=flared/amount*100
    center=(x1+145,y1+161);radius=87
    circle=(center[0]-radius,center[1]-radius,center[0]+radius,center[1]+radius)
    # Crea un pastel completo y real con áreas proporcionales.
    draw.pieslice(circle,start=-90,end=270,fill='#78939A')
    if energy>0:
        draw.pieslice(circle,start=-90,end=-90+360*energy/amount,fill=YELLOW)
    legendx=x1+269
    for i,(lab,pct,val,col) in enumerate([
        ('Gas for energy',pct_energy,energy,YELLOW),
        ('Flared gas',pct_flared,flared,SLATE)]):
        y=y1+96+i*92
        draw.rounded_rectangle((legendx,y+8,legendx+15,y+24),radius=4,fill=col)
        text(draw,(legendx+27,y),lab,18,MUTED,True)
        text(draw,(legendx+27,y+29),fmt(pct,2)+' %',23,NAVY,True)
        text(draw,(legendx+150,y+33),fmt(val,1)+' Mscf',14,MUTED)


def draw_page1(wor, images=None, injection_metadata=None):
    """Lámina 1: tres imágenes y tabla de reinyección sin OFM Evidence."""
    images=images or {}
    img=Image.new('RGB',(W,H),BG); d=ImageDraw.Draw(img);d._image=img
    p=wor.production
    dates=p['fecha'].tolist()
    header(d,1,'FIELD PRODUCTION & PERFORMANCE','',f'{dates[0]:%d/%m/%Y} - {dates[-1]:%d/%m/%Y}')
    defs=[('Oil rate',fmt(p.petroleo.mean()),'BOPD',GREEN),
          ('Water produced',fmt(p.agua.mean()),'BWPD',BLUE),
          ('Gas produced',fmt(p.gas.mean()),'MPCD',YELLOW),
          ('Oil delivery',fmt(p.entrega.mean()),'BOPD',DARKGREEN),
          ('Water cut',fmt(p.corte_agua.mean(),2)+'%','Field',BLUE),
          ('BSW delivery',fmt(p.bsw_entrega.mean()*100,3)+'%','Delivery',SLATE)]
    w=(1484-5*14)//6
    for i,vals in enumerate(defs):metric(d,58+i*(w+14),138,w,*vals)
    image_panel(d,images.get('oil'),(58,277,789,556),'Oil production and delivery',accent=GREEN,tag='USER IMAGE')
    image_panel(d,images.get('water'),(810,277,1542,556),'Water produced',accent=BLUE,tag='USER IMAGE')
    image_panel(d,images.get('forecast'),(58,575,789,846),'Production vs forecast',accent=GREEN,tag='USER IMAGE')
    injection_table(d,(810,575,1542,846),wor,injection_metadata)
    return img


def draw_page2(wor, losses, selected_loss=None, previous=None, current=None, images=None):
    images=images or {}
    img=Image.new('RGB',(W,H),BG);d=ImageDraw.Draw(img);d._image=img
    dates=wor.production['fecha'].tolist()
    header(d,2,'LOSSES & OPERATING SYSTEMS','',f'{dates[0]:%d/%m/%Y} - {dates[-1]:%d/%m/%Y}')
    recoverable=float(losses.recuperable.sum()) if not losses.empty else 0
    nonrecoverable=float(losses.no_recuperable.sum()) if not losses.empty else 0
    produced=wor.gas.producido.sum() if not wor.gas.empty else 0
    energy=wor.gas.combustible.sum() if not wor.gas.empty else 0
    use=(energy/produced*100) if produced else None
    inj=(wor.water.a10_agua+wor.water.a11_agua).mean() if not wor.water.empty else None
    defs=[('Recoverable loss',fmt(recoverable,2),'BOPD reported',ORANGE),
          ('Non-recoverable loss',fmt(nonrecoverable,2),'BOPD reported',RED),
          ('Water reinjection',fmt(inj),'BWPD average',BLUE),
          ('Gas to energy',fmt(use,2)+'%' if use is not None else 'N/D','Produced gas',YELLOW)]
    cw=(1484-3*16)//4
    for i,vals in enumerate(defs):metric(d,58+i*(cw+16),138,cw,*vals)
    boxa=(58,283,917,558)
    panel(d,boxa,'Reported loss vs previous day','SAME WELL / BOPD',RED)
    if selected_loss is not None:
        day=selected_loss['fecha']; well=selected_loss['pozo']
        loss=max(0,float(selected_loss['recuperable'])+float(selected_loss['no_recuperable']))
        # EXACTAMENTE DOS barras verticales: D-1 real y D real + pérdida reportada roja.
        # No se dibuja un valor ficticio de petróleo si D-1 o D son desconocidos.
        prev=float(previous) if previous is not None else None
        curr=float(current) if current is not None else None
        highest=max((prev or 0),(curr or 0)+loss,1)*1.13
        chart_top=boxa[1]+91
        bottom=502
        height=bottom-chart_top
        text(d,(82,347),f'{well}  ·  pérdida reportada: {fmt(loss,2)} BOPD',17,MUTED,True)
        for x,rate,title in [(330,prev,f'D-1  {(day-timedelta(days=1)):%d/%m}'),
                             (640,curr,f'D  {day:%d/%m}')]:
            d.line((x-82,bottom,x+82,bottom),fill=LINE,width=2)
            if rate is None:
                d.rounded_rectangle((x-52,bottom-78,x+52,bottom),radius=7,outline='#A3B6B9',width=2)
                if x==330:
                    text(d,(x,bottom-39),'N/D',18,MUTED,True,anchor='mm')
            else:
                oil_height=int(max(0,rate)/highest*height)
                if oil_height:
                    d.rectangle((x-52,bottom-oil_height,x+52,bottom),fill=GREEN)
                if x==640:
                    lost_h=int(loss/highest*height)
                    if lost_h:
                        d.rectangle((x-52,bottom-oil_height-lost_h,x+52,bottom-oil_height),fill=RED)
                        text(d,(x,bottom-oil_height-lost_h/2),fmt(loss,2),14,WHITE,True,anchor='mm')
                if x==640 and oil_height>=25:
                    text(d,(x,bottom-oil_height/2),fmt(rate,2),15,WHITE,True,anchor='mm')
                elif x==330:
                    text(d,(x,bottom-oil_height-8),fmt(rate,2),16,NAVY,True,anchor='mb')
                else:
                    text(d,(x+65,bottom-4),'Real: '+fmt(rate,2),13,GREEN,True,anchor='lm')
            text(d,(x,513),title,16,MUTED,True,anchor='mt')
        if curr is None and loss>0:
            # Falta producción real del evento: rojo reportado como dato aislado,
            # no como total D reconstruido.
            red_h=int(loss/highest*height)
            d.rectangle((640-52,bottom-red_h,640+52,bottom),fill=RED)
            text(d,(640,bottom-red_h/2),fmt(loss,2),14,WHITE,True,anchor='mm')
        if prev is None or curr is None:
            text(d,(87,547),'DATOS D-1/D INCOMPLETOS  ·  rojo = pérdida reportada',14,ORANGE,True)
        else:
            text(d,(87,547),'Rojo = pérdida reportada, NO descenso calculado entre días',13,MUTED)
    else:
        text(d,(92,430),'No reported losses during this week',23,MUTED,True)
    gas_distribution_pie(d,(937,283,1542,558),wor.gas)
    image_panel(d,images.get('rld'),(58,577,917,846),'Artificial lift / RLD',accent=GREEN,tag='USER IMAGE')
    image_panel(d,images.get('delivery'),(937,577,1542,846),'Delivery Sansahuari',accent=SLATE,tag='USER IMAGE')
    return img

def to_png(img):
    f=BytesIO();img.save(f,format='PNG',optimize=True);return f.getvalue()


def to_pdf(page1,page2):
    f=BytesIO()
    page1.save(f,format='PDF',save_all=True,append_images=[page2],resolution=135.0,title='Weekly Operations - Two Plate Summary',author='WOR Streamlit Dashboard')
    return f.getvalue()


def build_exports(wor, losses, selected_loss=None, previous=None, current=None, images=None, injection_metadata=None):
    """Genera dos láminas: petróleo, agua, pronóstico, RLD y Sansahuari."""
    p1=draw_page1(wor,images,injection_metadata)
    p2=draw_page2(wor,losses,selected_loss,previous,current,images)
    return to_png(p1),to_png(p2),to_pdf(p1,p2)
