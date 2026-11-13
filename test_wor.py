"""Pruebas unitarias sin archivos empresariales en el repositorio."""
import io
import unittest
import zipfile
from datetime import date
import os
import sys
sys.path.insert(0,os.path.dirname(os.path.dirname(__file__)))

from wor_parser import excel_date, load_wor, load_daily_wells, match_loss_to_daily, well_key
from report_export import build_exports, injection_rows
from wor_model import read_model, is_model_xlsx


def fixture():
    name=('Production and Pumping','Produccion Losses','BSW','GAS','FORECAST','RLD')
    ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    def c(ref,value):
        try:
            float(value);return f'<c r="{ref}"><v>{value}</v></c>'
        except (TypeError,ValueError):
            return f'<c r="{ref}" t="inlineStr"><is><t>{value}</t></is></c>'
    def row(i,d):
        return f'<row r="{i}">'+''.join(c(k+str(i),v) for k,v in d.items())+'</row>'
    p=[row(4,{'A':'DATE','B':'OIL RATE (BOPD)'})]
    b=[];g=[]
    for i in range(7):
        e=46318+i
        p.append(row(5+i,{'A':e,'B':4100+i,'C':29000,'D':33100+i,'E':1100,'F':87.5,'G':4000,'H':0.0036}))
        b.append(row(5+i,{'B':e,'C':15000,'D':2000,'E':40,'F':30,'G':14000,'H':1900,'I':40,'J':30}))
        g.append(row(4+i,{'B':e,'C':70,'D':0,'E':70,'F':1100,'G':1030}))
    loss=[row(4,{'B':'DATE','C':'WELL'}),row(5,{'B':46324,'C':'SB-2 "UI"','D':0,'E':129.45,'F':'E.W.O'})]
    forecast=[row(76,{'A':46296,'B':3569.77,'C':3789})]
    rld=[row(4,{'B':'B2','F':1309,'G':1200,'E':46330})]
    blocks=[p,loss,b,g,forecast,rld]
    sheets=''.join(f'<sheet name="{n}" sheetId="{i+1}" r:id="rId{i+1}"/>' for i,n in enumerate(name))
    wb=(f'<workbook xmlns="{ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{sheets}</sheets></workbook>')
    rels='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(
        f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>'
        for i in range(len(name)))+'</Relationships>'
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w',compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr('xl/workbook.xml',wb)
        z.writestr('xl/_rels/workbook.xml.rels',rels)
        for i,block in enumerate(blocks):
            z.writestr(f'xl/worksheets/sheet{i+1}.xml',f'<worksheet xmlns="{ns}"><sheetData>{"".join(block)}</sheetData></worksheet>')
    return out.getvalue()


class WORTests(unittest.TestCase):
    def test_excel_dates(self):
        self.assertEqual(excel_date(46324),date(2026,10,29))
        self.assertIsNone(excel_date(2026))
        self.assertEqual(excel_date('28/10/2026'),date(2026,10,28))
    def test_well_keys(self):
        self.assertEqual(well_key('SB-2 "UI"'),'B2')
        self.assertEqual(well_key('B2'),'B2')
        self.assertEqual(well_key('SB-9 RE "UI"'),'B9RE')
        self.assertNotEqual(well_key('SB-9 RE'),well_key('SB-9'))
        self.assertEqual(well_key('SB-18ST1 UI'),'B18ST1')
    def test_workbook_tables(self):
        w=load_wor(fixture())
        self.assertEqual(len(w.production),7)
        self.assertEqual(w.end_date,date(2026,10,29))
        self.assertAlmostEqual(float(w.losses.no_recuperable.sum()),129.45)
        self.assertAlmostEqual(float(w.water.a10_agua.mean()+w.water.a11_agua.mean()),29000)
        self.assertAlmostEqual(float(w.gas.combustible.sum()/w.gas.producido.sum()),70/1100)
        self.assertTrue(w.warnings)  # RLD fecha posterior al reporte
        self.assertEqual(w.forecast.iloc[0]['real'],3569.77)
    def test_same_well_prior_day_only(self):
        w=load_wor(fixture());event=w.losses.iloc[0]
        daily=load_daily_wells(b'fecha,pozo,petroleo_bopd\n2026-10-28,SB-2 UI,135.5\n2026-10-29,SB-2 UI,0\n2026-10-28,SB-3 UI,999\n','daily.csv')
        prev,now=match_loss_to_daily(event,daily)
        self.assertEqual((prev,now),(135.5,0.0))
        absent=load_daily_wells(b'fecha,pozo,petroleo_bopd\n2026-10-28,SB-3 UI,999\n','daily.csv')
        self.assertEqual(match_loss_to_daily(event,absent),(None,None))
    def test_no_duplicate_series(self):
        with self.assertRaises(ValueError):
            load_daily_wells(b'fecha,pozo,petroleo_bopd\n2026-10-28,B2,135\n2026-10-28,SB-2 UI,145\n','d.csv')
    def test_injection_table_keeps_missing_metadata_uninvented(self):
        w=load_wor(fixture())
        rows=injection_rows(w)
        self.assertEqual(len(rows),2)
        self.assertEqual(rows[0]['reservorio'],'N/D')
        self.assertIsNone(rows[0]['salinidad'])
        self.assertEqual(rows[0]['agua'],15000)
        added=injection_rows(w,{'a10_reservoir':'Tiyuyacu','a10_salinity':12045})
        self.assertEqual(added[0]['salinidad'],12045)
        self.assertIsNone(added[1]['salinidad'])

    def test_all_five_uploaded_images_are_inserted(self):
        from PIL import Image
        w=load_wor(fixture())
        colors={
           'oil':(200,20,20),'water':(20,180,210),'rld':(60,90,210),
           'forecast':(240,160,50),'delivery':(90,180,80),
        }
        images={}
        for name,color in colors.items():
            f=io.BytesIO()
            Image.new('RGB',(800,600),color).save(f,format='PNG')
            images[name]=f.getvalue()
        a,b,c=build_exports(w,w.losses,w.losses.iloc[0],135.5,0,images)
        self.assertEqual(Image.open(io.BytesIO(a)).getpixel((400,445)),colors['oil'])
        self.assertEqual(Image.open(io.BytesIO(a)).getpixel((1110,445)),colors['water'])
        self.assertEqual(Image.open(io.BytesIO(a)).getpixel((430,730)),colors['forecast'])
        self.assertEqual(Image.open(io.BytesIO(b)).getpixel((470,727)),colors['rld'])
        self.assertEqual(Image.open(io.BytesIO(b)).getpixel((1300,727)),colors['delivery'])
        self.assertTrue(c.startswith(b'%PDF'))

    def test_model_empty_is_valid_complement_but_not_main(self):
        from pathlib import Path
        src=Path(__file__).resolve().parents[1] / 'plantillas' / 'WOR_Modelo_Carga_Excel.xlsx'
        content=src.read_bytes()
        self.assertTrue(is_model_xlsx(content))
        result=read_model(content,require_weekly=False)
        self.assertIsNone(result.wor)
        self.assertTrue(result.daily.empty)
        with self.assertRaises(ValueError):
            read_model(content,require_weekly=True)

    def test_model_with_actual_well_values(self):
        from pandas import DataFrame, ExcelWriter
        from pathlib import Path
        from tempfile import NamedTemporaryFile
        from datetime import datetime
        with NamedTemporaryFile(suffix='.xlsx') as tmp:
            tabs={
                'Weekly Summary': DataFrame([
                    ['2026-10-28',4080,29084,33164,1142,87.69,4159,0.36],
                    ['2026-10-29',4041,28258,32299,1121,87.49,4183,0.36]],
                    columns=['Date','Oil Rate BOPD','Water Rate BWPD','Total Fluid BPD','Gas MPCD','Water Cut %','Delivery BOPD','Delivery BSW %']),
                'Daily Production': DataFrame([
                    ['2026-10-28','SB-2 UI',300],['2026-10-29','B2',170]],
                    columns=['Date','Well','Oil Rate BOPD']),
                'Reported Losses': DataFrame([
                    ['2026-10-29','SB-2',0,130,'EWO']],
                    columns=['Date','Well','Recoverable BOPD','Non-Recoverable BOPD','Comments']),
                'Injection Water': DataFrame([
                    ['2026-10-29','Singue A10 WDW','Tiyuyacu',14852,1986,42,33,12045],
                    ['2026-10-29','Singue A11 WDW','Tiyuyacu',14107,1907,42,33,12045]],
                    columns=['Date','Well','Reservoir','Water BWPD','Pressure psi','ppm SST','ppm oil','ppm Salinity']),
                'Gas Distribution': DataFrame([
                    ['2026-10-29',1142,72,1070]],
                    columns=['Date','Produced Gas Mscf','Energy Gas Mscf','Flared Gas Mscf']),
                'Forecast': DataFrame([
                    ['2026-10-01',3570,3789]],columns=['Month','Actual BOPD','Forecast BOPD']),
            }
            with ExcelWriter(tmp.name,engine='openpyxl') as writer:
                for name,data in tabs.items():data.to_excel(writer,sheet_name=name,index=False)
            parsed=read_model(Path(tmp.name).read_bytes())
        self.assertAlmostEqual(parsed.wor.production.iloc[1].bsw_entrega,.0036)
        event=parsed.wor.losses.iloc[0]
        self.assertEqual(match_loss_to_daily(event,parsed.daily),(300.0,170.0))
        injection=injection_rows(parsed.wor)
        self.assertEqual(len(injection),2)
        self.assertEqual(injection[0]['reservorio'],'Tiyuyacu')
        self.assertEqual(injection[0]['salinidad'],12045)
        self.assertEqual(parsed.wor.gas.producido.sum(),1142)

    def test_export_png_and_pdf(self):
        w=load_wor(fixture());loss=w.losses
        p1,p2,pdf=build_exports(w,loss,loss.iloc[0],135.5,0)
        self.assertTrue(p1.startswith(b'\x89PNG'))
        self.assertTrue(p2.startswith(b'\x89PNG'))
        self.assertTrue(pdf.startswith(b'%PDF'))

if __name__=='__main__':unittest.main()
