"""Template-preserving Word reports (.docx) only."""
import io
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import RGBColor

TEMPLATE_DIR = Path(__file__).resolve().parent / 'templates'


def _replace_paragraph(p, old, new):
    if old not in p.text: return
    for run in p.runs:
        if old in run.text:
            run.text=run.text.replace(old,new); return
    # fragmented text: replace in first run while preserving paragraph formatting
    if p.runs:
        p.runs[0].text=p.text.replace(old,new)
        for run in p.runs[1:]: run.text=''


def _cell_text(cell, value):
    # keep original cell formatting by editing first run rather than assigning cell.text
    p=cell.paragraphs[0]
    if p.runs:
        p.runs[0].text=str(value)
        for r in p.runs[1:]: r.text=''
    else: p.add_run(str(value))
    for extra in cell.paragraphs[1:]:
        for r in extra.runs: r.text=''


def _num(value, decimals=2):
    try:
        import math
        v=float(value)
        return f'{v:.{decimals}f}' if math.isfinite(v) else ''
    except (ValueError, TypeError): return ''


def _clear_cell_picture(cell):
    for drawing in cell._element.xpath('.//w:drawing'):
        drawing.getparent().remove(drawing)
    for pict in cell._element.xpath('.//w:pict'):
        pict.getparent().remove(pict)


def _set_result_shading(cell, result):
    """Set actual Word cell shading, overriding template theme colors."""
    colors = {'Normal': 'E2F0D9', 'Abnormal': 'F2DBDB'}
    fill = colors.get(result)
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in list(tc_pr):
        if old.tag == qn('w:shd'):
            tc_pr.remove(old)
    if fill:
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), fill)
        tc_pr.append(shd)


def word_report(project, block, inverter, rows, figures, stc_by_string=None, inverter_serial='', equipment='MULTI-TESTER (IV-CURVE)', equipment_serial='BOB01890', calibration_certificate='129647', calibration_expiry='18 August 2027', test_date=''):
    max_string = len(rows)
    if max_string not in (15, 18, 21) or {int(r['String']) for r in rows} != set(range(1, max_string+1)):
        raise ValueError('Expected exactly 15, 18 or 21 sequential string records, including cut/missing strings.')
    template = TEMPLATE_DIR / f'IV_Curve_Template_{max_string}.docx'
    if not template.is_file():
        raise FileNotFoundError(f'Missing template: {template.name}; upload the templates folder to GitHub.')
    doc=Document(template)
    stc_by_string=stc_by_string or {}
    # Keep original cover page, approval table, header, footer, logos and page breaks.
    for p in doc.paragraphs:
        if max_string == 15: _replace_paragraph(p, "Strings 13–18", "Strings 13–15")
        for old,new in [('Block 9 - Inverter 28',f'Block {block} - Inverter {inverter}'),
                        ('Block 9 · Inverter 28',f'Block {block} · Inverter {inverter}'),
                        ('Block 9 · Inverter 28·',f'Block {block} · Inverter {inverter} ·'),
                        ('Virginia Solar Park 240MW',project),
                        ('2026-10-22',test_date or '')]:
            _replace_paragraph(p,old,new)
    info=doc.tables[1]
    values={(0,1):f'BLOCK {block}',(1,1):f'INV {inverter}',(2,1):inverter_serial,
            (3,1):str(len(rows)),(0,3):equipment,(1,3):equipment_serial,
            (2,3):calibration_certificate,(3,3):calibration_expiry}
    for (r,c),v in values.items(): _cell_text(info.cell(r,c),v)
    # Traverse existing graph cells in row-major order, including the two-row final table in the 21-string template.
    summary = doc.tables[-1]
    slots = [cell for table in doc.tables[2:-1] for row in table.rows for cell in row.cells]
    if len(summary.rows) != 2*max_string+1 or len(slots) != max_string + max_string%2:
        raise ValueError('Template layout does not match selected string count.')
    by_num = {int(r['String']): r for r in rows}
    for s, cell in enumerate(slots, 1):
        drawings = cell._element.xpath('.//wp:extent')
        width = int(drawings[0].get('cx')) if drawings else Inches(3.15)
        height = int(drawings[0].get('cy')) if drawings else Inches(2.07)
        _clear_cell_picture(cell)
        for p in cell.paragraphs:
            for run in p.runs: run.text=''
        if s > max_string: continue
        if s in figures and by_num[s]['Result'] not in ('Strings Cut', 'Missing / Not Tested'):
            cell.paragraphs[0].add_run().add_picture(io.BytesIO(figures[s]),width=width,height=height)
        else:
            cell.paragraphs[0].add_run(f"String {s}\n{by_num[s]['Result']}")
    keys=['Pmax (W)','Voc (V)','Isc (A)','Vm (V)','Im (A)','Irradiance (W/m²)','PV Temp (°C)','En Temp (°C)','Incline','FF (%)']
    by_num={int(r['String']):r for r in rows}
    for s in range(1,max_string+1):
        r=by_num.get(s,{'Result':'Not Applicable'})
        row_idx=2*s-1
        cells=summary.rows[row_idx].cells
        _cell_text(cells[0],f'String {s}')
        for j,k in enumerate(keys,1): _cell_text(cells[j],_num(r.get(k),3 if k=='FF (%)' else 2))
        result=r.get('Result','Not Applicable')
        _cell_text(cells[11],result)
        _set_result_shading(cells[11], result)
        # Original summary has paired STC rows; never present original inverter's STC values as new measurements.
        stc=summary.rows[row_idx+1].cells
        _cell_text(stc[0],'STC')
        stc_values=stc_by_string.get(s,{}) if result not in ('Strings Cut','Missing / Not Tested') else {}
        for j,k in enumerate(keys,1): _cell_text(stc[j],_num(stc_values.get(k),3 if k=='FF (%)' else 2))
        # In merged Result cells, retain the classification and background styling.
        for run in cells[11].paragraphs[0].runs:
            if result=='Abnormal': run.font.color.rgb=RGBColor(190,0,0)
            elif result=='Normal': run.font.color.rgb=RGBColor(0,120,0)
    out=io.BytesIO();doc.save(out);return out.getvalue()

