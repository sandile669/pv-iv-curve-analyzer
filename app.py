import io
import re
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from analysis_engine import parse_csv, string_number, analyze
from reporting import word_report
from raw_metadata import parse_raw_metadata

st.set_page_config(page_title='PV IV Curve Analyzer',page_icon='☀️',layout='wide')
st.title('☀️ PV String IV Curve Analyzer')
st.caption('Version 1.7 • Inverter-level analysis • CSV → IV/PV plots → Word')
with st.sidebar:
    st.header('Report settings')
    project=st.text_input('Project','Virginia Solar Project')
    block=st.text_input('Block','2')
    inverter=st.text_input('Inverter','12')
    expected=st.number_input('Expected strings',min_value=1,max_value=18,value=18,help='The supplied report template has 18 string graph positions.')
    inverter_serial=st.text_input('Inverter Serial No.','',help='Entered into Test Information table in the Word report.')
    test_date=st.date_input('Test date').isoformat()
    with st.expander('Test equipment details'):
        equipment=st.text_input('Test equipment','MULTI-TESTER (IV-CURVE)')
        equipment_serial=st.text_input('Equipment serial number','BOB01890')
        calibration_certificate=st.text_input('Calibration certificate number','129647')
        calibration_expiry=st.text_input('Calibration expiry date','18 August 2027')
    cut_text=st.text_input('Cut string numbers (comma separated)','1,2,3')
    threshold=st.number_input('Abnormal if FF below (%)',value=75.0,min_value=0.0,max_value=100.0,step=0.5)
    st.info('The first version uses an inverter-specific upload. Do not mix files from different inverters.')
raw_file=st.file_uploader('Raw Data Files (CSV or Excel, measured + STC metrics)',type=['csv','xlsx','xls'],help='Upload the tester summary export containing File Name and measured/STC fields.')
metadata={}
if raw_file:
    try:
        metadata=parse_raw_metadata(raw_file.getvalue(),raw_file.name)
        st.success(f'Raw Data Files imported for {len(metadata)} strings')
        preview=[]
        for number,entry in sorted(metadata.items()):
            item={'String':number}
            for field,value in entry['stc'].items():item[field]=value
            preview.append(item)
        st.subheader('STC Data Preview — values read from Raw Data Files')
        st.dataframe(pd.DataFrame(preview),hide_index=True,use_container_width=True)
        missing=[str(number) for number,entry in sorted(metadata.items()) if not entry['stc']]
        if missing:st.warning('No STC fields detected for strings: '+', '.join(missing)+'. The report will not invent missing STC values. Please share the original export for exact mapping.')
    except Exception as exc: st.error(f'Raw Data Files import: {exc}')
files=st.file_uploader('Upload CSV file(s), one per tested string',type=['csv'],accept_multiple_files=True)
try:
    cut={int(x.strip()) for x in cut_text.split(',') if x.strip()}
except ValueError:
    st.error('Cut string list must contain integers separated by commas.'); st.stop()
if not files:
    st.info('Upload files named like b2-inv12-s4.csv. For a quick trial, use the included demo CSV files.'); st.stop()
records={}; warnings=[]
for f in files:
    s=string_number(f.name)
    if s is None:
        warnings.append(f'{f.name}: could not identify string number from filename'); continue
    if s < 1 or s > expected:
        warnings.append(f'{f.name}: string {s} outside configured range'); continue
    if s in cut:
        warnings.append(f'{f.name}: file ignored because string {s} is marked cut'); continue
    if s in records:
        warnings.append(f'{f.name}: duplicate string {s}; first file retained'); continue
    try:
        d=parse_csv(f.getvalue()); result=analyze(d)
        result['Result']='Abnormal' if pd.notna(result['FF (%)']) and result['FF (%)'] < threshold else result['Result']
        if result['Result'] == 'Abnormal' and result['FF (%)'] >= threshold: result['Result']='Normal'
        # Prefer certified tester summary measurements when explicitly provided.
        result.update(metadata.get(s,{}).get('measured',{}))
        if pd.notna(result.get('FF (%)')) and 0 <= result['FF (%)'] <= 100:
            result['Result']='Abnormal' if result['FF (%)'] < threshold else 'Normal'
        records[s]=(d,result)
    except Exception as exc: warnings.append(f'{f.name}: {exc}')
for w in warnings: st.warning(w)
if not records:
    st.error('No valid IV curves were imported. Check filenames and CSV headers.'); st.stop()
rows=[]
for s in range(1,expected+1):
    if s in cut: r={'String':s,'Result':'Strings Cut'}
    elif s not in records: r={'String':s,'Result':'Missing / Not Tested'}
    else: r={'String':s,**records[s][1]}
    rows.append(r)
normal=sum(r['Result']=='Normal' for r in rows); abnormal=sum(r['Result']=='Abnormal' for r in rows)
a,b,c,d=st.columns(4); a.metric('Expected',expected); b.metric('Tested',len(records)); c.metric('Normal',normal); d.metric('Abnormal',abnormal)
st.subheader('String test results')
st.dataframe(pd.DataFrame(rows),hide_index=True,use_container_width=True)
st.subheader('IV / PV curves')
figures={}
for s,(data,res) in sorted(records.items()):
    fig,ax=plt.subplots(figsize=(9,5)); v=data.V.to_numpy(); i=data.I.to_numpy(); p=v*i
    ax.plot(v,i,color='#1673b8',linewidth=2,label='Current (A)'); ax.set_xlabel('Voltage (V)'); ax.set_ylabel('Current (A)',color='#1673b8'); ax.set_ylim(bottom=0)
    ax2=ax.twinx(); ax2.plot(v,p,color='#ed8c25',linestyle='--',linewidth=2,label='Power (W)'); ax2.scatter([res['Vm (V)']],[res['Pmax (W)']],color='red',s=40,zorder=5,label='MPP'); ax2.set_ylabel('Power (W)',color='#ed8c25'); ax2.set_ylim(bottom=0)
    ax.set_title(f'Block {block} / Inverter {inverter} / String {s} — {res["Result"]} | FF {res["FF (%)"]:.2f}%')
    ax.grid(alpha=.25); h1,l1=ax.get_legend_handles_labels(); h2,l2=ax2.get_legend_handles_labels(); ax.legend(h1+h2,l1+l2,loc='upper right'); fig.tight_layout()
    buf=io.BytesIO(); fig.savefig(buf,format='png',dpi=145,bbox_inches='tight'); figures[s]=buf.getvalue()
    with st.expander(f'String {s} — {res["Result"]} (FF {res["FF (%)"]:.2f}%)',expanded=(len(records)<=3)):
        st.pyplot(fig)
        if res.get('Warnings'): st.warning(res['Warnings'])
    plt.close(fig)
st.subheader('STC data validation')
stc_count=sum(bool(v.get('stc')) for v in metadata.values())
st.caption(f'STC data found for {stc_count} strings. Missing STC values are left blank; no STC corrections are fabricated.')
st.subheader('Generate reports')
name=f'Block_{block}_Inverter_{inverter}_IV_Curve_Report'
report_kwargs=dict(stc_by_string={s:v['stc'] for s,v in metadata.items()},inverter_serial=inverter_serial,test_date=test_date,equipment=equipment,equipment_serial=equipment_serial,calibration_certificate=calibration_certificate,calibration_expiry=calibration_expiry)
st.download_button('⬇ Download Word report (.docx)',word_report(project,block,inverter,rows,figures,**report_kwargs),file_name=name+'.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',use_container_width=True)
st.caption('Engineering caution: The app computes Pmax from sampled V×I values and estimates Isc/Voc when intercepts are absent. Verify against tester-certified results before acceptance.')
