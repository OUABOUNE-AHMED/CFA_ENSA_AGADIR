import pandas as pd
import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import seaborn as sns
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
from matplotlib.pyplot import figure
figure(figsize=(10, 10), dpi=80)
from IPython.display import set_matplotlib_formats
set_matplotlib_formats('svg')
import base64

def app():
    def to_excel_M(df1,df2,df3,df4,df5,df6,df7,df8):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df1.to_excel(writer, index=False, sheet_name='Single Tenants Assets')
        df2.to_excel(writer, index=False, sheet_name='Shopping centers')
        df3.to_excel(writer, index=False, sheet_name='Commercial Galleries')
        df4.to_excel(writer, index=False, sheet_name='Industrial Unit')
        df5.to_excel(writer, index=False, sheet_name='Assets total revenue')
        df6.to_excel(writer, index=False, sheet_name='Total Revenues')
        df7.to_excel(writer, index=False, sheet_name='Net Income')
        df8.to_excel(writer, index=False, sheet_name='FFO')


        workbook = writer.book
        worksheet1 = writer.sheets['Single Tenants Assets']
        worksheet2 = writer.sheets['Shopping centers']
        worksheet3 = writer.sheets['Commercial Galleries']
        worksheet4 = writer.sheets['Industrial Unit']
        worksheet5 = writer.sheets['Assets total revenue']
        worksheet6 = writer.sheets['Total Revenues']
        worksheet7 = writer.sheets['Net Income']
        worksheet8 = writer.sheets['FFO']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet1.set_column('A:A', None, format1)
        worksheet2.set_column('A:A', None, format1)
        worksheet3.set_column('A:A', None, format1)
        worksheet4.set_column('A:A', None, format1)
        worksheet5.set_column('A:A', None, format1)
        worksheet6.set_column('A:A', None, format1)
        worksheet7.set_column('A:A', None, format1)
        worksheet8.set_column('A:A', None, format1)
        writer.save()
        processed_data = output.getvalue()
        return processed_data  

    df = pd.read_excel ('data.xlsx')
    ST = pd.read_excel('data.xlsx',sheet_name ='ST').dropna()
    SHC = pd.read_excel('data.xlsx',sheet_name ='SHC').dropna()
    CG = pd.read_excel('data.xlsx',sheet_name ='CG').dropna()
    IU = pd.read_excel('data.xlsx',sheet_name ='IU').dropna()
    TR1 = pd.read_excel('data.xlsx',sheet_name ='TR1').dropna()
    TR2 = pd.read_excel('data.xlsx',sheet_name ='TR2').dropna()
    NET_INC = pd.read_excel('data.xlsx',sheet_name ='NET_Income').dropna()
    FFO = pd.read_excel('data.xlsx',sheet_name ='FFO').dropna()

    hide_table_row_index = """
                <style>
                tbody th {display:none}
                .blank {display:none}
                </style>
                """

    # Inject CSS with Markdown
    st.markdown(hide_table_row_index, unsafe_allow_html=True)
    #st.table(ST.style.format({"E": "{:.2f}"}))
    #  Single Tenants Assets : (visualisation) 
    st.markdown('## Single Tenants Assets')
    st.table(ST.applymap(lambda x: int(round(x, 0)) if isinstance(x, (int, float)) else x))
    test=pd.read_excel('visualisation.xlsx',sheet_name ='ST')
    ST_v =test.set_index('Single Tenants Assets')
    fig = plt.figure(figsize=(10,6), tight_layout=True)
    #fig, ax = plt.subplot()
    #plotting
    plt.plot(ST_v, 'o-', linewidth=2)
    #customization
    #plt.xticks(['2019', '2020','2021','2021','2021E','2021E','2021E','2022E'])
    plt.xlabel('Years')
    plt.ylabel('Revenues')
    plt.title('Revenues and forecasting revenues troughtout the years')
    plt.legend(title='Single Tenants Assets', title_fontsize = 13, labels=ST_v.columns,bbox_to_anchor=(1.,0.3), loc="lower left")
    st.write(fig)


    
    st.markdown('## Shopping centers')
    st.table(SHC.applymap(lambda x: int(round(x, 0)) if isinstance(x, (int, float)) else x))
    test = pd.read_excel('visualisation.xlsx',sheet_name ='SHC').dropna()
    SHC_v=test.set_index('Shopping centers')
    #plt.figure(figsize=(10,6), tight_layout=True)
    fig, ax= plt.subplots(figsize=(10,6),tight_layout=True)
    #plotting
    ax.plot(SHC_v, 'o-', linewidth=2)
    plt.xlabel('Years')
    plt.ylabel('Revenues')
    plt.title('Revenues and forecasting revenues troughtout the years')
    plt.legend(title='Shopping centers', title_fontsize = 13, labels=SHC_v.columns,bbox_to_anchor=(1.,0.3), loc="lower left")
    st.write(fig)

    
    st.markdown('## Commercial Galleries')
    st.table(CG.applymap(lambda x: int(round(x, 0)) if isinstance(x, (int, float)) else x).astype(str))
    test = pd.read_excel('visualisation.xlsx',sheet_name ='CG').dropna()
    CG_V=test.set_index('Commercial Galleries')
    #plt.figure(figsize=(10,6), tight_layout=True)
    fig, ax= plt.subplots(figsize=(10,6),tight_layout=True)
    #plotting
    ax.plot(CG_V, 'o-', linewidth=2)
    plt.xlabel('Years')
    plt.ylabel('Revenues')
    plt.title('Revenues and forecasting revenues troughtout the years')
    plt.legend(title='Commercial Galleries', title_fontsize = 13, labels=CG_V.columns,bbox_to_anchor=(1.,0.3), loc="lower left")
    st.write(fig)
    
    st.markdown('## Industrial Unit')
    st.table(IU.applymap(lambda x: int(round(x, 0)) if isinstance(x, (int, float)) else x))
    test = pd.read_excel('visualisation.xlsx',sheet_name ='IU').dropna()
    IU_V=test.set_index('Industrial Unit')
    #plt.figure(figsize=(10,6), tight_layout=True)
    fig, ax= plt.subplots(figsize=(10,6),tight_layout=True)
    #plotting
    ax.plot(IU_V, 'o-', linewidth=2)
    plt.xlabel('Years')
    plt.ylabel('Revenues')
    plt.title('Revenues and forecasting revenues troughtout the years')
    plt.legend(title='Industrial Unit', title_fontsize = 13, labels=IU_V.columns,bbox_to_anchor=(1.,0.3), loc="lower left")
    st.write(fig)
    
    st.markdown('## Assets total revenue')
    st.table(TR1.applymap(lambda x: int(round(x, 0)) if isinstance(x, (int, float)) else x).astype(str))
    #st.markdown('## Shopping centers')
    st.table(TR2.applymap(lambda x: int(round(x, 0)) if isinstance(x, (int, float)) else x))
    test = pd.read_excel('visualisation.xlsx',sheet_name ='TR2').dropna()
    IU_V=test.set_index('(KMAD)')
    #plt.figure(figsize=(10,6), tight_layout=True)
    fig, ax= plt.subplots(figsize=(10,6),tight_layout=True)
    #plotting
    ax.plot(IU_V, 'o-', linewidth=2)
    plt.xlabel('Years')
    plt.ylabel('Revenues')
    plt.title('Revenues and forecasting revenues troughtout the years')
    plt.legend(title='Industrial Unit', title_fontsize = 13, labels=IU_V.columns,bbox_to_anchor=(1.,0.3), loc="lower left")
    st.write(fig)
    
    st.markdown('## Revenues ')
    st.table(NET_INC.applymap(lambda x: int(round(x, 0)) if isinstance(x, (int, float)) else x))
    test = pd.read_excel('visualisation.xlsx',sheet_name ='Net_Income').dropna()
    Net_Income =test.set_index('Years')
    #plt.figure(figsize=(10,6), tight_layout=True)
    fig, ax= plt.subplots(figsize=(15,9),tight_layout=True)
    #plotting
    ax.plot(Net_Income, 'o-', linewidth=2)
    plt.xlabel('Years')
    plt.ylabel('Revenues')
    plt.title('Revenues and forecasting revenues troughtout the years')
    plt.legend(title='Industrial Unit', title_fontsize = 13, labels=Net_Income.columns,bbox_to_anchor=(1.,0.3), loc="lower left")
    fig1, ax1= plt.subplots(figsize=(15,9),tight_layout=True)
    plt.legend(title_fontsize = 13, labels=Net_Income.columns,bbox_to_anchor=(1.,0.1), loc="lower left")
    Net_Income.plot(ax=ax1,kind='bar', stacked=True,figsize=(15, 8),color=["#ADD8E6","#00BFFF"])
    st.write(fig1)
    st.write(fig)
    
    st.markdown('## Fund From Operation')
    st.table(FFO.applymap(lambda x: int(round(x, 0)) if isinstance(x, (int, float)) else x))
    test = pd.read_excel('visualisation.xlsx',sheet_name ='FFO').dropna()
    FFO1 =test.set_index('Years')
    #plt.figure(figsize=(10,6), tight_layout=True)
    fig, ax= plt.subplots(figsize=(15,9),tight_layout=True)
    #plotting
    #ax.plot(FFO, 'o-', linewidth=2,color='red')
    plt.xlabel('Years')
    plt.ylabel('Revenues')
    plt.title('Revenues and forecasting revenues troughtout the years')
    plt.legend(title='Industrial Unit', title_fontsize = 13, labels=Net_Income.columns,bbox_to_anchor=(1.,0.3), loc="lower left")
    plt.legend(title_fontsize = 13, labels=Net_Income.columns,bbox_to_anchor=(1.,0.1), loc="lower left")
    FFO1.plot.bar(ax=ax, stacked=True,figsize=(15, 8),color=["#00BFFF"])
    st.write(fig)
    df_xlsx = to_excel_M(ST,SHC,CG,IU,TR1,TR2,NET_INC,FFO)
    st.download_button(label='📥 Download Current Result',
                                data=df_xlsx ,
                                file_name= 'FFO_prediction.xlsx')




    
    
