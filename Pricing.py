import pandas as pd
import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
#from mpl_toolkits.mplot3d import Axes3D
import seaborn as sns
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import base64

def app():


    apptitle = 'CFA ARADEI CAPITAL'

    st.image('ensa.png',width=500)
    #st.set_page_config(page_title=apptitle, page_icon=":chart_with_upwards_trend:")

    # Title the app
    st.title(' CFA Research Challenge Morocco')

    # @st.cache(ttl=3600, max_entries=10)   #-- Magic command to cache data

    # @st.cache(ttl=3600, max_entries=10)   #-- Magic command to cache data
        
    


    st.markdown('### Subject Company :  ARADEI CAPITAL  ')

    #st.markdown('<center><img src="ensa.png" width="300"  height="100" alt="Ensa logo"></center>', unsafe_allow_html=True)
   
    st.markdown('##')
    st.markdown('__________________________________________________________')
        
    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data

    def to_excel_M(df1,df2):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df1.to_excel(writer, index=False, sheet_name='Sheet1')
        df2.to_excel(writer, index=False, sheet_name='Sheet2')
        workbook = writer.book
        worksheet1 = writer.sheets['Sheet1']
        worksheet2 = writer.sheets['Sheet2']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet1.set_column('A:A', None, format1)
        worksheet2.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data                                                          

        
    df = pd.read_csv ('d.csv')
    st.markdown('## Valuation')
    rc = st.number_input('Cost of Equity % : ')/100
    g = st.number_input('Growth rate % : ')/100
    #st.write('The Rc number is ', rc)
    #st.write('The g number is ', g)


    if rc or g !=0:
        DP = list()
        #DP.append('Dividends paid')
        factor=list()
        #factor.append('Dividends factor')
        for i in range(1,6):
            DP.append(df[str(2021+i)][0])
            factor.append(1/(1+rc)**int(i))
            #print(1/(1+rc)**int(i))

        DDP=list()
        #DDP.append('Discounted dividends paid')
        somme=0
        for i in range(1,6):
            #print(df[str(2021+i)][0]*factor[i-1])
            DDP.append(df[str(2021+i)][0]*factor[i-1])
          
        TV =(DP[4]*(1+g))/(rc-g)
        DTV= TV*factor[4]
        land =596029
        nmbr_of_share = 10645783

        equity_value = sum(DDP)+DTV+land

        price = (equity_value/nmbr_of_share)*1000

        data=  pd.DataFrame([DP,factor,DDP], columns = ['2022', '2023','2024','2025','2026'])


        data.insert(loc=0, column="KMAD", value=['Dividends paid','Discount Factor','Discounted dividends paid'])


        data.loc[4]=['Terminal value',0,0,0,0,TV]

        hide_table_row_index = """
                    <style>
                    tbody th {display:none}
                    .blank {display:none}
                    </style>
                    """

        # Inject CSS with Markdown
        st.markdown(hide_table_row_index, unsafe_allow_html=True)

        st.table(data)

        data1=  pd.DataFrame(['Sum of discounted Dividends paid to shareholders','Discounted Terminal Value','Land Reserves','Equity Value','Number of Shares','Share Price in MAD'], columns = ['KMAD'])
        data1['value']= [sum(DDP),DTV,land,equity_value,nmbr_of_share,price]
        #change color of dataframe
        #data1.style.set_properties(**{'background-color': 'white','color': 'green'})

        #st.table(data1.style.set_properties(**{'background-color': 'white','color': 'green'}))
        st.table(data1)                           
        #st.write(price)

        st.download_button(label='📥 Download Current Result',
                                                data=to_excel_M(data,data1) ,
                                                file_name= 'DDM.xlsx')
    #---------------------------------------------sensitivity matrix
    i=0

    st.markdown('## Sensitivity matrix')
    rc= st.multiselect("Choose the range of Re ",[round(i,2) for i in np.arange(1,10,0.01)]
    )
    g= st.multiselect(" Choose the range of g",[round(i,2) for i in np.arange(1,10,0.01)]
    )
    #rc=[5.30/100,5.40/100,5.59/100,6.20/100,6.40/100]
    #g=[1.50/100,1.75/100,2.00/100,2.25/100,2.50/100]
    matrix=  pd.DataFrame(index=[str(round(i,2))+'%' for i in g])

    if rc and g!=None:
        while i<len(rc):
            j=0
            price=list()
            while j<len(g):
                

                DP = list()
                #DP.append('Dividends paid')
                factor=list()
                #factor.append('Dividends factor')
                for x in range(1,6):
                    DP.append(df[str(2021+x)][0])
                    factor.append(1/(1+rc[i]/100)**int(x))
                    #print(1/(1+rc)**int(i))

                DDP=list()
                #DDP.append('Discounted dividends paid')
                somme=0
                for x in range(1,6):
                    #print(df[str(2021+i)][0]*factor[i-1])
                    DDP.append(df[str(2021+x)][0]*factor[x-1])

                TV =DP[4]*(1+g[j]/100)/(rc[i]/100-g[j]/100)
                DTV= TV*factor[4]
                land =596029
                nmbr_of_share = 10645783

                equity_value = sum(DDP)+DTV+land
                price.append((equity_value/nmbr_of_share)*1000)
                j=j+1
                
            matrix[str(rc[i])+'%']=price
            
            i=i+1
        
        st.write(' Sensitivity matrix (Cost of equity /Growth) : ')
        st.write(matrix)

        fig, ax = plt.subplots()
        sns.heatmap(matrix,annot=True, fmt=".2f", 
               linewidths=5, 
               cbar_kws={"shrink": .2}, ax=ax)
        st.write(fig)

        
      
        matrix_dow =pd.DataFrame()
        matrix_dow = matrix
        matrix_dow.insert(loc=0, column="RE/G", value=[str(round(i,2))+'%' for i in g])
      
        
        df_xlsx = to_excel(matrix_dow)
        st.download_button(label='📥 Download Current Result',
                                            data=df_xlsx ,
                                            file_name= 'Sensitivity matrix.xlsx')

