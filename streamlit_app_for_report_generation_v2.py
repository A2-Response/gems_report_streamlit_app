#!/usr/bin/env python
# coding: utf-8

import streamlit as st
#from process_bom import *
import pandas as pd
import sys
import os
import copy
import base64
import io

def process_data(path):
    df = pd.read_csv(path,sep='\t')
    indices = list(df[df['Level']==1].index)
    copy_indices = copy.deepcopy(indices)
    copy_indices.append(df.shape[0])
    
    for j,i in enumerate(indices):
        df.loc[i:copy_indices[j+1],'Level_1'] = df.loc[i,'Parent PN']+'_'+str(j)+'_1'
        df.loc[i:copy_indices[j+1],'Asset PN'] = df.loc[i,'Name']
        df.loc[i:copy_indices[j+1],'Asset Description'] = df.loc[i,'Desc']        
    return df
    
def groupd_by(df_sub):
    for j in df_sub.loc[df_sub['Type'].isin(['Circuit Board Assembly','Tested Hybrid-Mcm'])].index:
        level = df_sub.loc[j,'Level']
        pwapn = df_sub.loc[j,'Name']
        pwadesc = df_sub.loc[j,'Desc']    
        pwaproj = df_sub.loc[j,'Project']
        pwaqty = df_sub.loc[j,'Qty']
        test_ind = df_sub.loc[j:,].shape[0] - 1
        i = 1
        while i <= test_ind:
            if ((df_sub.loc[j+i,'Level'] > level) and (i != test_ind)):
                i += 1
                pass
            elif ((df_sub.loc[j+i,'Level'] > level) and (i == test_ind)):
                k = j+i
                df_sub.loc[j+1:k,'PWA PN'] = pwapn
                df_sub.loc[j+1:k,'PWA Description'] = pwadesc
                df_sub.loc[j+1:k,'PWA Qty'] = pwaqty
                df_sub.loc[j+1:k,'PWA Project'] = pwaproj
                break
            else:
                k = j+i-1
                df_sub.loc[j+1:k,'PWA PN'] = pwapn
                df_sub.loc[j+1:k,'PWA Description'] = pwadesc
                df_sub.loc[j+1:k,'PWA Qty'] = pwaqty
                df_sub.loc[j+1:k,'PWA Project'] = pwaproj
                break
    return df_sub

def get_pwa_qty(df):
    df_temp = df.groupby(['Level_1']).apply(lambda x: groupd_by(x))
    df_temp.loc[:,'Temp5'] = 1
    df_temp.loc[:,'Temp6'] = df_temp.loc[:,'Qty']
    df_temp.reset_index(drop=True,inplace=True)
    lastrow = df_temp.shape[0]-1
    max_level = max(df_temp['Level'])
    df_temp.loc[df_temp['Type'].isin(['Circuit Board Assembly','Tested Hybrid-Mcm']),'Temp6'] = 1 
    for a in range(2,max_level+1):
        i = 1
        while i < lastrow:
            if df_temp.iloc[i,2] == a:
                rowstart = i
                qty = df_temp.loc[i,'Temp6']
                df_temp.loc[i,'Temp5'] = df_temp.loc[i,'Temp5']*qty
                if rowstart != lastrow:
                    for j in range(rowstart+1,lastrow+1):
                        if df_temp.iloc[j,2] > a:
                            df_temp.loc[j,'Temp5'] = df_temp.loc[j,'Temp5']*qty
                            rowend = j
                        else:
                            rowend = j
                            break
                else:
                    rowend = lastrow
                if rowend != lastrow:
                    i = rowend
                else:
                    i = lastrow
            else:
                i+=1
    df_temp = df_temp.rename(columns={'Temp5':'Component Qty'})
    df_temp.drop('Temp6',axis=1,inplace=True)
    return df_temp
    
def get_asset_to_pwa(df):
    df_temp = df.groupby(['Level_1']).apply(lambda x: groupd_by(x))
    df1 = df_temp.dropna(subset=['PWA PN','PWA Description','PWA Qty','PWA Project'])
    df2 = df1[['Asset PN','Asset Description','PWA PN','PWA Description','PWA Qty']].drop_duplicates()
    df2.columns = ['Asset PN','Asset Description','PWA PN','PWA Desc','Quantity']
    df2 = df2.drop_duplicates(subset=['Asset PN','PWA PN'])
    return df2
    
def get_pwa_bom(df):
    df_temp = df.groupby(['Level_1']).apply(lambda x: groupd_by(x))
    df1 = df_temp.dropna(subset=['PWA PN'])
    df2 = df1[['Asset PN','PWA PN','Name','Component Qty']]
    df_pwa_bom = df2.groupby(['Asset PN','PWA PN','Name'])['Component Qty'].sum().reset_index(drop=False)
    df_pwa_bom = df_pwa_bom.drop('Asset PN',axis=1)
    df_pwa_bom = df_pwa_bom.drop_duplicates(subset=['PWA PN','Name'])
    df_pwa_bom = df_pwa_bom.loc[(df_pwa_bom['PWA PN'] != df_pwa_bom['Name']),]
    df_pwa_bom.columns = ['PWA PN','Child PN','Quantity']
    return df_pwa_bom.reset_index(drop=True)
    
def get_part_list(df):
    df_temp = df.groupby(['Level_1']).apply(lambda x: groupd_by(x))
    ind = df_temp.columns
    ind = list(ind)
    strt_ind = ind.index('Level')
    end_ind = ind.index('Date of Intro')+1
    df2 = df_temp[ind[strt_ind:end_ind]]
    df2_ = df2[~((df2['Level']==1)|(df2['Type'].isin(['Circuit Board Assembly','Tested Hybrid-Mcm'])))]
    l = [i for i in df2_.columns if i not in ['Usage','Qty','Level','Rev']]
    df2_1 = df2_[l]
    df2_1 = df2_1.drop_duplicates(subset=['Type'])
    l1 = [i for i in df2_1.columns if i not in ['PB-Free','CE Mark']]
    df2_2 = df2_1[l1]
    l4 = ['Name',
     'Type',
     'Desc',
     'Design Responsibility',
     'RoHS',
     'EU RoHS Exemption',
     'Proc Code',
     'Proc Date',
     'Date of Intro',
     'Green'
     ]
    df2_3 = df2_2[l4]
    df2_3.columns = ['Child PN','Child Part Type','Child Desc','Child Project',
                 'RoHS','EU RoHS Exemption','Procurability','Proc Date','Date of Introduction','Green']
    df2_3 = df2_3.reset_index(drop=True)
    return df2_3
    
def get_component_qty(df):
    df_temp = df.groupby(['Level_1']).apply(lambda x: groupd_by(x))
    df_temp.loc[:,'Temp5'] = 1
    df_temp.reset_index(drop=True,inplace=True)
    lastrow = df_temp.shape[0]-1
    max_level = max(df_temp['Level'])
    for a in range(2,max_level+1):
        i = 1
        while i < lastrow:
            if df_temp.iloc[i,2] == a:
                rowstart = i
                qty = df_temp.loc[i,'Qty']
                df_temp.loc[i,'Temp5'] = df_temp.loc[i,'Temp5']*qty
                if rowstart != lastrow:
                    for j in range(rowstart+1,lastrow+1):
                        if df_temp.iloc[j,2] > a:
                            df_temp.loc[j,'Temp5'] = df_temp.loc[j,'Temp5']*qty
                            rowend = j
                        else:
                            rowend = j
                            break
                else:
                    rowend = lastrow
                if rowend != lastrow:
                    i = rowend
                else:
                    i = lastrow
            else:
                i+=1
    df_temp['Component Qty'] = df_temp['Temp5']
    df_temp.drop('Temp5',axis=1,inplace=True)
    return df_temp
    
def get_asset_to_component(df):
    df = get_component_qty(df)
    df = df[df['PWA PN'].isnull()]
    df['Asset PN'] = df['Asset PN'].astype('str')
    df['Name'] = df['Name'].astype('str')
    df_asset_to_bom = df.groupby(['Asset PN','Name'])['Component Qty'].sum()
    df_asset_to_bom = df_asset_to_bom.reset_index(drop=False)
    df_asset_to_bom = df_asset_to_bom[df_asset_to_bom['Name']!=df_asset_to_bom['Asset PN']]
    df_asset_to_bom.reset_index(drop=True,inplace=True)
    df_asset_to_bom.columns = ['Asset PN','Child PN','Quantity']
    return df_asset_to_bom
    
def get_whereused(df):
    df = get_component_qty(df)
    ind = df.columns
    ind = list(ind)
    strt_ind = ind.index('Type')
    end_ind = ind.index('Component Qty')+1
    df2 = df[ind[strt_ind:end_ind]]
    #df2['PWA PN'] = df2['PWA PN'].fillna('')
    df2.loc[:,'CONCAT'] = df2.loc[:,['Asset PN','PWA PN','Name']].apply(lambda x: 'A'+x[0]+x[2] if pd.isnull(x[1]) else 'A'+x[0]+x[1]+x[2], axis=1) 

    df_sub_filt_concat_grpd = df2[['CONCAT','Component Qty']].groupby('CONCAT')['Component Qty'].sum().reset_index()
    df_sub_filt_concat_grpd.columns = ['CONCAT','Qty_1']
    df_sub_filt = pd.merge(df2, df_sub_filt_concat_grpd,on='CONCAT',how='left')
    df_sub_filt = df_sub_filt.drop_duplicates(subset=['CONCAT'])
    df_sub_filt.drop(['Component Qty','CONCAT'],axis=1,inplace=True)
    df_sub_filt = df_sub_filt.rename(columns={'Qty_1':'Component Qty'})

    df_sub_filt = df_sub_filt[df_sub_filt['Name']!=df_sub_filt['Asset PN']]
    df_sub_filt = df_sub_filt[df_sub_filt['Name']!=df_sub_filt['PWA PN']]
    df_sub_filt = df_sub_filt[df_sub_filt['Asset PN']!=df_sub_filt['PWA PN']]
    
    df_sub_filt = df_sub_filt[df_sub_filt['Component Qty']!=0]
    df_sub_filt_sub = df_sub_filt[['Type','Name','Rev','PWA PN', 
                                   'PWA Description', 'PWA Project','Asset PN', 
                                   'Asset Description','Component Qty','RoHS', 
                                   'Proc Code', 'Proc Date', 'CE Mark','EU RoHS Exemption', 
                                   'Green', 'Date of Intro']]
    df_final = df_sub_filt_sub.drop_duplicates()
    df_final.reset_index(drop=True,inplace=True)
    
    return df_final

def main():
    st.title("Application For BoM Report Creation...")
    uploaded_file = st.file_uploader("Upload Files",type=['xls'])
    if uploaded_file is not None:
        #file_details = {"FileName":uploaded_file.name,"FileType":uploaded_file.type,"FileSize":uploaded_file.size}
        st.write('File Uploaded!')
    if st.button('Process'):
        data_path = uploaded_file
        with st.spinner('The Report is getting generated...'):
            df = process_data(data_path)
            
            part_list = get_part_list(df)
            asset_to_pwa = get_asset_to_pwa(df)
            df = get_pwa_qty(df)
            pwa_bom = get_pwa_bom(df)
            asset_bom = get_asset_to_component(df)
            whereused = get_whereused(df)

            df = df.drop('Level_1',axis=1)
                   
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer,'BoM',index=False,encoding='utf-8')
            asset_to_pwa.to_excel(writer,'Asset To PWA',index=False,encoding='utf-8')
            asset_bom.to_excel(writer,'Asset BoM',index=False,encoding='utf-8')
            pwa_bom.to_excel(writer,'PWA BoM',index=False,encoding='utf-8')
            whereused.to_excel(writer,'Whereused',index=False,encoding='utf-8')
            part_list.to_excel(writer,'Part List',index=False,encoding='utf-8')
            writer.save()
            processed_data = output.getvalue()
            val = processed_data
            b64 = base64.b64encode(val)
            st.success(f'Process Done!')
            linko= f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="BoM_Report.xlsx">Download Report file</a>'
            st.markdown(linko, unsafe_allow_html=True)
    
if __name__ == '__main__':
    main()

