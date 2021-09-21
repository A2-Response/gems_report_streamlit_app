#!/usr/bin/env python
# coding: utf-8

import streamlit as st
#from process_bom import *
import pandas as pd
import sys
import os
import copy

def process_data(path):
    if path.name.endswith('.csv'):
        df = pd.read_csv(path,sep=',')
    else:
        df = pd.read_csv(path,sep='\t')
    indices = list(df[df['Level']==1].index)
    copy_indices = copy.deepcopy(indices)
    copy_indices.append(df.shape[0])
    
    for j,i in enumerate(indices):
        df.loc[i:copy_indices[j+1],'Level_1'] = df.loc[i,'Parent PN']+'_'+str(j)+'_1'
        df.loc[i:copy_indices[j+1],'Component Qty'] = df.loc[i,'Name']
        df.loc[i:copy_indices[j+1],'UoM'] = df.loc[i,'Desc']        
    return df
    
def groupd_by(df_sub):
    for j in df_sub.loc[df_sub['Type'].isin(['Circuit Board Assembly','Tested Hybrid-Mcm'])].index:
        level = df_sub.loc[j,'Level']
        pwapn = df_sub.loc[j,'Name']
        pwadesc = df_sub.loc[j,'Desc']    
        pwaproj = df_sub.loc[j,'Project']
        pwaqty = df_sub.loc[j,'Qty']

        for i in range(1,df_sub.loc[j:,].shape[0]):
            if df_sub.loc[j+i,'Level'] > level:
                pass
            else:
                k = j+i-1
                df_sub.loc[j+1:k,'Temp1'] = pwapn
                df_sub.loc[j+1:k,'Temp2'] = pwadesc
                df_sub.loc[j+1:k,'Temp3'] = pwaqty
                df_sub.loc[j+1:k,'Temp4'] = pwaproj
                break
    return df_sub
    
def get_asset_to_pwa(df):
    df_temp = df.groupby(['Level_1']).apply(lambda x: groupd_by(x))
    df1 = df_temp.dropna(subset=['Temp1','Temp2','Temp3','Temp4'])
    df2 = df1[['Component Qty','Temp1','Temp2','Temp3']].drop_duplicates()
    df2.columns = ['Asset PN','PWA PN','PWA Desc','Quantity']
    return df2
    
def get_pwa_bom(df):
    df_temp = df.groupby(['Level_1']).apply(lambda x: groupd_by(x))
    df1 = df_temp.dropna(subset=['Temp1','Temp2','Temp3','Temp4'])
    df2 = df1[['Temp1','Name','Qty']]
    df2 = df2.loc[(df2['Temp1'] != df2['Name']),]
    df2.columns = ['PWA PN','Name','Qty']
    return df2.reset_index(drop=True)
    
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
    df2_1 = df2_1.drop_duplicates()
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
    df_temp.loc[(df_temp['Qty']==0),'Qty'] = 0
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
    return df_temp
    
def get_asset_to_component(df):
    df = get_component_qty(df)
    df_asset_to_bom = df.groupby(['Component Qty','Name'])['Temp5'].sum()
    df_asset_to_bom = df_asset_to_bom.reset_index(drop=False)
    df_asset_to_bom = df_asset_to_bom[df_asset_to_bom['Name']!=df_asset_to_bom['Component Qty']]
    df_asset_to_bom.reset_index(drop=True,inplace=True)
    df_asset_to_bom.columns = ['Asset PN','Child PN','Quantity']
    return df_asset_to_bom
    
def get_whereused(df):
    df = get_component_qty(df)
    df_sub = df.loc[:,['UoM', 'Component Qty', 'Temp1', 'Temp2',
       'Temp3', 'Temp4','Temp5']]
    df_sub['Temp3'] = df_sub['Temp3'].fillna(1)
    df_sub['Temp6'] = df_sub['Temp3']*df_sub['Temp5']
    df_sub = df_sub.drop('Temp3',axis=1)
    df_sub = df_sub.drop('Temp5',axis=1)
    df_sub = df_sub[['Temp1', 'Temp2', 'Temp4', 'Component Qty','UoM','Temp6']]
    df_sub = pd.concat([df.loc[:,'Name':'Project'],df_sub],axis=1)
    df_sub['Rev'] = df_sub['Desc']
    df_sub['Desc'] = df['Type']
    df_sub['RoHS'] = df['RoHS']
    df_sub['Proc Code'] = df['Proc Code']
    df_sub = pd.concat([df_sub,df.loc[:,'Green':'Date of Intro']],axis=1)
    
    df_sub_non_filt = df_sub[~df_sub['Temp1'].isnull()]
    df_sub_non_filt = df_sub_non_filt.drop_duplicates()
    
    df_sub_filt = df_sub[df_sub['Temp1'].isnull()]
    
    df_sub_filt['CONCAT'] = df_sub_filt[['Name','Component Qty']].apply(lambda x: x[0]+'A'+x[1],axis=1)
    df_sub_filt = df_sub_filt.rename(columns={'Temp6':'Qty'})
    df_sub_filt_concat_grpd = df_sub_filt[['CONCAT','Qty']].groupby('CONCAT')['Qty'].sum().reset_index()
    df_sub_filt_concat_grpd.columns = ['CONCAT','Qty_1']
    df_sub_filt = pd.merge(df_sub_filt, df_sub_filt_concat_grpd,on='CONCAT',how='left')
    
    df_sub_filt['Qty'] = df_sub_filt['Qty_1']
    df_sub_filt.drop('Qty_1',axis=1,inplace=True)
    df_sub_filt.drop('CONCAT',axis=1,inplace=True)
    df_sub_non_filt = df_sub_non_filt.rename(columns={'Temp6':'Qty'})
    df_final = pd.concat([df_sub_non_filt,df_sub_filt])
    df_final.columns = ["Child PN","Child Desc","Child Part Type","Child Project","PWA PN","PWA Desc","PWA Project","Asset PN","Asset Desc","Quantity",
    "RoHS","Procurability","Green","Date of Introduction"]
    df_final = df_final[["Child PN","Child Desc","Child Part Type","Child Project","RoHS",
                         "Procurability","Green","Date of Introduction","PWA PN","PWA Desc","PWA Project",
                         "Asset PN","Asset Desc","Quantity"]]
    df_final = df_final.drop_duplicates()
    df_final.reset_index(drop=True,inplace=True)
    
    return df_final

def main():
    st.title("Application For BoM Report Creation...")
    uploaded_file = st.file_uploader("Upload Files",type=['csv','xls'])
    if uploaded_file is not None:
        file_details = {"FileName":uploaded_file.name,"FileType":uploaded_file.type,"FileSize":uploaded_file.size}
        st.write(file_details)
    if st.button('Process'):
        data_path = uploaded_file
        with st.spinner('The Report is getting generated...'):
            df = process_data(data_path)

            asset_to_pwa = get_asset_to_pwa(df)

            pwa_bom = get_pwa_bom(df)

            part_list = get_part_list(df)

            asset_bom = get_asset_to_component(df)

            whereused = get_whereused(df)

            df = df.drop('Level_1',axis=1)
            
            loc = os.path.expanduser('~')
            
            writer = pd.ExcelWriter(loc+'final.xlsx')
            df.to_excel(writer,'BoM',index=False)
            asset_to_pwa.to_excel(writer,'Asset To PWA',index=False)
            asset_bom.to_excel(writer,'Asset BoM',index=False)
            pwa_bom.to_excel(writer,'PWA BoM',index=False)
            whereused.to_excel(writer,'Whereused',index=False)
            part_list.to_excel(writer,'Part List',index=False)
            writer.save()
            writer.close()
            st.success(f'Process Done! Report saved at {loc}file.xlsx')
    
if __name__ == '__main__':
    main()

