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
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import time
from selenium.common.exceptions import NoSuchElementException
from time import sleep
import pandas as pd
import shutil
import os
import pyautogui
import ast
import shutil
os.environ['DISPLAY'] = ':0'

def enable_download_headless(browser,download_dir):
        browser.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd':'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_dir}}
        browser.execute("send_command", params)
        return browser

def set_chrome_options(headless=True):
    chrome_options = Options()
    if headless == True:
        chrome_options.headless = True
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
    else:
        chrome_options.headless = False
    chrome_options.add_argument("--window-size=1920,1200")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    return chrome_options

def get_vault_info(handles,browser=None):
    seq = browser.find_elements_by_tag_name('iframe')
    browser.switch_to.frame(seq[0])
    browser.implicitly_wait(10)
    seq2 = browser.find_elements_by_tag_name('iframe')
    browser.switch_to.frame(seq2[0])
    s = browser.find_element_by_id("calc_vault")
    text = s.find_element_by_class_name('field').text
    #browser.switch_to_window(handles[1])
    return text.strip().replace(' ','%20')

def remwin(browser=None):
    default_handle = browser.current_window_handle
    handles = list(browser.window_handles)
    assert len(handles) > 1

    handles.remove(default_handle)
    assert len(handles) > 0

    browser.switch_to_window(handles[0])
    # do your stuffs
    browser.close()
    browser.switch_to_window(default_handle)
    return browser
    
def initiate_and_login_browser(headless=True):
    chrome_options = set_chrome_options(headless=headless)
    browser = webdriver.Chrome(chrome_options=chrome_options,executable_path='chromedriver') # give the path_to_chromedriver here
    browser.get('''
    https://www.gems.slb.com/''')
    time.sleep(15)
    username = 'Amaity2'
    password = '112358@Cg#Am'
    elementID = browser.find_element_by_xpath("//input[@placeholder='LDAP Alias']")
    elementID.send_keys(username)
    elementID = browser.find_element_by_xpath("//input[@placeholder='Password']")
    elementID.send_keys(password)
    elementID.submit()
    WebDriverWait(browser,150).until(EC.element_to_be_clickable((By.XPATH,"//font[@color='white']")))
    browser.find_element(By.XPATH,"//font[@color='white']").click()
    time.sleep(15)
    return browser

def get_latest_revisions(df,browser=None):
    if browser == None:
        browser = initiate_and_login_browser()
    for i in df.index:
        url='''https://www.gems.slb.com/enovia/engineeringcentral/SLBemxEngrSearchSummaryFS.jsp?txtWhere=&ckSaveQuery=&txtQueryName=&ckChangeQueryLimit=&queryLimit=100&fromDirection=&toDirection=&selType=*&ComboType=Part&sFindPartByName=*&txtName='''+str(df.loc[i,'PN'])+'''&txtNameHidden=&txtDesc=*&txtOwner=*&hiddenOrig=*&sFindPartByName=*&selVaults=ALL_VAULTS&selType=*&txtRev=*&txtOriginator=*&vaultOption=ALL_VAULTS&txtProject=*&revPattern=HIGHEST_PROTOTYPE_RELEASED_OBSOLETE&typeAheadFormName=%2Fengineeringcentral%2FSLBemxEngrSearchDialog.jsp&mode=powersearch&saveQuery=false&changeQueryLimit=true&FromCollection=FALSE'''
        browser.get(url)
        time.sleep(5)
        iframe = browser.find_element_by_xpath("//iframe[@name='pagecontent']")
        browser.switch_to.frame(iframe)

        for num in range(3,20):
            try:
                browser.find_element_by_xpath("/html/body/form/table/tbody/tr["+str(num)+"]/td[4]")
                trval=num
            except:
                break

        try:
            revtext=str(browser.find_element_by_xpath("/html/body/form/table/tbody/tr["+str(trval)+"]/td[4]").text).strip()
        except:
            revtext=""
        try:
            reltext=str(browser.find_element_by_xpath("/html/body/form/table/tbody/tr["+str(trval)+"]/td[7]").text).strip()
        except:
            reltext=""
        try:
            prttext=[str(browser.find_element_by_xpath("/html/body/form/table/tbody/tr["+str(trval)+"]/td[5]").text).strip()]
            try:
                revtext_1 = str(browser.find_element_by_xpath("/html/body/form/table/tbody/tr["+str(trval-1)+"]/td[4]").text).strip()
                if revtext == revtext_1:
                    prttext_1=str(browser.find_element_by_xpath("/html/body/form/table/tbody/tr["+str(trval-1)+"]/td[5]").text).strip()
                    prttext.append(prttext_1)
            except:
                pass
        except:
            prttext=""
        df.loc[i,['RelVer','Stat','PART']]=[revtext.strip(),reltext.strip(),prttext]
        print("PN:"+str(df.loc[i,'PN'])+" "+revtext.strip()+" "+reltext.strip()+" ",prttext)
        browser.switch_to.default_content()
    df['PART'] = df['PART'].apply(lambda x: [i.replace(' ','%20') for i in x])
    return df, browser

def new_clickable_table_expand(tr_len,browser=None):
    iframe = browser.find_element_by_xpath("//iframe[@name='pagecontent']")
    browser.switch_to.frame(iframe)
    browser.find_element_by_xpath(f"/html/body/form/table/tbody/tr[{str(tr_len)}]/td[8]/a").click()
    return browser

def get_total_no_tab_rows(browser=None):
    iframe = browser.find_element_by_xpath("//iframe[@name='pagecontent']")
    browser.switch_to.frame(iframe)
    s = browser.find_element_by_xpath("/html/body/form/table/tbody")
    tr_len = len(s.find_elements_by_tag_name('tr'))
    browser.switch_to.default_content()
    return tr_len

def run_vault_process(df,rem_win=True,browser=None):
    for i in df.index:
        url='''https://www.gems.slb.com/enovia/engineeringcentral/SLBemxEngrSearchSummaryFS.jsp?txtWhere=&ckSaveQuery=&txtQueryName=&ckChangeQueryLimit=&queryLimit=100&fromDirection=&toDirection=&selType=*&ComboType=Part&sFindPartByName=*&txtName='''+str(df.loc[i,'PN'])+'''&txtNameHidden=&txtDesc=*&txtOwner=*&hiddenOrig=*&sFindPartByName=*&selVaults=ALL_VAULTS&selType=*&txtRev='''+str(df.loc[i,'RelVer']).strip()+'''&txtOriginator=*&vaultOption=ALL_VAULTS&txtProject=*&revPattern=HIGHEST_PROTOTYPE_RELEASED_OBSOLETE&typeAheadFormName=%2Fengineeringcentral%2FSLBemxEngrSearchDialog.jsp&mode=powersearch&saveQuery=false&changeQueryLimit=true&FromCollection=FALSE'''
        browser.get(url)
        time.sleep(5)
        total_no_rows = get_total_no_tab_rows(browser=browser)
    
        vault_info = 'na'
        row_check = total_no_rows
        while vault_info == 'na':
            browser = new_clickable_table_expand(row_check,browser=browser)
            time.sleep(5)
            handles = list(browser.window_handles)
            browser.switch_to_window(handles[1])
            time.sleep(8)
            try:
                vault_info = get_vault_info(handles,browser=browser)
            except Exception as e:
                vault_info = 'na'
                row_check -= 1
                browser.close()
                browser.switch_to_window(handles[0])
        if rem_win:
            browser = remwin(browser=browser)
        time.sleep(8)
        df.loc[i,'Vault'] = vault_info.strip()
        print("PN:"+str(df.loc[i,'PN'])+" "+vault_info.strip())
        browser.switch_to.default_content()
    return df, browser

def movetofolder(s,cwd_file):
    src= os.path.expanduser('~') + '\\Downloads'+s
    if not os.path.exists(fr"{cwd_file}"):
        os.mkdir(fr"{cwd_file}")
    dest = fr"{cwd_file}" + s
    try:
        os.rename(src, dest)
        print('The file:'+s[1:]+' is successfully moved to '+dest)
        return dest
    except Exception as e:
        print(e)
        print('Error Unable to Move File')

def prepare_df_for_download(path,browser=None):
    df=pd.read_excel(path,engine='openpyxl')
    centre = df['Centre'].unique()[0]
    for i in df.index:
        df.loc[i,'PN']=str(df.loc[i,'PN']).strip()
        df.loc[i,['RelVer','Stat']]=['na','non']
    df , browser = get_latest_revisions(df,browser=browser)
    df , browser = run_vault_process(df,browser=browser)
    df.to_csv('revparts.csv',index=False)
    browser.close()
    return centre

def get_bom_eol_reports(browser=None,headless=False):
    df=pd.read_csv(r'revparts.csv')
    df['PART'] = df['PART'].apply(lambda x: ast.literal_eval(x))
    cwd = os.getcwd()#os.path.dirname(os.path.abspath(__file__))
    cwd_file = cwd+"\\BOMSTAT"
    if not os.path.exists(cwd_file):
        os.mkdir(cwd_file)
    f2 = "\BoM EOL Stat.xls"
    strt=time.perf_counter()
    if browser == None:
        browser = initiate_and_login_browser(headless=headless)
    time.sleep(10)
    download_dir = os.path.expanduser('~')+'\\Downloads'
    browser = enable_download_headless(browser,download_dir)
    for i in range(df.shape[0]):
        for j in range(len(df.loc[i,'PART'])):
            try:
                url='''https://www.gems.slb.com/enovia/report/SLBIntermidiateBuildReport.jsp?reportId=63947.16629.36608.63451&busObjType='''+str(df.loc[i,'PART'][j])+'''&busObjName='''+str(df.loc[i,'PN'])+'''&busObjRevision='''+str(df.loc[i,'RelVer'])+'''
                &busObjVault='''+str(df.loc[i,'Vault'])+'''&targetLocation=popup'''
                try:
                    browser.get(url)
                    time.sleep(2)
                    browser.find_element(By.XPATH,"//*[@id='GenerateReport']").click()
                    time.sleep(5)
                    WebDriverWait(browser,150).until(EC.element_to_be_clickable((By.XPATH,"/html/body/table/tbody/tr[3]/td[2]/table[2]/tbody/tr/td[1]/table/tbody/tr/td[2]/a/img")))
                    browser.find_element(By.XPATH,"/html/body/table/tbody/tr[3]/td[2]/table[2]/tbody/tr/td[1]/table/tbody/tr/td[2]/a/img").click()
                    time.sleep(3)
                    movetofolder(f2,cwd_file)
                    os.rename(cwd_file+f2,cwd_file+'%s_%s.xls'%(f2.split('.xls')[0],str(df.loc[i,'PN']).strip()))
                except NoSuchElementException as e:
                    print("not this")
                print("PN:"+df.loc[i,'PN'])
                time.sleep(2)
                browser = remwin(browser=browser)
                time.sleep(10)
                break
            except Exception as e1:
                print(e1)
                pass
    browser.close()
    os.remove('revparts.csv')
    
def make_consolidated_file():
    files = os.listdir('BOMSTAT/')
    df = pd.DataFrame()
    for f in files:
        temp_df = pd.read_csv('BOMSTAT/'+f,sep='\t')
        df = pd.concat([df,temp_df])
    save_path = os.getcwd()+'\\consolidate_BOM_report'
    if not os.path.exists(save_path):
        os.mkdir(save_path)
    df.to_csv(save_path+'\\BOM_EOL_Stat_full.csv',index=False)
    shutil.rmtree('BOMSTAT')

def download_main(path):
    centre = prepare_df_for_download(path)
    get_bom_eol_reports()
    print("Saving the consolidated file...")
    make_consolidated_file()
    print("\n Task Done!")
    return centre

def process_data(path):
    df = pd.read_csv(path,sep=',')
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
    l = [i for i in df2_.columns if i not in ['Usage','Qty','Level']]
    df2_1 = df2_[l]
    df2_1 = df2_1.drop_duplicates(subset=['Name'])
    l1 = [i for i in df2_1.columns if i not in ['PB-Free','CE Mark']]
    df2_2 = df2_1[l1]
    l4 = ['Name',
     'Type',
     'Desc',
     'Rev',
     'Design Responsibility',
     'RoHS',
     'EU RoHS Exemption',
     'Proc Code',
     'Proc Date',
     'Date of Intro',
     'Green'
     ]
    df2_3 = df2_2[l4]
    df2_3.columns = ['Child PN','Child Part Type','Child Desc','Rev','Child Project',
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
    st.title("Application For BoM Report Creation")
    st.write("Sample Part Number file: https://github.com/A2-Response/working_monitor/blob/main/PN.xlsx?raw=true")
    uploaded_file = st.file_uploader("Upload Part Number Excel File", type=['xlsx'])
    if uploaded_file is not None:
        file_details = {"FileName":uploaded_file.name,"FileType":uploaded_file.type,"FileSize":uploaded_file.size}
        st.write('File Uploaded!')
    if st.button('Process'):
        st.write("Downloading of the data has started at the backend. The process may take 10 mins or longer depending upon the uploaded data size. Please have patience")
        centre = download_main(uploaded_file)
        st.write("Download Completed and Report Generation has started...")
        name = f"BoM EOL Stat_{centre}"
        with st.spinner('The Report is getting generated...'):
            data_path = "consolidate_BOM_report/BOM_EOL_Stat_full.csv"
            df = process_data(data_path)

            part_list = get_part_list(df)
            asset_to_pwa = get_asset_to_pwa(df)
            df = get_pwa_qty(df)
            pwa_bom = get_pwa_bom(df)
            asset_bom = get_asset_to_component(df)
            whereused = get_whereused(df)

            df = df.drop('Level_1',axis=1)
            shutil.rmtree("consolidate_BOM_report")

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
            linko= f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{name}.xlsx">Download Report file</a>'
            st.markdown(linko, unsafe_allow_html=True)
    
if __name__ == '__main__':
    main()