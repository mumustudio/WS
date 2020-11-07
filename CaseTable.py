#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Apr 26 23:27:31 2019

@author: ws
"""
from bs4 import BeautifulSoup
import re
import multiprocessing as mp
import time
from openpyxl import Workbook
import requests
from openpyxl import load_workbook

table_key_words = ['commitments(\s\n*)and(\s\n*)contingencies','commitments(\s\n*)and(\s\n*)guarantees',
                   'leasing(\s\n*)arrangements','commitments(\s\n*)and(\s\n*)contingent(\s\n*)liabilities']

money_key_words = ['millions','thousands']

find_key_words = ['long-term','long term','long-run','long run',
                  'look ahead','looking ahead','outlook','short-run',
                  'short run','short-term','short term']

def getDataUrls():
    wb = load_workbook('DataUrls.xlsx')
    sheet = wb['Sheet']
    ciks=[]
    urls=[]
    calyears=[]
    
    for cik in sheet['A']:
        ciks.append(cik.value)
    for calyear in sheet['B']:
        calyears.append(calyear.value)
    for url in sheet['C']:
        urls.append(url.value)
    return ciks,urls,calyears

def getErrorCIK():
    wb = load_workbook('ErrorCIK.xlsx')
    sheet = wb['Sheet']
    error_ciks=[]
    
    for error_cik in sheet['A']:
        error_ciks.append(error_cik.value)
    return error_ciks

def crawl(url):
    response = requests.get(url)
    time.sleep(0.1)
    return response.content

def getTable(html):
    soup = BeautifulSoup(html,'lxml')
    soup3 = 'no table'
    
    #if re.findall('(.*)form(.*)10-k(.*)',str(soup).lower()):
    try:
        if re.search('contractual(.*)obligations',str(soup).lower()):
            soup2 = BeautifulSoup(str(soup)[re.search('contractual(.*)obligations',str(soup).lower()).span()[0]:],'lxml')
            soup3 = BeautifulSoup(str(soup2)[:re.search('</table>',str(soup2)).span()[1]],'lxml')
        else:
            for word in table_key_words:
                if re.search(word,str(soup).lower()):
                    soup2 = BeautifulSoup(str(soup)[re.search(word,str(soup).lower()).span()[0]:],'lxml')
                    soup3 = BeautifulSoup(str(soup2)[:re.search('</table>',str(soup2)).span()[1]],'lxml')
                    break
        excel_datas=[]
        if soup3 != 'no table':
            if soup3.find_all('tr'):
                for tr in soup3.find_all('tr'):
                    if tr.find_all('th'):
                        row = [th.text for th in tr.find_all('th')]
                    elif tr.find_all('td'):
                        row = [td.text for td in tr.find_all('td')]
                    excel_datas.append(row)
            else:
                row = ['no data']
                excel_datas.append(row)
        else:
            row = ['no data']
            excel_datas.append(row)
    except:
        row = ['no data']
        excel_datas.append(row)
    return excel_datas
    
def getMoneyUnit(html):
    result=''
    soup = BeautifulSoup(html,'lxml')

    #if re.findall('(.*)form(.*)10-k(.*)',str(soup).lower()):
    try:
        for word in money_key_words:
            if re.findall(word,str(soup).lower()):
                result+=(','+word)
        if result == '':
            result = 'no data'
    except:
        result = 'no data'
    return result

def getWordCount(html):
    soup = BeautifulSoup(html,'lxml')
    result=[]
    try:
        for word in find_key_words:
            result.append(len(re.findall(word,str(soup).lower())))
    except:
        for word in find_key_words:
            result.append(0)
    return result

if __name__ == '__main__':
    

    mp.freeze_support()
    pool = mp.Pool()
    
    wb = Workbook()
    ws = wb.active
    wb2 = Workbook()
    ws2 = wb2.active
    
    cik_index=0
    size_of_row=1
    AllTable_index=0
    AllWordCount_index=0
    
    print('program start.')
    all_ciks,all_urls,all_calyears = getDataUrls()
    error_ciks = getErrorCIK()
    print('read excel done.')
    
    ws2.cell(row=1, column=1).value = 'CIK'
    ws2.cell(row=1, column=2).value = 'CalYear'
    ws2.cell(row=1, column=3).value = '網址'
    for col,word in enumerate(find_key_words):
        ws2.cell(row=1, column=col+4).value = word
    
    while len(all_urls) !=0:
        urls=[]
        if len(all_urls) < 20:
            for times in range(0,len(all_urls)):
                urls.append(all_urls.pop(0))
        else:
            for times in range(0,19):
                urls.append(all_urls.pop(0))
        
        print('crawl start.')
        crawl_jobs = [pool.apply_async(crawl,args=(url,)) for url in urls]
        htmls = [j.get() for j in crawl_jobs]
        print('crawl done.')
        
        print('getTable start.')
        getTable_jobs = [pool.apply_async(getTable,args=(html,)) for html in htmls]
        all_excel_datas = [j.get() for j in getTable_jobs]
        print('getTable done.')
        
        print('getMoneyUnit start.')
        getMoneyUnit_jobs = [pool.apply_async(getMoneyUnit,args=(html,)) for html in htmls]
        moneyUnits = [j.get() for j in getMoneyUnit_jobs]
        print('getMoneyUnit done.')
        
        print('getWordCount start.')
        getWordCount_jobs = [pool.apply_async(getWordCount,args=(html,)) for html in htmls]
        all_word_counts = [j.get() for j in getWordCount_jobs]
        print('getWordCount done.')
        
        print('go to excel start.')
        for i,excel_datas in enumerate(all_excel_datas):
            ws.cell(row=size_of_row, column=1).value = 'CIK:'
            ws.cell(row=size_of_row, column=2).value = str(all_ciks[AllTable_index])
            ws.cell(row=size_of_row, column=3).value = 'Calyear:'
            ws.cell(row=size_of_row, column=4).value = str(all_calyears[AllTable_index])
            ws.cell(row=size_of_row, column=5).value = '金錢單位:'
            ws.cell(row=size_of_row, column=6).value = moneyUnits[i]
            ws.cell(row=size_of_row, column=7).value = urls[i]
            AllTable_index+=1
            size_of_row+=1
            for row,datas in enumerate(excel_datas):
                for column,data in enumerate(datas):
                    ws.cell(row=row+size_of_row, column=column+2).value = data.replace('\n','')
            size_of_row+=len(excel_datas)
        
        for i,wordCounts in enumerate(all_word_counts):
            ws2.cell(row=AllWordCount_index+2, column=1).value = str(all_ciks[AllWordCount_index])
            ws2.cell(row=AllWordCount_index+2, column=2).value = str(all_calyears[AllWordCount_index])
            ws2.cell(row=AllWordCount_index+2, column=3).value = str(urls[i])
            AllWordCount_index+=1
            for col,count in enumerate(wordCounts):
                ws2.cell(row=AllWordCount_index+1, column=col+4).value = str(count)
        
            
    print('program done.')      
    size_of_row+=1
    ws.cell(row=size_of_row, column=1).value = '以下CIK用自動化抓不到資料（沒有10-k資料）'
    size_of_row+=1
    for i,error_cik in enumerate(error_ciks):
        ws.cell(row=i+size_of_row+1, column=1).value = str(error_cik)
            
    wb.save(filename='AllTable.xlsx')
    wb2.save(filename='AllWordCount.xlsx')
