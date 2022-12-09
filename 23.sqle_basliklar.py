# -*- coding: utf-8 -*-
"""
Created on Wed Aug 17 08:24:17 2022

@author: Mahmut Gazi Altun
"""
import PyPDF2 
import re    
import xlsxwriter
import pyodbc
import pandas as pd
import numpy as np

def kelimeara(kelimeler):
    
    for kelime in kelimeler:
        print(kelime)
        if kelime in aranankelimeler:
            print("bu kelime arandı")
        else:
            aranankelimeler.append(kelime)
            inds=[s.start() for s in re.finditer(kelime, page_content)]
            inde=[s.end() for s in re.finditer(kelime, page_content)]
            kelime=inds[0],inde[0]
            satir=page_content[int(kelime[0]):int(kelime[1])+80]
            liste=satir.splitlines()
            bolunensatır=liste[0]
            indekslistesi.append(bolunensatır)
    print("----------------")
    return indekslistesi

def exceleYaz(indekslistesi):
    planWorkbook = xlsxwriter.Workbook("C:\\Users\Mahmut Gazi Altun\Desktop\PDFtoExcel\Python\donusen6.xlsx")
    planSheet = planWorkbook.add_worksheet("sonuclar")              
    planSheet.write("A1","Başlıklar")
    planSheet.write("B1","Değerler") 
    j=0
    k=0
    counter=0
    for indeks in indekslistesi:
        bolunensatir=indeks.split(" ",4)
        if(bolunensatir[1]=="prod." or
           bolunensatir[1]=="Power"or
           bolunensatir[1]=="ratio"or
           bolunensatir[1]=="Energy"or
           bolunensatir[1]==":" or
           bolunensatir[1]==" "):
            print(bolunensatir[0],bolunensatir[1],":",bolunensatir[2])
            basliklar.append(bolunensatir[0]+" "+bolunensatir[1])
            degerler.append(bolunensatir[2])
        else:
            print(bolunensatir[0],":",bolunensatir[1])
            basliklar.append(bolunensatir[0])
            degerler.append(bolunensatir[1])
                        
    for b in basliklar:
        j=j+1
        planSheet.write(j,k,b)
        planSheet.write(j,k+1,degerler[counter])
        counter=counter+1
                    
    planWorkbook.close()
    return basliklar

def insertSql(basliklar):
     conn = pyodbc.connect(
         "Driver={SQL Server Native Client 11.0};"
         "Server=(localdb)\MSSQLLocalDB;"
         "Database=TestDb;"
         "Trusted_Connection=yes")   
     counter=0
     try:
         for i in basliklar:
             sql='''insert into test(baslik,deger) values(?,?)'''
             cursor=conn.cursor()
             data=str(i).split(" ")             
             cursor.execute(sql,data)
             conn.commit()
         return True
     except Exception as e:
         print("hata:{}".format(e))
         return False
     finally:
         cursor.close()
         
if __name__ =="__main__":
    
    kelimeler=["Project","Specific prod.","Total Power","Produced Energy","Pnom"]
    aranankelimeler=[]
    indekslistesi=[]
    basliklar=[]
    degerler=[]
    yazilanbaslik=[] 
    yazilandegerler=[]         
    pdfFileObj = open('C:\\Users\\Mahmut Gazi Altun\\Desktop\\Pvsyst\\20201021_GTC_Solar Çatı_Bifacial.pdf.pdf', 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
    count = pdfReader.numPages
    print("sayfa sayısı:",count)

    page_content = ""
    for i in range(count):
        pageObj = pdfReader.getPage(i)
        page_content += pageObj.extractText()
    print("------------------------------")

insertSql(exceleYaz(kelimeara(kelimeler)))