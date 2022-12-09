# -*- coding: utf-8 -*-
"""
Created on Tue Aug 23 10:41:16 2022

@author: Mahmut Gazi Altun
"""

import PyPDF2 
import re    
import xlsxwriter
import pyodbc
import pandas as pd

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

    planWorkbook = xlsxwriter.Workbook("C:\\Users\Mahmut Gazi Altun\Desktop\PDFtoExcel\Python\donusen8.xlsx")
    planSheet = planWorkbook.add_worksheet("sonuclar")                 
    planSheet.write("A1","Başlıklar")
    planSheet.write("B1","Değerler") 
    j=0
    k=0
    counter=0
    for indeks in indekslistesi:
        bolunensatir=indeks.split(" ",4)
        if(bolunensatir[1]=="prod."or
           bolunensatir[1]=="Power"or
           bolunensatir[1]=="ratio"or
           bolunensatir[1]=="Energy"or
           bolunensatir[1]==":"or
           bolunensatir[1]==" "):
            print(bolunensatir[0],bolunensatir[1],":",bolunensatir[2])
            basliklar.append(bolunensatir[0]+" "+bolunensatir[1])
            degerler.append(bolunensatir[2])            
        else:
            print(bolunensatir[0],":",bolunensatir[1])
            basliklar.append(bolunensatir[0])
            degerler.append(bolunensatir[1])
                        
    for b in basliklar:
        b=b.strip(":")        
        planSheet.write(j,k,b)
        planSheet.write(j+1,k,degerler[counter])
        counter=counter+1        
        k=k+1            
    planWorkbook.close()
    
    return degerler
def tablobirlestir():
    df1=pd.read_excel("C:\\Users\Mahmut Gazi Altun\Desktop\PDFtoExcel\Python\donusen7.xlsx")

    df2=pd.read_excel("C:\\Users\Mahmut Gazi Altun\Desktop\PDFtoExcel\Python\donusen8.xlsx")
    df1=pd.concat([df1,df2])
    df1.to_excel("C:\\Users\Mahmut Gazi Altun\Desktop\PDFtoExcel\Python\donusen7.xlsx",index=False)
    
def insertSql(degerler):
     conn = pyodbc.connect(
         "Driver={SQL Server Native Client 11.0};"
         "Server=(localdb)\MSSQLLocalDB;"
         "Database=TestDb;"
         "Trusted_Connection=yes") 
     try:
         cursor=conn.cursor()                 
         sql='''insert into test2(Project,Specific_prod,Total_Power,Produced_Energy,Pnom_ratio) values(?,?,?,?,?)'''         
         cursor.execute(sql,degerler)         
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
                            # 20201021_Gunam_Mono
                            # 20201021_Tiryaki_Aypi1
                            # 20201021_TYT_
    pdfFileObj = open('C:\\Users\\Mahmut Gazi Altun\\Desktop\\Pvsyst\\20201021_Tiryaki_Aypi3.pdf', 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
    count = pdfReader.numPages
    print("sayfa sayısı:",count)
    page_content = ""
    for i in range(count):
        pageObj = pdfReader.getPage(i)
        page_content += pageObj.extractText()
    print("------------------------------")

insertSql(exceleYaz(kelimeara(kelimeler)))
tablobirlestir()