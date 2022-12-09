# -*- coding: utf-8 -*-
"""
Created on Wed Aug 31 09:35:43 2022

@author: Mahmut Gazi Altun
"""
import PyPDF4
import re

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
            satir=page_content[int(kelime[0]):int(kelime[1])+30]
            liste=satir.splitlines()
            bolunensatir=str(liste[0])+" "+str(liste[1])
            indekslistesi.append(bolunensatir)
            print("Bölünen satır:",bolunensatir)
            return indekslistesi
        
def sonucyazdir(indekslistesi):  
    for indeks in indekslistesi:
        #print("indeks listesi:",indekslistesi)
        bolunensatir=indeks.split(" ")
        if(
            bolunensatir[1]=="Power"or
            bolunensatir[1]=="ratio"or
            bolunensatir[1]=="Energy"or
            bolunensatir[1]==":"or
            bolunensatir[0]==""):
            print(bolunensatir[0],bolunensatir[1],":",bolunensatir[2])
            basliklar.append(bolunensatir[0]+" "+bolunensatir[1])
            degerler.append(bolunensatir[2])
        elif bolunensatir[0]=="Project":
            print(bolunensatir[0],":",bolunensatir[1])
        
        else:
            print(bolunensatir[0],bolunensatir[1],bolunensatir[2],":",bolunensatir[3])
            
if __name__ =="__main__":
    
    kelimeler=["Produced Energy","Unit Nom. Power"]
    aranankelimeler=[]
    indekslistesi=[]
    basliklar=[]
    degerler=[]
    pdfFileObj=open('C:\\Users\\Mahmut Gazi Altun\\Desktop\\Pvsyst\\220606_GTC_Bifi335_Hua_GB_1.pdf', 'rb')
    reader=PyPDF4.PdfFileReader(pdfFileObj)
    count = reader.numPages
    print("sayfa sayısı:",count)
    page_content = ""
    
    for i in range(count):
        pageObj = reader.getPage(i)
        page_content += pageObj.extractText()
    print("-----------") 
   
sonucyazdir(kelimeara(kelimeler))