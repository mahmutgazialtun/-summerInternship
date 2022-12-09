# -*- coding: utf-8 -*-
"""
Created on Wed Jul 20 10:16:50 2022

@author: Mahmut Gazi Altun
"""

import PyPDF2 
import re    

# pdf'nin nesnesini oluşturma
pdfFileObj = open(
'C:\\Users\\Mahmut Gazi Altun\\Desktop\\pdfler\\20201021_Gunam_Bifi.pdf', 'rb') 
    
# pdf okumak için nesne 
pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
    
# sayfa sayısını yazdırma
print("sayfa sayısı:",pdfReader.numPages) 
    
# sayfanın nesnesini oluşturma 
pageObj = pdfReader.getPage(3) 
 
page_content = pageObj.extractText()

# sayfadan metin çıkarma
print(pageObj.extractText()) 

print("*******************************************") 

#kelime arama
aranan=re.search('azimuth', page_content) # aranan kelimenin yeri
aranan1=re.search('Pnom', page_content)
aranan2=re.search("Pnom total",page_content)

print(aranan)
print(aranan1)
print(aranan2)
#print("aranan kelime:",page_content[334:337])

print("**************************")

print("aranan veri:",page_content[326:333],":", page_content[334:337])
print(page_content[389:393] ,":",page_content[394:400])
print(page_content[427:437],":",page_content[438:446])

# pdf nesnesini kapatma 
pdfFileObj.close() 
