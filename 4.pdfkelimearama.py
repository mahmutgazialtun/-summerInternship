# -*- coding: utf-8 -*-
"""
Created on Wed Jul 20 09:27:11 2022

@author: Mahmut Gazi Altun
"""

import PyPDF2 
import re    

# pdf'nin nesnesini oluşturma
pdfFileObj = open('C:\\Users\\Mahmut Gazi Altun\\Desktop\\deneme.pdf', 'rb') 
    
# pdf okumak için nesne 
pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
    
# sayfa sayısını yazdırma
print("sayfa sayısı:",pdfReader.numPages) 
    
# sayfanın nesnesini oluşturma 
pageObj = pdfReader.getPage(0) 
 
page_content = pageObj.extractText()
# sayfadan metin çıkarma
print(pageObj.extractText()) 
    
#kelime arama
aranan=re.search('belge', page_content)
print(aranan)
print("aranan kelime:",page_content[43:48])

# pdf nesnesini kapatma 
pdfFileObj.close() 