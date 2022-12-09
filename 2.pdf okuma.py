# -*- coding: utf-8 -*-
"""
Created on Mon Jul 18 15:06:18 2022

@author: Mahmut Gazi Altun
"""

import PyPDF2 
    
# okunacak pdf nesnesi
pdfFileObj = open
('C:\\Users\\Mahmut Gazi Altun\\Downloads\\20201021_Gunam_Mono.pdf', 'rb') 
    
# pdf okuma nesnesi
pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
    
# yazdırma 
print(pdfReader.numPages) 
    
# sayfa nesnesi 
pageObj = pdfReader.getPage(0) 
    
# pdf'ten veri çekme 
print(pageObj.extractText()) 
    
# pdf okuma nesnesini kapatma 
pdfFileObj.close() 
