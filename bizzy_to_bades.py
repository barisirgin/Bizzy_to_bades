# -*- coding: utf-8 -*-
import xlrd
import time
import xlsxwriter as xls
import colorama
from colorama import Fore, Back, Style
colorama.init(autoreset=True)
                    
#İslem yapacagimiz xlsx dosyasini acar ve islem yapilacak tabloyu seccer
while True:
    file = input("Excel dosya adi : ")
    file2 = str(file) + ".xlsx"
    try: 
        allGroupsWorkbook = xlrd.open_workbook(file2)
        allGroupsSheet = allGroupsWorkbook.sheet_by_index(0) #Excel dosyasindaki 0'inci indexteki sayfaa islem yapicagimizi belirtiyoruz.
        break
    except FileNotFoundError:
        print(f"{Fore.RED}! ! ! Girdiginiz dosya adi hatali veya uygulama ile ayni dizinde bulunmuyor.Tekrar deneyin.")
        #file = input("Islem yapilacak excel dosyasini .xlsx uzantisi ile giriniz : ")
        
print("\n")
t1 = time.time()

#İstenilen formata donusturulmus excel dosyasini yaratir ve gerekli basliklari ekler
duzenlenmisWorkbook = xls.Workbook(f"{file}_(Duzenlenmis).xlsx") #Duzenlenmis  dosyayi olusturur.
duzenlenmisSheet = duzenlenmisWorkbook.add_worksheet("Bades") #Excel dosyasina tablo ismi verir

cell_format = duzenlenmisWorkbook.add_format()
cell_format.set_bg_color('#00B0F0')
titles = ["Ref. No","Bulgu Adı","Önem ","Etkisi ","Erişim Noktası ","Kullanıcı Profili ","Bulgunun Kategorisi ","Bulgunun Tespit Edildiği Bileşenler ","Bulgu Açıklaması ","Çözüm Önerisi ","Referanslar ","Tarih ","Durum ","İyileştirme-Ek kontrol önerisi ","Bulgu Cevabı/Aksiyon ","PluginID ","Port ","Images "]
x = 65
for title in titles:  
    duzenlenmisSheet.write(f"{chr(x)}1",title,cell_format)
    x += 1

for x in range(1,allGroupsSheet.nrows):#1 den baslamamizin sebebi basliklari alsin istemiyoruz
    duzenlenmisSheet.write(f"A{x + 1}", allGroupsSheet.cell_value(x, 0))
    duzenlenmisSheet.write(f"B{x + 1}", allGroupsSheet.cell_value(x, 12))
    duzenlenmisSheet.write(f"C{x + 1}", allGroupsSheet.cell_value(x, 9))
    duzenlenmisSheet.write(f"D{x + 1}", allGroupsSheet.cell_value(x, 10))
    duzenlenmisSheet.write(f"E{x + 1}", allGroupsSheet.cell_value(x, 14))
    #duzenlenmisSheet.write(f"F{x + 1}", allGroupsSheet.cell_value(x, 1))
    duzenlenmisSheet.write(f"G{x + 1}", allGroupsSheet.cell_value(x, 11))
    duzenlenmisSheet.write(f"H{x + 1}", "".join(allGroupsSheet.cell_value(x, 1)) + " - " + "".join(allGroupsSheet.cell_value(x, 5)))
    duzenlenmisSheet.write(f"I{x + 1}", allGroupsSheet.cell_value(x, 34))
    duzenlenmisSheet.write(f"J{x + 1}", allGroupsSheet.cell_value(x, 36))
    #duzenlenmisSheet.write(f"K{x + 1}", allGroupsSheet.cell_value(x, 1))
    duzenlenmisSheet.write(f"L{x + 1}", allGroupsSheet.cell_value(x, 39))
    duzenlenmisSheet.write(f"M{x + 1}", allGroupsSheet.cell_value(x, 20))
    duzenlenmisSheet.write(f"N{x + 1}", allGroupsSheet.cell_value(x, 35))
    #duzenlenmisSheet.write(f"O{x + 1}", allGroupsSheet.cell_value(x, 1))
    duzenlenmisSheet.write(f"P{x + 1}", allGroupsSheet.cell_value(x, 15))
    duzenlenmisSheet.write(f"Q{x + 1}", allGroupsSheet.cell_value(x, 7))
    #duzenlenmisSheet.write(f"R{x + 1}", allGroupsSheet.cell_value(x, 1))

#satir sayisi
print(f"{Fore.LIGHTYELLOW_EX}Bu excel sayfasinda {allGroupsSheet.nrows} satir bulunmaktadir")
#sutun sayisi
print(f"{Fore.LIGHTYELLOW_EX}Bu excel sayfasinda {allGroupsSheet.ncols} sutun bulunmaktadir" + "\n")    

t2 = time.time()
tm = t2- t1
tm = round(tm,2)

print(f"{Fore.CYAN}{file}_(Duzenlenmis).xlsx dosyasi basariyla olusturuldu.")
print(f"{Fore.CYAN}{tm} saniye sürdü." + "\n")
print(f"{Fore.RED}Cikmak icin ENTER tusuna basiniz")
duzenlenmisWorkbook.close()

input()

