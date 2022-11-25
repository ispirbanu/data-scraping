from tkinter import *
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, 
NavigationToolbar2Tk)

#excelden belirli bazı sütünları almak hata-1
#Urunlerin bulunduğu data
data1=pd.read_excel('veriler_.xlsx',sheet_name='Urunler',usecols=('B,C,D,E,F'))

#Yorumların bulunduğu data
data2=pd.read_excel('veriler_.xlsx',sheet_name='Yorumlar',usecols=('B,C,D,E'))


#sıralamaları algoritmayla yapmak için
oran1=data1['Positive Rate'].tolist()
oran2=data1['Negative Rate'].tolist()

#pencerenin oluşturulması
root = tk.Tk()
root.geometry('1000x600')
root.resizable(False, False)
root.title('YORUM FİLTRELEME')


#grafiği oluşturma
def plot(show_r,positive,negative):
    #olumlu,olumsuz yorum sayılarını ve ortalamalarını label üzerinen göstermek.
    yaz="Olumlu Puan Yüzdesi: " + str((positive/(positive+negative))*100)
    yaz2="Olumsuz Puan Yüzdesi: " + str((negative/(positive+negative))*100)
    yaz3="Olumlu Yorum Sayısı: "+str(positive)
    yaz4="Olumsuz Yorum Sayısı: "+str(negative)
    labels2.pack(padx=5, pady=5)
    labels1.pack(padx=5, pady=5)
    labels3.pack(padx=5, pady=5)
    labels4.pack(padx=5, pady=5)
    labels1.config(text=yaz)
    labels2.config(text=yaz2)
    labels3.config(text=yaz3)
    labels4.config(text=yaz4)
    
    #histogram çizmek
    #aralığı 1-5 olan kenar renkleri siyah, arkaplanı kırmızı olan histogram
    ax1.hist(show_r,range=[1,5], edgecolor='black',facecolor='red')
    #historamın x verileri puanlamalar
    plt.xlabel("Puan")
    #histogramın y verileri kişi sayısını temsil etmekte
    plt.ylabel("Kişi Sayısı")
    #canvas ile cizdirme işlemi yapılıyor.
    canvas.draw()
    #canvas üzerinden pencerede konumlandırılması yapılıyor
    canvas.get_tk_widget().pack(side=tk.LEFT)
    
    
def get_review(event):
    
    global sayac
    #scrollbarlar tekrar oluşturulmaması için
    sayac+=1
    global ax1
    #histogram için puanlamaları tutan liste
    show_rate=list()
    positive=0
    negative=0
    
    #gerekli verileri ve comboboxdan seçilen kategori
    kategorial=kategori_cb.get()
    # excel yorumlar sayfasındaki kategoriler
    kategoriler=data2['Category'].tolist()
    #yorumlar
    yorumlar=data2['body'].tolist()
    #excel yorumlar sayfasındaki modeller
    modeller=data2['Model'].tolist()
    #puanlar
    puanlar=data2['rating'].tolist()
    
     #her severinde yorumların yeni yüklenmesi için listbox temizlenmeli
    listbox2.delete(0,'end')
    
    #evet olayı yani butonla iletilen model combobox'da bulunan değer
    marka = event.get()
    
    #scrollbarların sadece bir kere oluşturulması için
    if sayac==0:
        #scrollbar oluşturmak
        scrollbar=Scrollbar(root)
        #y yönünde scrollbar konumunu ve boyutunu belirlemesi
        scrollbar.pack(side=RIGHT, fill='y')
        #x yönünde hareket edecek scrollbar için orient özelliği ve konum bout belirlemesi
        scrollbarx = Scrollbar(root, orient=HORIZONTAL)
        scrollbarx.pack(side=BOTTOM, fill='x')
        
        #scrollbarların yorum listbox'ına dahil edilmesi
        scrollbar.config(command=listbox2.yview)
        scrollbarx.config(command=listbox2.xview)
        
        #listboxın scrollbarlarına dahil edilmesi
        listbox2.config(yscrollcommand=scrollbar.set)
        listbox2.config(xscrollcommand=scrollbarx.set) 
    ka=len(yorumlar)
    #listbox'ı görüntülemek
    listbox2.pack( fill=tk.BOTH,side=tk.RIGHT,padx=5, pady=5)
    
    for i in range(ka):
        #kategoriler uyuşuyorsa
        if kategorial in kategoriler[i]:
            #modeller uyuşuyorsa
            if marka in modeller[i]:
                #yorum ve puanları göster
                strp="Verilen Puan: "+str(puanlar[i])
                listbox2.insert(END, yorumlar[i])
                listbox2.insert(END, strp)
                listbox2.insert(END,"\n")
                #çizim işin puanlamaların tutulduğu liste
                show_rate.append(puanlar[i])
                #yorumların sınıfını bulmak 
                if puanlar[i]>3:
                    positive+=1
                if puanlar[i]<3:
                    negative+=1
   #scrollbar yeniden çizilmemesi için
    sayac+=1
    #yeni değerler ekleneceği için grafiği temizleme
    ax1.clear()
    #plot fonksiyonuna değerleri gönderek yazdırma işlemleri
    plot(show_rate,positive,negative) 
    



        
#modelleri tutacak liste
def kategori_changed(event):
     kategorileri_al=kategori_cb.get()
   
    #Modellerin listesi
     model_listele=data1['Model'].tolist()
    #kategorilerin listesi
     kategori_listele=data1['Category'].tolist()
     dene=list()
     dene.clear()
     for i in range(len(kategori_listele)):
         if kategorileri_al in kategori_listele[i]:
             dene.append(model_listele[i])
     label.pack(fill='x', padx=5, pady=5)
    #Combobox'a verileri yükleme 
     model_cb['values'] = dene #verileri comboboxa atama 
     model_cb['state'] = 'readonly'  
     #combobox ve butonu görüntüle
     model_cb.pack(fill='x', padx=5, pady=5)
     buton.pack(fill='x',padx=5,pady=5)

    
    
    
#kategorileri çekmek (her kategori bir kez alınmalı)
kategorilistele=data1['Category'].tolist()
kategori=[]
for i in kategorilistele: 
    if i in kategori:#önceden varsa bu adımı atla
        pass
    else:
        kategori.append(i)#yoksa listeye ekle


#scrollbar tekrarlanmaması için oluşturulan sayaç
sayac=-1

#buton oluşturma
buton=Button(root,text="Yorumları Göster",bg='darkgrey',height=2,command=lambda: get_review(model_cb))
#yorumlar Listbox'ı
listbox2=tk.Listbox(root,width=85,height=20,bg='lightgrey')

#verileri ortalamaya göre sıralama hata-2
data1=data1.set_index('Positive Rate')
data1=data1.sort_index(ascending=False)

#kategori label
label = ttk.Label(text="Kategori seçiniz:")
label.pack(fill='x', padx=5, pady=5)

#Tkinter StringVar ile Label gibi araçları etkili bir şekilde yönetmeye yardımcı olur .
#kullanılması gerekmiyor.
stringvar= tk.StringVar()


#kategorileri combobox'a eklemek
kategori_cb=ttk.Combobox(root, textvar=stringvar)
kategori_cb['values']=kategori
kategori_cb['state']='readonly' #sadece okunur yapmak
kategori_cb.pack(fill='x', padx=5, pady=10) #x ve y için boşluklar ve x eksenini doldurma


#model label
label = ttk.Label(text="Model seçiniz:")


#Modellerin listesi
model_listele=data1['Model'].tolist()

#model combobox
model_cb=ttk.Combobox(root)


#ürünler listbox'ı
#listbox=tk.Listbox(root,width=85,height=20,bg='lightgrey')




#grafiği oluşturmak için boyutlandırma 
fig=plt.figure(figsize=(6,4))
ax1=fig.add_subplot()
#canvas ile grafiği çizmek
canvas=FigureCanvasTkAgg(fig,master=root)

#olumlu olumsuz yorum sayısı ve ortalamaların yazılacağı labellar
labels1 = ttk.Label(text="")
labels2 = ttk.Label(text="")
labels3 = ttk.Label(text="")
labels4 = ttk.Label(text="")


#kategori değiştiğinde çalıştırılması gereken fonksiyonu çağırmak
k=kategori_cb.bind('<<ComboboxSelected>>',kategori_changed)

#arayüzün ekranda kalması için
root.mainloop()

