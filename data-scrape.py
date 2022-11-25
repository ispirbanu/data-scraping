from tkinter import *
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, 
NavigationToolbar2Tk)

#pencere oluşturmak
root = tk.Tk()
root.geometry('1000x600')
root.resizable(False, False)
root.title('YORUM FİLTRELEME')
#excelden belirli bazı sütünları almak hata-1
data1=pd.read_excel('veriler_.xlsx',sheet_name='Urunler',usecols=('B,C,D,E,F'))

#ürüne göre yorum seçmeye çalışmak
data2=pd.read_excel('veriler_.xlsx',sheet_name='Yorumlar',usecols=('B,C,D,E'))



    

show_rate=list()
#resim boyutunu ayarlamak
fig=plt.figure(figsize=(6,4))
ax1=fig.add_subplot()
ax1.grid()
canvas=FigureCanvasTkAgg(fig,master=root)
labels1 = ttk.Label(text="")
labels2 = ttk.Label(text="")
labels3 = ttk.Label(text="")
labels4 = ttk.Label(text="")

def plot(show_r,positive,negative):
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
    print(show_r)
    print("pos say: ",positive)
    print("neg say: ", negative)
    yar=[1,2,3,4,5]
   # ax1.clear()
   #histogramın aralığı, parametre kenarlık rengi siyah ve arkaplan rengi kırmızı olarak
   #histogramı oluşturur.
    ax1.hist(show_r,range=[1,5], edgecolor='black',facecolor='red')
    plt.xlabel("Puan")
    plt.ylabel("Kişi Sayısı")
    #canvas ile grafiği çiz.
    canvas.draw()
    #pencerenin soluna yerleştir.
    canvas.get_tk_widget().pack(side=tk.LEFT)
    #temizle(labels)


def get_review(event):

    global sayac,ax1
    global show_rate
    positive=0
    negative=0
    show_rate.clear()
    sayac+=1
    marka = event.get()
    kategorial=kategori_cb.get()
    print(kategorial)
    kategoriler=data2['Category'].tolist()
    yorumlar=data2['body'].tolist()
    modeller=data2['Model'].tolist()
    puanlar=data2['rating'].tolist()
    ka=len(yorumlar)
    
    
    listbox.delete(0,'end')
    if sayac>0:
        pass
    if sayac==0:
        
        scrollbar=Scrollbar(root)
        scrollbar.pack(side=RIGHT, fill='y')
        scrollbarx = Scrollbar(root, orient=HORIZONTAL)
        scrollbarx.pack(side=BOTTOM, fill='x')
        listbox.config(yscrollcommand=scrollbar.set)
        listbox.config(xscrollcommand=scrollbarx.set)
        scrollbar.config(command=listbox.yview)
        scrollbarx.config(command=listbox.xview)
    listbox.pack( fill=tk.BOTH,side=tk.RIGHT,padx=5, pady=5)
    #print(marka)
    
    for i in range(ka):
        if kategorial in kategoriler[i]:
            if marka in modeller[i]:
                #show_reviews.append(yorumlar[i])
                s=yorumlar[i]
                strp="Verilen Puan: "+str(puanlar[i])
                listbox.insert(1,"\n")
                listbox.insert(1,strp)
                listbox.insert(1,yorumlar[i])
                show_rate.append(puanlar[i])
                if puanlar[i]>3:
                    positive+=1
                if puanlar[i]<3:
                    negative+=1
    ax1.clear()
    plot(show_rate,positive,negative)         
    
   #create listbox
    
    #label=ttk.Label(text="Yorumlar:")
  
    

def kategori_changed(event):
    
    
    kategorileri_al=kategori_cb.get()
    #selected_month2 = tk.StringVar()
    #label = ttk.Label(text="Model seçiniz:")
    #Modellerin listesi
    model_listele=data1['Model'].tolist()
    #kategorilerin listesi
    kategori_listele=data1['Category'].tolist()
    dene=list()
    dene.clear()
    print(kategorileri_al)
    for i in range(len(kategori_listele)):
        if kategorileri_al in kategori_listele[i]:
            dene.append(model_listele[i])
    label.pack(fill='x', padx=5, pady=5)
    #month_cb = ttk.Combobox(root, textvariable=selected_month2)
    month_cb['values'] = dene #verileri comboboxa atama 
    month_cb['state'] = 'readonly'  
    month_cb.pack(fill='x', padx=5, pady=5)
    
   # month_cb.bind('<<ComboboxSelected>>', month_changed)
   
    #buton=Button(root,text="Yorumları Göster",bg='darkgrey',height=2,command=lambda: get_review(month_cb))
    #buton=Button(root,text="Ara",command=partial(get_review(month_cb))
    buton.pack(fill='x',padx=5,pady=5)
   
    

'''
wb=load_workbook('veriler.xlsx')
ws=wb.get_sheet_by_name(name='Urunler') 
#data=wb.set_index('Postive Rate')
#data=data.sort_values(by='Positive Rate')
#data=sort_index(ascending=False)
col_a=ws['C']
col_b=ws['E']
modeller=list()
for item in col_a:
    modeller.append(item.value)

print(modeller)
#Velileri büyükten küçüğe sıralayarak çekmek...
'''

#data=pd.ExcelFile("veriler.xlsx")
'''
#excelden belirli bazı sütünları almak hata-1
data=pd.read_excel('veriler.xlsx',sheet_name='Urunler',usecols=('C,D,E'))

#ürüne göre yorum seçmeye çalışmak
data2=pd.read_excel('veriler.xlsx',sheet_name='Yorumlar',usecols=('C,D,E'))
#data2=data2.set_index('body')'''

#Yorumları filtreleme
'''yorum1=data2['body'].tolist()
yorum2=data2['Model'].tolist()
ka=len(yorum2)

e='Xiaomi Redmi 8 64GB Yeşil Cep'
for i in range(ka):
    if e in yorum2[i]:
        print(yorum1[i])
#yorum filtreleme bitti      
'''
    
'''a=list(zip(yorum2,yorum1))

ka=len(yorum1)
for items in range(ka):
    if yorum2[items] in e:
        print(items)
#aa=[item[0] for item in a]
'''

#print("DATA VALUES")
#print(data.values)
#dene=data.values.tolist()


#3. sorun çözümü dataframe den listeye dönüşüm
#dene=data['Model'].tolist()

sayac=-1


#verileri ortalamaya göre sıralama hata-2
data1=data1.set_index('Positive Rate')
#data=data.sort_values(by='Positive Rate')
data1=data1.sort_index(ascending=False)

#combobox 
selected_month = tk.StringVar()
selected_month2 = tk.StringVar()
selected_month3 = tk.StringVar()

#listbox oluşturma
listbox=Listbox(root,width=85,height=20,bg='lightgrey')

buton=Button(root,text="Yorumları Göster",bg='darkgrey',height=2,command=lambda: get_review(month_cb))
#kategori için combobox

label = ttk.Label(text="Kategori seçiniz:")
label.pack(fill='x', padx=5, pady=5)
kategori=["Telefon",
          "Tablet",
          "AkıllıSaat",
          ]

kategori_cb=ttk.Combobox(root, textvar=selected_month)
kategori_cb['values']=kategori
kategori_cb['state']='readonly'
kategori_cb.pack(fill='x', padx=5, pady=10)
label = ttk.Label(text="Model seçiniz:")
month_cb = ttk.Combobox(root, textvariable=selected_month2)

k=kategori_cb.bind('<<ComboboxSelected>>',kategori_changed)
c=kategori_cb.get()

#kategori sonu




# create a combobox
'''

label = ttk.Label(text="Model seçiniz:")
label.pack(fill='x', padx=5, pady=5)


month_cb = ttk.Combobox(root, textvariable=selected_month)
month_cb['values'] = dene #verileri comboboxa atama 
month_cb['state'] = 'readonly'  
month_cb.pack(fill='x', padx=5, pady=25)

month_cb.bind('<<ComboboxSelected>>', month_changed)

buton=Button(root,text="Ara",command=lambda: get_review(month_cb,kategori_cb))
#buton=Button(root,text="Ara",command=partial(get_review(month_cb))
buton.pack(fill='x',padx=5,pady=30)

'''

root.mainloop()


'''
root=tk.Tk()
root.geometry("500x500")

deneme2=Listbox(root,width=35)
deneme2.pack(pady=20)
wb=load_workbook('veriler.xlsx')
ws=wb.get_sheet_by_name(name = 'Urunler') 
col_a=ws['C']
col_b=ws['E']

for item in col_a:
    deneme2.insert(END,item.value)
root.mainloop()

#Combobox oluşturma 
#Verileri sheete göre çekme 

liste=[]
root= Tk()
root.geometry("300x300")

def selected():
    myLabel=Label(root,text=clicked.get()).pack()
    
a=[]

a=pd.read_excel('veriler.xlsx',sheet_name='Urunler')
data2=pd.read_excel('veriler.xlsx',sheet_name='Yorumlar')


options=[
    "monday",
    "tuesday",
    "wednesday",
    "thursday",
    "friday"
    ]

clicked=StringVar()
#clicked.set(options[0])
drop=OptionMenu(root, clicked, *options)
drop.pack(pady=20)

myButton=Button(root,text="select",command=selected)

myButton.pack()
#print(a)
print(liste)
root.mainloop()
'''