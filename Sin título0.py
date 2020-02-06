# -*- coding: utf-8 -*-
"""
Created on Tue Feb  4 18:37:08 2020

@author: Cesar
"""
import openpyxl

wb = openpyxl.load_workbook('prueba.xlsx')
ws = wb.active
hoja=wb['Hoja1']

data = [
    ["Fruit", "Quantity"],
    ["Kiwi", 3],
    ["Grape", 15],
    ["Apple", 3],
    ["Peach", 3],
    ["Pomegranate", 3],
    ["Pear", 3],
    ["Tangerine", 3],
    ["Blueberry", 3],
    ["Mango", 3],
    ["Watermelon", 3],
    ["Blackberry", 3],
    ["Orange", 3],
    ["Raspberry", 3],
    ["Banana", 3]
]

lista=[]
for r in data:
    ws.append(r)

ws.auto_filter.ref = "A:B"
ws.auto_filter.add_filter_column(0, ["Kiwi", "Apple", "Mango"])

for i in hoja.iter_rows():
    lista.append(i) 
print(lista)
#ws.auto_filter.add_sort_condition("B2:B15")

wb.save("filered.xlsx")