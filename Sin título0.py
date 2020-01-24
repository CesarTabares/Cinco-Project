# -*- coding: utf-8 -*-
"""
Created on Thu Jan 23 19:44:17 2020

@author: Cesar
"""

for i in range (len(lista_fecha_actuaciones)):
        
        
        main_sheet.cell(row=(empty_row+i),column=3).value=lista_fecha_actuaciones[i]
        main_sheet.cell(row=(empty_row+i),column=7).value=lista_actuaciones[i]
        main_sheet.cell(row=(empty_row+i),column=11).value=lista_anotaciones[i]
        main_sheet.cell(row=(empty_row+i),column=37).value=lista_fecha_inicia[i]
        main_sheet.cell(row=(empty_row+i),column=41).value=lista_fecha_termina[i]
        main_sheet.cell(row=(empty_row+i),column=45).value=lista_fecha_registro[i]
        main_sheet.row_dimensions[empty_row+i].height = 33
        
        main_sheet.cell(row=(empty_row+i), column=3)._style=deepcopy(main_sheet[style_source]._style)
        main_sheet.merge_cells(start_row=empty_row+i, start_column=3, end_row=empty_row+i, end_column=3+3)
        main_sheet.cell(row=(empty_row+i), column=7)._style=deepcopy(main_sheet[style_source]._style)
        main_sheet.merge_cells(start_row=empty_row+i, start_column=7, end_row=empty_row+i, end_column=7+3)
        main_sheet.cell(row=(empty_row+i), column=11)._style=deepcopy(main_sheet[style_source]._style)
        main_sheet.merge_cells(start_row=empty_row+i, start_column=11, end_row=empty_row+i, end_column=11+25)
        main_sheet.cell(row=(empty_row+i), column=37)._style=deepcopy(main_sheet[style_source]._style)
        main_sheet.merge_cells(start_row=empty_row+i, start_column=37, end_row=empty_row+i, end_column=37+3)
        main_sheet.cell(row=(empty_row+i), column=41)._style=deepcopy(main_sheet[style_source]._style)
        main_sheet.merge_cells(start_row=empty_row+i, start_column=41, end_row=empty_row+i, end_column=41+3)
        main_sheet.cell(row=(empty_row+i), column=45)._style=deepcopy(main_sheet[style_source]._style)
        main_sheet.merge_cells(start_row=empty_row+i, start_column=45, end_row=empty_row+i, end_column=45+3)
