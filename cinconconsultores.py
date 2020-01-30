# -*- coding: utf-8 -*-
"""
Created on Thu Nov 28 14:28:41 2019

@author: user
"""
from selenium import webdriver
import wx
import time
import openpyxl
from webscraping import asignar_nro_proceso, get_the_web
from get_lists import get_cities_entities_web, make_cities_entities_dictionary, make_others_list
import os

col_radicado_ini=2
col_radicado_completo=3
col_fecha_radicacion=4
col_tipo_general_proceso=5
col_tipo_especifico_proceso=6
col_cuantia=7
col_instancia=8
col_responsable=9
col_apoderado=10
col_ciudad=11
col_entidad=12
col_jurisdiccion=13
col_tipo_sujeto_cliente=14
col_tipo_persona_demandante=15
col_razon_social_demandante=16
col_nit_demandate=17
col_tipo_persona_demandado=18
col_razon_social_demandado=19
col_nit_demandado=20
col_tipo_persona_tercero=21
col_razon_social_tercero=22
col_nit_tercero=23



DB = openpyxl.load_workbook('Database-Process.xlsx')
sheet = DB['Hoja1']


class MyFrame(wx.Frame):
    
    
    def OnKeyDown(self, event):
        """quit if user press q or Esc"""
        if event.GetKeyCode() == 27 or event.GetKeyCode() == ord('Q'): #27 is Esc
            self.Close(force=True)
            
        else:
            event.Skip()
 
    def __init__(self):
        
        wx.Frame.__init__(self, None, wx.ID_ANY, "Software Legal", size=(1200, 700))  
        self.Bind(wx.EVT_KEY_UP, self.OnKeyDown)
        
        try:
            image_file = 'CINCO CONSULTORES.jpg'
            bmp1 = wx.Image(image_file, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            
            self.panel = wx.StaticBitmap(self, -1, bmp1, (0, 0))
            
        except IOError:
            print ("Image file %s not found"  )
            raise SystemExit
            
            
        button = wx.Button(self.panel, id=wx.ID_ANY, label="Ingresar Proceso" ,pos=(900, 100), size=(200, 50))
        button.Bind(wx.EVT_BUTTON, self.Ingresarproceso)
        
        button2 = wx.Button(self.panel, id=wx.ID_ANY, label="Consultar Proceso" ,pos=(900, 150), size=(200, 50))
        button2.Bind(wx.EVT_BUTTON, self.BtnConsultaProceso)
        
        button3 = wx.Button(self.panel, id=wx.ID_ANY, label="Actualizar información proceso" ,pos=(900, 200), size=(200, 50))
        button3.Bind(wx.EVT_BUTTON, self.onButton3)
        
        btn_asignar_procesos = wx.Button(self.panel, id=wx.ID_ANY, label="Ident. Nro Proceso" ,pos=(900, 250), size=(200, 50))
        btn_asignar_procesos.Bind(wx.EVT_BUTTON, self.onBtn_asignar_procesos)
        
        ico = wx.Icon('Icono.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)

    
    #-------------Button Functions-----------------#
    def Ingresarproceso(self, event):
        secondWindow = ww_Ingresar_Proceso(parent=self.panel)
        secondWindow.Show()

    def BtnConsultaProceso(self, event): 
        consultawindow=ww_Consultar_Proceso(parent=self.panel)
        consultawindow.Show()

        
    def onButton3(self, event):
        print ("Button pressed!")
        
    def onBtn_asignar_procesos(self, event):
        asignar_nro_proceso()
        
    #-------------Button Functions-----------------#    

        
class ww_Ingresar_Proceso(wx.Frame):
   
    
    def __init__(self,parent):
        
        wx.Frame.__init__(self,parent, -1,'Ingresar Proceso', size=(880,530))
        ciudades_entidades=make_cities_entities_dictionary()
        self.other_lists=make_others_list()
        self.ciudad='MEDELLIN '
        
        try:
            
            image_file = 'CINCO CONSULTORES.jpg'
            bmp1 = wx.Image(
                image_file, 
                wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            
            #self.panel = wx.StaticBitmap(
                #self, -1, bmp1, (0, 0)
            self.panel=wx.Panel(self)
            self.panel.SetBackgroundColour(wx.Colour('white'))
            
            
        except IOError:
            print ("Image file %s not found"  )
            raise SystemExit
        
        
        ico = wx.Icon('Icono.ico', wx.BITMAP_TYPE_ICO)
        title_font= wx.Font(25, wx.FONTFAMILY_DECORATIVE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        
        self.fgs= wx.GridBagSizer(0,0)
        
        self.SetIcon(ico)
        
        self.lbltitle =wx.StaticText(self.panel, label='Nuevo Proceso')
        self.lbltitle.SetFont(title_font)
        self.lbltitle.SetBackgroundColour('white')
        self.fgs.Add(self.lbltitle,pos=(0,4),span=(1,3), flag=wx.ALL | wx.ALIGN_CENTER, border=5)
        
        ################################ CIUDAD ############################################
        self.lblciudad = wx.StaticText(self.panel, label="Ciudad:")
        self.lblciudad.SetBackgroundColour("white")
        self.fgs.Add(self.lblciudad,pos=(2,1),span=(1,1), flag= wx.ALL, border=5)
        self.Ciudad=wx.ComboBox(self.panel, value='Ciudad', choices=ciudades_entidades[1])
        self.Ciudad.Bind(wx.EVT_COMBOBOX, self.get_entidades)
        self.fgs.Add(self.Ciudad,pos=(2,2),span=(1,2), flag= wx.ALL , border=5)
        ################################ CIUDAD ############################################
    
        ################################ ENTIDAD ############################################
        self.lblentidad = wx.StaticText(self.panel, label="Entidad:")
        self.lblentidad.SetBackgroundColour("white")
        self.fgs.Add(self.lblentidad,pos=(3,1),span=(1,1), flag= wx.ALL, border=5)
        self.Entidades=wx.ComboBox(self.panel, choices=[""],size=(520,-1))
        self.fgs.Add(self.Entidades,pos=(3,2),span=(1,5), flag=wx.ALL , border=5)
        ################################ ENTIDAD ############################################

        ################################ JURISDICCION ############################################
        self.lbljurisdiccion = wx.StaticText(self.panel, label="Jurisdicción:")
        self.lbljurisdiccion.SetBackgroundColour("white")
        self.fgs.Add(self.lbljurisdiccion, pos=(4,1),span=(1,1), flag= wx.ALL, border=5)
        self.Jurisdi = wx.TextCtrl(self.panel)
        self.fgs.Add(self.Jurisdi, pos=(4,2),span=(1,1), flag= wx.ALL, border=5)
        ################################ JURISDICCION ############################################
        
        ################################ TIPO_SUJETO ############################################
        self.lbltipo_sujeto = wx.StaticText(self.panel, label="Tipo Sujeto:")
        self.lbltipo_sujeto.SetBackgroundColour("white")
        self.fgs.Add(self.lbltipo_sujeto, pos=(5,1),span=(1,1), flag= wx.ALL, border=5)
        self.Tipsuj = wx.ComboBox(self.panel ,value=self.other_lists[0][0], choices=self.other_lists[0])
        self.Tipsuj.Bind(wx.EVT_COMBOBOX, self.get_tercero)
        self.fgs.Add(self.Tipsuj, pos=(5,2),span=(1,1), flag= wx.ALL, border=5)
        ################################ TIPO SUJETO ############################################
        
        ################################ SECCION DEMANDANTE ############################################
        self.lbldemandante=wx.StaticText(self.panel, label='Demandante')
        self.lbldemandante.SetBackgroundColour("white")
        self.fgs.Add(self.lbldemandante , pos=(7,1),span=(1,2), flag=wx.ALL | wx.ALIGN_CENTER, border=5)
        self.lbltipo_persona_demandante=wx.StaticText(self.panel, label='Tipo Persona')
        self.lbltipo_persona_demandante.SetBackgroundColour("white")
        self.fgs.Add(self.lbltipo_persona_demandante , pos=(8,1),span=(1,1), flag= wx.ALL, border=5)
        self.tipo_persona_demandante = wx.ComboBox(self.panel,value=self.other_lists[1][0], choices=self.other_lists[1])
        self.tipo_persona_demandante.Bind(wx.EVT_COMBOBOX, self.get_labels_demandante)
        self.fgs.Add(self.tipo_persona_demandante , pos=(8,2),span=(1,1), flag= wx.SHAPED|wx.ALL, border=5)
        self.lblrazon_social_demandante=wx.StaticText(self.panel, label='Razon Social')
        self.lblrazon_social_demandante.SetBackgroundColour("white")
        self.fgs.Add(self.lblrazon_social_demandante , pos=(9,1),span=(1,1), flag= wx.ALL, border=5)
        self.razon_social_demandante = wx.TextCtrl(self.panel)
        self.fgs.Add(self.razon_social_demandante , pos=(9,2),span=(1,1), flag= wx.ALL, border=5)
        self.lblid_demandante=wx.StaticText(self.panel, label='NIT')
        self.lblid_demandante.SetBackgroundColour("white")
        self.fgs.Add(self.lblid_demandante , pos=(10,1),span=(1,1), flag= wx.ALL, border=5)
        self.id_demandante = wx.TextCtrl(self.panel)
        self.fgs.Add(self.id_demandante , pos=(10,2),span=(1,1), flag= wx.ALL, border=5)
        ################################ SECCION DEMANDANTE ############################################
        
        ################################ SECCION DEMANDADO ############################################
        self.lbldemandado=wx.StaticText(self.panel, label='Demandado')
        self.lbldemandado.SetBackgroundColour("white")
        self.fgs.Add(self.lbldemandado , pos=(7,4),span=(1,2), flag=wx.ALL | wx.ALIGN_CENTER, border=5)
        self.lbltipo_persona_demandado=wx.StaticText(self.panel, label='Tipo Persona')
        self.lbltipo_persona_demandado.SetBackgroundColour("white")
        self.fgs.Add(self.lbltipo_persona_demandado , pos=(8,4),span=(1,1), flag= wx.ALL, border=5)
        self.tipo_persona_demandado = wx.ComboBox(self.panel, value=self.other_lists[1][0],choices=self.other_lists[1])
        self.tipo_persona_demandado.Bind(wx.EVT_COMBOBOX, self.get_labels_demandado)
        self.fgs.Add(self.tipo_persona_demandado , pos=(8,5),span=(1,1), flag= wx.ALL, border=5)
        self.lblrazon_social_demandado=wx.StaticText(self.panel, label='Razon Social')
        self.lblrazon_social_demandado.SetBackgroundColour("white")
        self.fgs.Add(self.lblrazon_social_demandado , pos=(9,4),span=(1,1), flag= wx.ALL, border=5)
        self.razon_social_demandado = wx.TextCtrl(self.panel)
        self.fgs.Add(self.razon_social_demandado , pos=(9,5),span=(1,1), flag= wx.ALL, border=5)
        self.lblid_demandado=wx.StaticText(self.panel, label='NIT')
        self.lblid_demandado.SetBackgroundColour("white")
        self.fgs.Add(self.lblid_demandado , pos=(10,4),span=(1,1), flag= wx.ALL, border=5)
        self.id_demandado = wx.TextCtrl(self.panel)
        self.fgs.Add(self.id_demandado , pos=(10,5),span=(1,1), flag= wx.ALL, border=5)
        ################################ SECCION DEMANDADO ############################################
        
        ################################ SECCION TERCERO ############################################
        self.lbltercero=wx.StaticText(self.panel, label='Tercero')
        self.lbltercero.SetBackgroundColour("white")
        self.fgs.Add(self.lbltercero , pos=(7,7),span=(1,2), flag=wx.ALL | wx.ALIGN_CENTER, border=5)
        self.lbltipo_persona_tercero=wx.StaticText(self.panel, label='Tipo Persona')
        self.lbltipo_persona_tercero.SetBackgroundColour("white")
        self.fgs.Add(self.lbltipo_persona_tercero , pos=(8,7),span=(1,1), flag= wx.ALL, border=5)
        self.tipo_persona_tercero = wx.ComboBox(self.panel,value=self.other_lists[1][0], choices=self.other_lists[1])
        self.tipo_persona_tercero.Bind(wx.EVT_COMBOBOX, self.get_labels_tercero)
        self.fgs.Add(self.tipo_persona_tercero , pos=(8,8),span=(1,1), flag= wx.ALL, border=5)
        self.lblrazon_social_tercero=wx.StaticText(self.panel, label='Razon Social')
        self.lblrazon_social_tercero.SetBackgroundColour("white")
        self.fgs.Add(self.lblrazon_social_tercero , pos=(9,7),span=(1,1), flag= wx.ALL, border=5)
        self.razon_social_tercero = wx.TextCtrl(self.panel)
        self.fgs.Add(self.razon_social_tercero , pos=(9,8),span=(1,1), flag= wx.ALL, border=5)
        self.lblid_tercero=wx.StaticText(self.panel, label='NIT')
        self.lblid_tercero.SetBackgroundColour("white")
        self.fgs.Add(self.lblid_tercero , pos=(10,7),span=(1,1), flag= wx.ALL, border=5)
        self.id_tercero = wx.TextCtrl(self.panel)
        self.fgs.Add(self.id_tercero , pos=(10,8),span=(1,4), flag= wx.ALL, border=5)
        ################################ SECCION TERCERO ############################################ 
        
        ################################ TIPO PROCESO ############################################
        self.lbltipo_proceso=wx.StaticText(self.panel, label='Tipo Proceso')
        self.lbltipo_proceso.SetBackgroundColour("white")
        self.fgs.Add(self.lbltipo_proceso , pos=(12,1),span=(1,1), flag= wx.ALL, border=5)
        self.tipo_proceso = wx.ComboBox(self.panel,value=self.other_lists[2][0], choices=self.other_lists[2])
        self.fgs.Add(self.tipo_proceso , pos=(12,2),span=(1,1), flag= wx.ALL, border=5)
        ################################ TIPO PROCESO ############################################
        
        ################################ CUANTIA ############################################
        self.lblcuantia_ini = wx.StaticText(self.panel, label="Cuantia:")
        self.lblcuantia_ini.SetBackgroundColour("white")
        self.fgs.Add(self.lblcuantia_ini, pos=(12,4),span=(1,1), flag= wx.ALL, border=5)
        self.cuantia_ini = wx.TextCtrl(self.panel)
        self.fgs.Add(self.cuantia_ini, pos=(12,5),span=(1,1), flag= wx.ALL, border=5)        
        ################################ CUANTIA ############################################
        
        ################################ RADICADO ############################################
        self.lblradicado_ini = wx.StaticText(self.panel, label="Radicado Inicial:")
        self.lblradicado_ini.SetBackgroundColour("white")
        self.fgs.Add(self.lblradicado_ini, pos=(13,1),span=(1,1), flag= wx.ALL, border=5)
        self.radicado_ini = wx.TextCtrl(self.panel)
        self.fgs.Add(self.radicado_ini, pos=(13,2),span=(1,1), flag= wx.ALL, border=5)
        ################################ RADICADO ############################################
        
        ################################ FECHA_RADICADO ############################################
        self.lblfecha_rad = wx.StaticText(self.panel, label="Fecha de Radicacion:")
        self.lblfecha_rad.SetBackgroundColour("white")
        self.fgs.Add(self.lblfecha_rad, pos=(13,4),span=(1,1), flag= wx.ALL, border=5)
        self.Fechara = wx.TextCtrl(self.panel)
        self.fgs.Add(self.Fechara, pos=(13,5),span=(1,1), flag= wx.ALL, border=5)        
        ################################ FECHA_RADICADO ############################################
        
        ################################ RESPONSABLE ############################################
        self.lblresponsable = wx.StaticText(self.panel, label="Responsable:")
        self.lblresponsable.SetBackgroundColour("white")
        self.fgs.Add(self.lblresponsable, pos=(14,1),span=(1,1), flag= wx.ALL, border=5)
        self.Responsable = wx.TextCtrl(self.panel)
        self.fgs.Add(self.Responsable, pos=(14,2),span=(1,1), flag= wx.ALL, border=5)
        ################################ RESPONSABLE ############################################
        
        ################################ APODERADO ############################################
        self.lblapoderado_ini = wx.StaticText(self.panel, label="apoderado:")
        self.lblapoderado_ini.SetBackgroundColour("white")
        self.fgs.Add(self.lblapoderado_ini, pos=(14,4),span=(2,1), flag= wx.ALL, border=5)
        self.apoderado_ini = wx.TextCtrl(self.panel)
        self.fgs.Add(self.apoderado_ini, pos=(14,5),span=(5,1), flag= wx.ALL, border=5)      
        ################################ APODERADO ############################################
        
        ################################ BOTONES ############################################
        btn_crear = wx.Button(self.panel, id=wx.ID_ANY, label="Crear Proceso", size=(200,40))
        self.fgs.Add(btn_crear, pos=(12,7),span=(2,2), flag= wx.ALL, border=0)
        btn_crear.Bind(wx.EVT_BUTTON, self.Crearproceso)
        
        btn_cancelar = wx.Button(self.panel, id=wx.ID_ANY, label="Cancelar",size=(200,40))
        self.fgs.Add(btn_cancelar, pos=(14,7),span=(2,2), flag= wx.ALL, border=0)
        btn_cancelar.Bind(wx.EVT_BUTTON, self.OnCloseWindow)
        ################################ BOTONES ############################################
        
        self.SetBackgroundColour(wx.Colour(100,100,100))
        self.Centre(True)
        self.Show()

        mainSizer= wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.fgs,0, flag=wx.ALIGN_CENTER)
        self.panel.SetSizerAndFit(mainSizer)


    def get_entidades(self,event):

        ciudades_entidades=make_cities_entities_dictionary()
        Ciudad = self.Ciudad.GetValue()
        choices=ciudades_entidades[0][Ciudad]
        self.Entidades.Clear()
        self.Entidades.AppendItems(choices)

    def get_tercero(self,event):
        
        TipSuj=self.Tipsuj.GetValue()
        
        if TipSuj != self.other_lists[0][3]:
            self.lbltercero.Hide()
            self.lbltipo_persona_tercero.Hide()
            self.tipo_persona_tercero.Hide()
            self.lblrazon_social_tercero.Hide()
            self.razon_social_tercero.Hide()
            self.lblid_tercero.Hide()
            self.id_tercero.Hide()
        else:
            self.lbltercero.Show()
            self.lbltipo_persona_tercero.Show()
            self.tipo_persona_tercero.Show()
            self.lblrazon_social_tercero.Show()
            self.razon_social_tercero.Show()
            self.lblid_tercero.Show()
            self.id_tercero.Show()
        
    def get_labels_demandante(self,event):
        tipo_persona=self.tipo_persona_demandante.GetValue()
        
        if tipo_persona==self.other_lists[1][1]:
            self.lblrazon_social_demandante.SetLabel('Nombre') 
            self.lblid_demandante.SetLabel('Cedula')
        else:
            self.lblrazon_social_demandante.SetLabel('Razon Social')
            self.lblid_demandante.SetLabel('NIT')
            
    def get_labels_demandado(self,event):        
        tipo_persona=self.tipo_persona_demandado.GetValue()
        
        if tipo_persona==self.other_lists[1][1]:
            self.lblrazon_social_demandado.SetLabel('Nombre') 
            self.lblid_demandado.SetLabel('Cedula')
        else:
            self.lblrazon_social_demandado.SetLabel('Razon Social')
            self.lblid_demandado.SetLabel('NIT')
            
    def get_labels_tercero(self,event):
        tipo_persona=self.tipo_persona_tercero.GetValue()
        
        if tipo_persona==self.other_lists[1][1]:
            self.lblrazon_social_tercero.SetLabel('Nombre') 
            self.lblid_tercero.SetLabel('Cedula')
        else:
            self.lblrazon_social_tercero.SetLabel('Razon Social')
            self.lblid_tercero.SetLabel('NIT')
            
    def OnCloseWindow(self, event):
        self.Destroy()
    
    def Crearproceso(self, event):
        
        
        global col_radicado_ini
        global col_radicado_completo
        global col_fecha_radicacion
        global col_tipo_general_proceso
        global col_tipo_especifico_proceso
        global col_cuantia
        global col_instancia
        global col_responsable
        global col_apoderado
        global col_ciudad
        global col_entidad
        global col_jurisdiccion
        global col_tipo_sujeto_cliente
        global col_tipo_persona_demandante
        global col_razon_social_demandante
        global col_nit_demandate
        global col_tipo_persona_demandado
        global col_razon_social_demandado
        global col_nit_demandado
        global col_tipo_persona_tercero
        global col_razon_social_tercero
        global col_nit_tercero
        
        

        smmlv=int(self.other_lists[4])
        limite=40*smmlv
        
        Nproce = 1
        
        while (sheet.cell(row = Nproce, column = 1).value != None) :
          Nproce = Nproce + 1
        sheet.cell(row = Nproce  , column = 1).value = Nproce
        

        Ciudad = self.Ciudad.GetValue()
        sheet.cell(row = Nproce, column = col_ciudad).value = Ciudad
        self.Ciudad.Value=""

        Entidad = self.Entidades.GetValue()
        sheet.cell(row = Nproce, column = col_entidad).value = Entidad
        self.Entidades.Value=""        
        
        Jurisdi = self.Jurisdi.GetValue()
        sheet.cell(row = Nproce, column = col_jurisdiccion).value = Jurisdi
        self.Jurisdi.Value=""
        
        Tipo_sujeto = self.Tipsuj.GetValue()
        sheet.cell(row = Nproce, column = col_tipo_sujeto_cliente).value = Tipo_sujeto 
        self.Tipsuj.Value=self.other_lists[0][0]

        Tipo_persona_demandante= self.tipo_persona_demandante.GetValue()
        sheet.cell(row = Nproce, column = col_tipo_persona_demandante).value = Tipo_persona_demandante
        self.tipo_persona_demandante.Value=self.other_lists[1][0]
        
        Razon_social_demandante=self.razon_social_demandante.GetValue()
        sheet.cell(row = Nproce, column = col_razon_social_demandante).value = Razon_social_demandante
        self.razon_social_demandante.Value=""
        
        Id_demandante=self.id_demandante.GetValue()
        sheet.cell(row = Nproce, column = col_nit_demandate).value = Id_demandante
        self.id_demandante.Value=""        
        
        Tipo_persona_demandado= self.tipo_persona_demandado.GetValue()
        sheet.cell(row = Nproce, column = col_tipo_persona_demandado).value = Tipo_persona_demandado
        self.tipo_persona_demandado.Value=self.other_lists[1][0]
        
        Razon_social_demandado=self.razon_social_demandado.GetValue()
        sheet.cell(row = Nproce, column = col_razon_social_demandado).value = Razon_social_demandado
        self.razon_social_demandado.Value=""
        
        Id_demandado=self.id_demandado.GetValue()
        sheet.cell(row = Nproce, column = col_nit_demandado).value = Id_demandado
        self.id_demandado.Value=""
        
        Tipo_persona_tercero= self.tipo_persona_tercero.GetValue()
        sheet.cell(row = Nproce, column = col_tipo_persona_tercero).value = Tipo_persona_tercero
        self.tipo_persona_tercero.Value=self.other_lists[1][0]
        
        Razon_social_tercero=self.razon_social_tercero.GetValue()
        sheet.cell(row = Nproce, column = col_razon_social_tercero).value = Razon_social_tercero
        self.razon_social_tercero.Value=""
        
        Id_tercero=self.id_tercero.GetValue()
        sheet.cell(row = Nproce, column = col_nit_tercero).value = Id_tercero
        self.id_tercero.Value=""
        
        Cuantia = self.cuantia_ini.GetValue()
        sheet.cell(row = Nproce, column = col_cuantia).value = Cuantia
        self.cuantia_ini.Value=""

        Tipo_proceso=self.tipo_proceso.GetValue()
        sheet.cell(row = Nproce, column = col_tipo_especifico_proceso).value = Tipo_proceso
        self.tipo_proceso.Value=self.other_lists[2][0]
        
        index_tipo_proceso= self.other_lists[2].index(Tipo_proceso)
        self.tipo_proceso_general=self.other_lists[3][index_tipo_proceso]
        sheet.cell(row = Nproce, column = col_tipo_general_proceso).value = self.tipo_proceso_general
        
        if self.tipo_proceso_general=="Declarativo":
            if Tipo_proceso==self.other_lists[2][1]:
                self.instancia="Doble Instancia"
            else:
                self.instancia="Unica Instancia"
        
        elif self.tipo_proceso_general=="De Jurisdicción Voluntaria":
            self.instancia="Unica Instancia"
        else:
            if int(Cuantia) < limite:
                self.instancia="Unica Instancia"
            else:
                self.instancia="Doble Instancia"
         
        sheet.cell(row = Nproce, column = col_instancia).value = self.instancia

        Radicado_ini=self.radicado_ini.GetValue()
        sheet.cell(row = Nproce, column = col_radicado_ini).value = Radicado_ini
        self.radicado_ini.Value=""
        
        Responsable = self.Responsable.GetValue()
        sheet.cell(row = Nproce, column = col_responsable).value = Responsable
        self.Responsable.Value=""
        

        Fechara  = self.Fechara.GetValue()
        sheet.cell(row = Nproce, column = col_fecha_radicacion).value = Fechara
        self.Fechara.Value=""
                
        Apoderado = self.apoderado_ini.GetValue()
        sheet.cell(row = Nproce, column = col_apoderado).value =Apoderado
        self.apoderado_ini.Value=""
         
        success_msgbox=wx.MessageDialog(None,'Registro añadido - Este mensaje aun no garantiza que nada haya fallado en el proceso de agregar el registro - /n EnConstruccion','Status',wx.OK)
        success_msgbox.ShowModal()
        
        DB.save('Database-Process.xlsx')

class ww_Consultar_Proceso(wx.Frame):
    

    def __init__(self,parent):
        wx.Frame.__init__(self,parent, -1,'Consultar Proceso', size=(1200,700))   

        try:
            
            image_file = 'CINCO CONSULTORES.jpg'
            bmp1 = wx.Image(
                image_file, 
                wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            
            self.panel = wx.StaticBitmap(
                self, -1, bmp1, (0, 0))
           
       
          
        except IOError:
            print ("Image file %s not found"  )
            raise SystemExit
        
        
        ico = wx.Icon('Icono.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)
        
        self.lblname1 = wx.StaticText(self.panel, label="Ingrese Numero de Proceso", pos=(830, 250))
        self.lblname1.SetBackgroundColour("white")
        self.numero_consulta=wx.TextCtrl(self.panel, size=(300, -1),pos=(750, 270))

        btn_consultar = wx.Button(self.panel, id=wx.ID_ANY, label="Consultar" ,pos=(840, 300), size=(100, 30))
        btn_consultar.Bind(wx.EVT_BUTTON, self.Consultar_Excel)
        
    def Consultar_Excel(self,event):
        
        numero_consulta=self.numero_consulta.GetValue()
        print(os.getcwd())
        workbook_path=os.getcwd()+'/Procesos/'+ numero_consulta + '.xlsx'
        os.startfile(workbook_path)
        
class MyApp(wx.App):
    def OnInit(self):
        self.frame= MyFrame()
        self.frame.Show()
        return True       
 
# Run the program     
app=MyApp()
app.MainLoop()
del app