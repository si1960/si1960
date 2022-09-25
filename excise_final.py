# -*- coding: utf-8 -*-
#!/usr/bin/env python
#!C:/Python39/

import sys
import importlib
import os
import shutil
sys.path.append('c:\program files (x86)\python38-32\lib\site-packages')
sys.path.append('C:\Program Files (x86)\gs\gs9.56.1\bin')
sys.path.append('C:\\\\Program Files (x86)\\\\gs\\\\gs9.56.1\\\\bin\\gsdll32.dll')


from datetime import datetime
import pandas as pd
import time
import re
import ctypes
import wx

import openpyxl
import numpy as np
import pickle
import PyPDF2
import matplotlib
#import tabula
import camelot.io as camelot
import ghostscript
from PyPDF2 import PdfFileReader
from io import StringIO

def main():
    global os
    def onButton(event):

        print("Button pressed.")

    app = wx.App()

    frame = wx.Frame(None, -1, 'win.py')





    # Create open file dialog
    openFileDialog = wx.FileDialog(frame, "Open", "", "",
                                          "PDF files (*.pdf)|*.pdf",
                                           wx.FD_MULTIPLE | wx.FD_FILE_MUST_EXIST)

    openFileDialog.ShowModal()
    global source
    source=openFileDialog.GetPaths()
    #print(source)
    openFileDialog.Destroy()
    #source = raw_input('Enter path with filename,pdf file: ')

    input_files = source
    print(source)
    source_a = str(source[0])
    folder1 = os.path.dirname(source[0])
    print (folder1)
    
    pdf = PdfFileReader(source_a)
    number_of_pages = pdf.getNumPages()
    print(number_of_pages)
    print(type(number_of_pages))
##    #----------------------------------
##
##    
##    
####    mylist =['  ASSESSED','OUT OF CHARGE COPY']
####    window = wx.ListBox(-1,(20,20),(80,60),mylist,wx.LB_SINGLE)
####    #window = wx.Frame(None, title = "wxPython Frame", size = (300,200)) 
####    #panel = wx.Panel(window) 
####
####    
####    #cunt =wx.ListBox(panel,-1,(20,20),(80,60),mylist,wx.LB_SINGLE)
####    
####     
####    #label = wx.StaticText(panel, label = "Hello World", pos = (100,50)) 
####
####    app = wx.App() 
####
####    window.Show(True) 
####    #cunt.Show(True)
####    app.MainLoop()
##    #------------------------------------
####    class Mywin(wx.Frame): 
####                
####        def __init__(self, parent, title): 
####          super(Mywin, self).__init__(parent, title = title,size = (350,300))
####                    
####          panel = wx.Panel(self) 
####          box = wx.BoxSizer(wx.HORIZONTAL) 
####                    
####          self.text = wx.TextCtrl(panel,style = wx.TE_MULTILINE) 
####             
####          languages = ['  ASSESSED','OUT OF CHARGE COPY']   
####          lst = wx.ListBox(panel, size = (100,-1), choices = languages, style = wx.LB_SINGLE)
####                    
####          box.Add(lst,0,wx.EXPAND) 
####          box.Add(self.text, 1, wx.EXPAND) 
####                    
####          panel.SetSizer(box) 
####          panel.Fit() 
####                    
####          self.Centre() 
####          self.Bind(wx.EVT_LISTBOX, self.onListBox, lst) 
####          self.Show(True)  
####                    
####        def onListBox(self, event): 
####          self.text.AppendText( "Current selection:"
####
####                    +event.GetEventObject().GetStringSelection()+"\n")
####                    
####    ex = wx.App() 
####    Mywin(None,'ListBox Demo') 
####    ex.MainLoop()
###--------------------------------
####from tkinter import *
#### 
##### create a root window.
####top = Tk()
#### 
##### create listbox object
####listbox = Listbox(top, height = 10,
####                  width = 15,
####                  bg = "grey",
####                  activestyle = 'dotbox',
####                  font = "Helvetica",
####                  fg = "yellow")
#### 
##### Define the size of the window.
####top.geometry("300x250") 
#### 
##### Define a label for the list. 
####label = Label(top, text = " FOOD ITEMS")
#### 
##### insert elements by their
##### index and names.
####listbox.insert(1, "Nachos")
####listbox.insert(2, "Sandwich")
####listbox.insert(3, "Burger")
####listbox.insert(4, "Pizza")
####listbox.insert(5, "Burrito")
#### 
##### pack the widgets
####label.pack()
####listbox.pack()
#### 
#### 
##### Display until User
##### exits themselves.
####top.mainloop()
#####-------------------
####
####import tkinter as tk
####
####root = tk.Tk()
####
####listbox = tk.Listbox(root)
####listbox.pack()
####for item in ["1", "2", "3"]:
####    listbox.insert("end", item)
####listbox.select_set(0)
####listbox.focus_set()
####
####def exit_gui(event):
####    global result
####    result = listbox.curselection()
####    root.destroy()
####
####root.bind("<Return>",exit_gui)
####root.mainloop()
####
####print (result)
####    
##    #----------------------------------------
####    wx.choice
####
####    Choice(parent, id=ID_ANY, pos=DefaultPosition, size=DefaultSize,
####       choices=['  ASSESSED','OUT OF CHARGE COPY'],
####           style=0, validator=DefaultValidator, name=ChoiceNameStr)    
####        
##
##
##    
##    
    #-------------------------------  WORKED  ----------- 
    def remove_watermark(wm_text, inputFile, outputFile):
        from PyPDF4 import PdfFileReader, PdfFileWriter
        from PyPDF4.pdf import ContentStream
        from PyPDF4.generic import TextStringObject, NameObject
        from PyPDF4.utils import b_
        
        with open(inputFile, "rb") as f:
            source = PdfFileReader(f, "rb")
            output = PdfFileWriter()

            for page in range(source.getNumPages()):
                page = source.getPage(page)
                content_object = page["/Contents"].getObject()
                content = ContentStream(content_object, source)

                for operands, operator in content.operations:
                    if operator == b_("Tj"):
                        text = operands[0]

                        if isinstance(text, str) and text.startswith(wm_text):
                            operands[0] = TextStringObject('')

                page.__setitem__(NameObject('/Contents'), content)
                output.addPage(page)
                print ('ok1')
            with open(outputFile, "wb") as outputStream:
                output.write(outputStream)

    wm_text = input('Enter WATERMARK ON DOCUMENT HERE  :')
    print (wm_text) 


    wm_text = 'OUT OF CHARGE COPY'            
    #wm_text = '  ASSESSED'
    inputFile = source_a
    outputFile = folder1+"\\output.pdf"
    
     
    remove_watermark(wm_text, inputFile, outputFile)
    print ('ok')
    
    #--------------------------------------------------------
    ''' get data in lattice mode'''
    try:
        os.remove('F:\\data_lattice.csv')
    except OSError:
        pass    


    for i in range(number_of_pages):
        #tables =  tabula.read_pdf(outputfile,lattice=True,pages=str(i))#,table_areas=['60,605,570,497'])#,columns=['100,145,195,242,290,347,394,445,495,525'], edge_tool='4500',strip_text='\n'),#
        tables =  camelot.read_pdf(outputFile,flavor='lattice',pages=str(i),split_text=True)#,table_areas=['60,605,570,497'])#,columns=['100,145,195,242,290,347,394,445,495,525'], edge_tool='4500',strip_text='\n'),#
             
        print (tables)

        ###camelot.plot(tables[0], kind='grid').show()
        ##
        for table in tables:
            df_1 =table.df
            df_1.to_csv('F:\data_lattice.csv', mode='a+',na_rep='Nan',index=False)

            #print(df_1.columns)
        ##            for i in df_1.columns:
        ##                df_1 = df_1.assign( i=df_1[i].str.strip().str.split('\n')).explode(i).reset_index(drop=True)
        ##                print  (df_1.shape)
        ##                df_1.to_csv('F:\data_lattice.csv', mode='a+',na_rep='Nan',index=False)
        ## df_explode = df.assign(names=df.names.str.split(",")).explode('names')
  
    print('ok_lattice')
    
   
    #---------------------------------------------------
    
    ''' get page 1 in stream mode'''
    try:
        os.remove('F:\\data.csv')
    except OSError:
        pass    

    source_i = folder1+"\\output.pdf"
    print (source_i)

    outfile = 'F:\\data.csv'
    list_1=[]
    
    
    list_tables = [
        
        camelot.read_pdf(source_i,flavor='stream',strip_text='\n',pages="1", edge_tool='4500',table_areas=['300,790,540,680']),#w=816 L=1056 total L=8560,
        camelot.read_pdf(source_i,flavor='stream',pages="1", table_areas=['60,660,560,630'],columns=['114,147,190,230,269,320,360,405,449,480,515'],edge_tool='4500',split_text=True,strip_text='\n'),
        camelot.read_pdf(source_i,flavor='stream',strip_text='\n',pages="1", edge_tool='4500',table_areas=['60,620,560,550']),#part a and b
        camelot.read_pdf(source_i,flavor='stream',strip_text='\n',pages="1", edge_tool='4500',table_areas=['60,540,560,525']),#columns['121,173,229,276,323,393,478,520'],
        camelot.read_pdf(source_i,flavor='stream',pages="1", table_areas=['60,525,560,500'],columns=['121,173,229,276,323,393,435,478,520'],edge_tool='4500',strip_text='\n'),
        camelot.read_pdf(source_i,flavor='stream',pages="1", table_areas=['60,490,570,460'],columns=['121,173,229,276,323,385,436,487,524,572'],strip_text='\n',edge_tool='4500'),#provide columns
        camelot.read_pdf(source_i,flavor='stream',strip_text='\n',pages="1", edge_tool='4500',table_areas=['60,412,285,330'],columns=['121,173,229,276']),# wh processingdt not getting processing time
        camelot.read_pdf(source_i,flavor='stream',strip_text='\n',pages="1", edge_tool='4500',table_areas=['60,320,285,262']),# GET LINE 2
        camelot.read_pdf(source_i,flavor='stream',strip_text='\n',pages="1",table_areas=['60,230,285,180'],columns=['85,120,178,229'], edge_tool='4500'),#container details notsplitiing header 1 and2
        #camelot.read_pdf(source_i,flavor='stream',strip_text='\n',guess='false',pages="1", edge_tool='4500',table_areas=['305,420,570,100']),#provide columns 
        camelot.read_pdf(source_i,flavor='stream',strip_text='\n',guess='false',pages="1", edge_tool='4500',table_areas=['345,420,570,350']),#payment details
        camelot.read_pdf(source_i,flavor='stream',strip_text='\n',guess='false',pages="1", edge_tool='4500',table_areas=['355,342,570,250']),#invoice details 
        camelot.read_pdf(source_i,flavor='stream',strip_text='\n',guess='false',pages="1", edge_tool='4500',table_areas=['305,180,570,150']),#occ no 
        ]

    
        
    #outfile = folder1+'\\out.csv'
    outfile  =  'F:\\data.csv'

    for tables in list_tables:
        print (tables)
        #tables.to_csv(outfile, mode='a+',na_rep='Nan',index=False)
        #camelot.plot(tables[0], kind='grid').show()
        
        for table in tables:
            print(table)
            table.to_csv(outfile, mode='a+',na_rep='Nan',index=False)
            #table.to_csv(outfile_S, mode='a+', na_rep='Nan', index=False)
            #tables.export('foo.csv', f='csv', compress=True) # json, excel, html, markdown, sqlite
    df_1 = pd.read_csv(outfile, delimiter=',', names=list(range(30))).dropna(axis='columns', how='all')
    print(df_1)
    # splitting column names which are combined
    search = '1.IGM NO'
    ls=df_1.loc[df_1.isin([search]).any(axis=1)].index.tolist()
        
    df_1.iloc[ls[0]:ls[0]+1,10:11]=df_1.iat[ls[0],9][7:12]
    df_1.iloc[ls[0]:ls[0]+1,9:10]=df_1.iat[ls[0],9][:7]
    df_1.iloc[ls[0]:ls[0]+1,4:5]=df_1.iat[ls[0],3][9:17]
    df_1.iloc[ls[0]:ls[0]+1,3:4]=df_1.iat[ls[0],3][:9]
    df_1.iloc[ls[0]:ls[0]+1,2:3]=df_1.iat[ls[0],1][11:21]
    df_1.iloc[ls[0]:ls[0]+1,1:2]=df_1.iat[ls[0],1][:11]
    print (df_1.iat[ls[0],1])

    search = '1.BOND NO.'
    ls=df_1.loc[df_1.isin([search]).any(axis=1)].index.tolist()

    df_1.iloc[ls[0]:ls[0]+1,3:4]=df_1.iat[ls[0],2][10:21] 
    df_1.iloc[ls[0]:ls[0]+1,4:5]=df_1.iat[ls[0],2][21:31] 
    df_1.iloc[ls[0]:ls[0]+1,2:3]=df_1.iat[ls[0],2][:10] 

    search = '1.SNO 2.LCL/'
    ls=df_1.loc[df_1.isin([search]).any(axis=1)].index.tolist()

    df_1.iloc[ls[0]:ls[0]+1,1:2]=df_1.iat[ls[0],0][6:13] 
    df_1.iloc[ls[0]:ls[0]+1,0:1]=df_1.iat[ls[0],0][:6]

    search = '2.IGM DATE'
    ls=df_1.loc[df_1.isin([search]).any(axis=1)].index.tolist() 
    print (ls)
    
    df_1.to_csv(outfile, mode='w',na_rep='Nan',index=False) 
    
                      
    print ('F:\data.csv saved')
    
    #------------------------------ 

   
    





    ''' get data file with stream flow and get details in part -I--
    then get dadtain another csv file with lattice and proces part -III
    and then part -II.'''
    # with stream part I----

    df_1 = pd.read_csv(r'F:\data.csv', delimiter=',', names=list(range(30))).dropna(axis='columns', how='all')
     
    #print (df_1)


    lsiv =[]  #listOfItemsvaluepair

    def getIndexes(dfObj, value):
        ''' Get index positions of value in dataframe i.e. dfObj.'''
        listOfPos = list()
        # Get bool dataframe with True at positions where the given value exists
        result = dfObj.isin([value])
        # Get list of columns that contains the value
        seriesObj = result.any()
        columnNames = list(seriesObj[seriesObj == True].index)
        # Iterate over list of columns and fetch the rows indexes where value exists
        for col in columnNames:
            rows = list(result[col][result[col] == True].index)
            for row in rows:
                listOfPos.append((row, col))
        # Return a list of tuples indicating the positions of value in the dataframe
        return listOfPos


    listOfElems = ["Port Code","BE No",'BE Date','BE Type','INV','ITEM','CONT',
                  '1.IGM NO','3.INW DATE', '18.TOT.ASS VAL','7.IGST',
                   '6.MAWB NO','7.DATE','8.HAWB NO','9.DATE','15.INT',
                   '16.PNLTY','17.FINE','19.TOT. AMOUNT','2.LCL/','4.SEAL','5.CONTAINER NUMBER','4.CUR'

                    ]#'4.GIGMNO','5.GIGMDT','2.LCL/\nFCL','2.IGM DATE',


    dictOfPos = {elem: getIndexes(df_1, elem) for elem in listOfElems}
    #print('Position of given elements in Dataframe are : ')

    for key, value in dictOfPos.items():
        #print(key, ' : ', value)
        value_new=[value[0][0]+1,value[0][1]]
        #print (value_new)
        print (key, ':', df_1.iloc[value[0][0]+1,value[0][1]])
        # one row below
        lsiv.append((key,df_1.iloc[value[0][0]+1,value[0][1]])) 


    listOfElems = ['EXCHANGE RATE','2.MODE','2.LCL/','4.SEAL','5.CONTAINER NUMBER']


    dictOfPos = {elem: getIndexes(df_1, elem) for elem in listOfElems}
    print('Position of given elements in Dataframe are : ')

    for key, value in dictOfPos.items():
        #print(key, ' : ', value),
        value_new=[value[0][0]+2,value[0][1]]
        #print (value_new)
        print (key, ':', df_1.iloc[value[0][0]+2,value[0][1]])
        # two row below
        lsiv.append((key,df_1.iloc[value[0][0]+2,value[0][1]]))


        
    listOfElems = ["IEC/Br","PKG","G.WT(KGS)",'14.COUNTRY OF CONSIGNMENT',
                   '16.PORT OF SHIPMENT','2.CB NAME','AD CODE']
    dictOfPos = {elem: getIndexes(df_1, elem) for elem in listOfElems}
    #print('Position of given elements in Dataframe are : ')
    for key, value in dictOfPos.items():
        #print(key, ' : ', value)
        value_new=[value[0][0],value[0][1]+1]
        #print (value_new)
        print (key, ':', df_1.iloc[value[0][0],value[0][1]+1])
        # one column right
        lsiv.append((key,df_1.iloc[value[0][0],value[0][1]+1]))

        
    listOfElems = ["GSTIN/TYPE","CB CODE",'13.COUNTRY OF ORIGIN',
                   '15.PORT OF LOADING','16.PORT OF SHIPMENT','OOC NO.','OOC DATE']

    
    dictOfPos = {elem: getIndexes(df_1, elem) for elem in listOfElems}
    #print('Position of given elements in Dataframe are : ')
    for key, value in dictOfPos.items():
        #print(key, ' : ', value)
        value_new=[value[0][0],value[0][1]+2]
        #print (value_new)
        print (key, ':', df_1.iloc[value[0][0],value[0][1]+2])
        # two column right
        lsiv.append((key,df_1.iloc[value[0][0],value[0][1]+2]))


    print (lsiv)

    with open('F:\lsiv', "wb") as fp:   #Pickling
       pickle.dump(lsiv, fp)

    # with open("F:\lsiv", "rb") as fp:   # Unpickling
    #     b = pickle.load(fp)
    df = pd.DataFrame(data=lsiv)

    #convert into excel
    df.to_excel("F:\lsiv_data.xlsx", index=False)






    #-------------------------------------------------------------

    # DATA WITH LATTICE for part III and Part II and part IV 

    df_1 = pd.read_csv(r'F:\data_lattice.csv', delimiter=',', names=list(range(30))).dropna(axis='columns', how='all')

    #print(df_1)
    # ----------------for licence details  from Part -IV-----------------------        
    search = 'F. LICENCE DETAILS'
    i_1 = df_1.loc[df_1.isin([search]).any(axis=1)].index.tolist()
    print(i_1[0])
    search = 'G. CERTIFICATE DETAILS\nH.HSS  DETAILS'
    i_2 = df_1.loc[df_1.isin([search]).any(axis=1)].index

    no_of_licence_details = (i_2[0] - i_1[0])-2
    print('no of licence details: '+str(no_of_licence_details))

    
    dc_licence = {}
    lsdictlicence=[]
    lslicence =[]
    ls3=[]


    print (df_1.iloc[i_1[0]+2,3])
    
    if str(df_1.iloc[i_1[0]+2,3]) != 'nan':
        for k in range(i_1[0],i_2[0]-1):
            for j in range(1,15):
                #print(df_1.iloc[k+1,j].split('\n'))
                #print(df_1.iloc[k+1,j])
                ls3.append(df_1.iloc[k+1,j])
                
            lslicence.append(ls3)
            ls3 =[]
    else:
       for k in range(i_1[0],i_2[0]-1):
            for j in range(0,1):
                #print(df_1.iloc[k+1,j].split('\n'))
                print(df_1.iloc[k+1,j])
                ls3=str(df_1.iloc[k+1,j]).split('\n')
                
            lslicence.append(ls3)
            ls3 =[] 
            
    print (lslicence)    
    
    df = pd.DataFrame(data=lslicence)

    #convert into excel
    df.to_excel("F:\lslicence_data.xlsx", index=False)
    #dc_invoice[key] = df_1.iloc[value[i][0]+3,value[i][1]]
    
    print ('ok3')
 
    #-------------- part II--------------------------

    def getIndexes(dfObj, value):
        ''' Get index positions of value in dataframe i.e. dfObj.'''
        listOfPos = list()
        # Get bool dataframe with True at positions where the given value exists
        dfresult = df_1.isin([value])
        # Get list of ROWS DUE TO axis = 1 that contains the value
        seriesObj = dfresult.any(axis=1)
        rowNumbers = list(seriesObj[seriesObj == True].index)
        # Iterate over list of rows and fetch the column indexes where value exists
        for row in rowNumbers:
            col = list(dfresult.iloc[row,:][dfresult.iloc[row,:] == True].index)
            listOfPos.append((row, col[0]))
        # Return a list of tuples indicating the positions of value in the dataframe
        
        return listOfPos



    # find index of partII ,PartIII and Licence table then pick  details from that referance

    search = 'A.\nINVOICE'
    ls=df_1.loc[df_1.isin([search]).any(axis=1)].index.tolist()
    #print(ls)

    listofinvoice = []
    dc_invoice={}

    for i in range(len(ls)):
        
        listOfElems = ["1.BUYER'S NAME & ADDRESS",'3.SUPPLIER NAME & ADDRESS','2.FREIGHT',
                       '3.INSURANCE','10.SVB CH','11.SVB NO',
                        ]#

        dictOfPos = {elem: getIndexes(df_1, elem) for elem in listOfElems}                              
           
        for key, value in dictOfPos.items():
            #print(key, ' : ', value)
            value_new=[value[i][0]+1,value[i][1]]
            #print (value_new)
            #print (key, ':', df_1.iloc[value[i][0]+3,value[i][1]])
            dc_invoice[key] = df_1.iloc[value[i][0]+1,value[i][1]]
            # one row below

        
        listOfElems = ['1. BCD','3.SWS']

        dictOfPos = {elem: getIndexes(df_1, elem) for elem in listOfElems}
        #print('Position of given elements in Dataframe are : ')

        for key, value in dictOfPos.items():
            #print(key, ' : ', value)
            value_new=[value[i][0]+3,value[i][1]]
            #print (value_new)
            #print (key, ':', df_1.iloc[value[0][0]+3,value[0][1]])
            dc_invoice[key] = df_1.iloc[value[i][0]+3,value[i][1]]
            # three row below

            
        listOfElems = []

        dictOfPos = {elem: getIndexes(df_1, elem) for elem in listOfElems}
        #print('Position of given elements in Dataframe are : ')
        for key, value in dictOfPos.items():
            #print(key, ' : ', value)
            value_new=[value[i][0],value[i][1]+1]
            #print (value_new)
            #print (key, ':', df_1.iloc[value[0][0],value[0][1]+1])
            dc_invoice[key] = df_1.iloc[value[i][0],value[i][1]+1] 
            # one column right

        listOfElems = ["14.Cur\nUSD"] #

        dictOfPos = {elem: getIndexes(df_1, elem) for elem in listOfElems}
        #print('Position of given elements in Dataframe are : ')
        for key, value in dictOfPos.items():
            #print(key, ' : ', value)
            value_new=[value[i][0],value[i][1]+2]
            #print (value_new)
            #print (key, ':', df_1.iloc[value[0][0],value[0][1]+2])
            dc_invoice[key] = df_1.iloc[value[i][0],value[i][1]+2]
            # two column right


        listOfElems = []
        dictOfPos = {elem: getIndexes(df_1, elem) for elem in listOfElems}
        #print('Position of given elements in Dataframe are : ')
        for key, value in dictOfPos.items():
            #print(key, ' : ', value)
            value_new=[value[i][0],value[i][1]+3]
            #print (value_new)
            #print (key, ':', df_1.iloc[value[0][0],value[0][1]+3])
            dc_invoice[key] = df_1.iloc[value[i][0],value[i][1]+3]
            # three column right

        listOfElems = ['A.\nINVOICE' ]
        dictOfPos = {elem: getIndexes(df_1, elem) for elem in listOfElems}
        
        #print('Position of given elements in Dataframe are : ')
        for key, value in dictOfPos.items():
            #print(key, ' : ', value)
            value_new=[value[0][0]+1,value[0][1]+1]
            #print (value_new)
            #print (key, ':', df_1.iloc[value[0][0]+1,value[0][1]+1])
            dc_invoice[key] = df_1.iloc[value[i][0]+1,value[i][1]+1]
            # one row below,one column right

            value_new=[value[0][0]+2,value[0][1]+1]
            #print (value_new)
            #print (key, 'date:', df_1.iloc[value[0][0]+2,value[0][1]+1])
            dc_invoice[key+'date'] = df_1.iloc[value[i][0]+2,value[i][1]+1]
            # two row below,one column right




    
        listOfElems =['1.BUYER\'S NAME & ADDRESS']

        for elem in listOfElems:
            ls1= getIndexes(df_1, elem) 
        print (ls1)
        a = ''
        for  i in range (ls1[0][0]+2,ls1[0][0]+6):
            b = str(df_1.iloc[i,ls1[0][1]])
            a=a+b+';'
            a1=a.replace(',','')
        dc_invoice['BUYER ADDRESS'] = a1
     
        print (a1)
        
        ls2=[]
        listOfElems =['3.SUPPLIER NAME & ADDRESS']
        for elem in listOfElems:
            ls2= getIndexes(df_1, elem) 
        #print (ls2)
        a = ''
        i=0
        for  i in range (ls2[0][0]+2,ls2[0][0]+6):
            b = str(df_1.iloc[i,ls2[0][1]])
            a=a+b+';'
            a1=a.replace(',','')
            
        dc_invoice['3.SUPPLIER ADDRESS'] = a1
     
        #print (a)
       
        #print (dc_invoice)
        listofinvoice.append(dc_invoice)
        dc_invoice = {}
    
    
    #print (listofinvoice)
    print (len(listofinvoice))
    z=len(listofinvoice)
    
    ##    dfinv = pd.DataFrame(data=listofinvoice)
    ##    #convert into excel
    ##    dfinv.to_excel("F:\invoiceall_data.xlsx", index=False)
    
    

    #------------part III  duties -----------------

    ''' approach different as item position is fixed.taking location as fixed'''

    search = 'A. ITEM\nDETAILS'
    ls=df_1.loc[df_1.isin([search]).any(axis=1)].index.tolist()
    #print(ls)
    # ls is list of items index --so no of items
    #print (ls)  # position of invoices in  df_1 part iii

    dict_duties = {}
    list_duties = []

    for i in ls:
        
        dict_duties['invsno']= df_1.iloc[i+1,1]
        dict_duties['itemsno'] = df_1.iloc[i+1,2]
        dict_duties['cth'] = df_1.iloc[i+1,3]
        dict_duties['5.ITEM DESCRIPTION'] = df_1.iloc[i+1,5]
        dict_duties['11.upi'] = df_1.iloc[i+3,1]
        dict_duties['15.S.QTY'] = df_1.iloc[i+3,3]
        dict_duties['16.S.UQC'] = df_1.iloc[i+3,4]
        dict_duties['17.SCH'] = df_1.iloc[i+3,6]
       
        dict_duties['14.ass val'] = df_1.iloc[i+5,9]
        dict_duties['30. total duty'] = df_1.iloc[i+5,11]
        dict_duties['5.igst(lnotesrno)'] = df_1.iloc[i+8,6]
        dict_duties['5.igst(lnotificno)'] = df_1.iloc[i+7,6]

        dict_duties['1.bcd'] = df_1.iloc[i+9,2]
        dict_duties['BCD_AMOUNT'] = df_1.iloc[i+10,2]
        dict_duties['3.sws rate'] = df_1.iloc[i+9,4]
        dict_duties['sws_amount'] = df_1.iloc[i+10,4]
        dict_duties['igst_rate'] = df_1.iloc[i+9,6]
        dict_duties['igst_amount'] = df_1.iloc[i+10,6]
        dict_duties['igst_duty_fg'] = df_1.iloc[i+11,6]
        
        dict_duties['2.ACD rate'] = df_1.iloc[i+9,3]
        dict_duties['2.ACD amount'] = df_1.iloc[i+10,3]
        dict_duties['4.SAD RATE'] = df_1.iloc[i+9,5]
        dict_duties['4.SAD amount'] = df_1.iloc[i+10,5]
        dict_duties['5.caidc'] = df_1.iloc[i+17,6]
        dict_duties['caidc_duty_fg'] = df_1.iloc[i+18,6]  #
        #print (dict_duties)
        list_duties.append(dict_duties)
        dict_duties = {}
        
    #print (list_duties)

        
    # ----  list of licence to add to duties based on iINV N0 and item no other places add zero or nan
    i = 0
    for k in range(0,len(list_duties)): 
        for i in range (1,len(lslicence)):
            if str(list_duties[k]['invsno']) !=  str(lslicence[i][0])  :
                print('ok4')
                for j in (2,3,4,5,6,8,10,11,13):
                    list_duties[k][lslicence[0][j]] = '0'

    for k in range(0,len(list_duties)): 
        for i in range (1,len(lslicence)):
            if str(list_duties[k]['invsno']) ==  str(lslicence[i][0]) and str(list_duties[k]['itemsno']) !=  str(lslicence[i][1])  :
                print('ok4')
                for j in (2,3,4,5,6,8,10,11,13):
                    list_duties[k][lslicence[0][j]] = '0'

                    
    for k in range(0,len(list_duties)): 
        for i in range (1,len(lslicence)):
            if str(list_duties[k]['invsno']) ==  str(lslicence[i][0])  and str(list_duties[k]['itemsno']) ==  lslicence[i][1]  :
                print('ok4')
                for j in (2,3,4,5,6,8,10,11,13):
                    list_duties[k][lslicence[0][j]] = lslicence[i][j]
            
##                for j in (2,3,4,5,6,8,13):
##                    list_duties[k][lslicence[0][j]] = '0'

    

            
                    
                        #print (list_duties[0]['4.LIC NO'])
    #print (list_duties)
    

    #  combine 3 ---lsiv(page 1 data --common to all
    #             listofinvoice--gives invoicewise details,
    #             list_duties -- gives itemwise --datails and taxes

    # added  lsiv items to dictinary listofinvoices
    for lista in listofinvoice:
        for i in range (len(lsiv)):
            lista[lsiv[i][0]] = lsiv[i][1]
    #print (listofinvoice)
   
    ## to combined lsiv+listofinvoice append list_duties give fifnal list

    #-----------final list ------------------
    l_all = []
    ll_duties =[]
    for dictduties in list_duties:
        #print (dictduties['invsno'])
        v=dictduties['invsno']
        print (v)
        for dictb in listofinvoice:
            #print (dictb['A.\nINVOICE'][0])
            #print(dictb['A.\nINVOICE'][0])
            
            if dictb['A.\nINVOICE'][0] == v:
               dictduties.update(dictb)  # adding items of dictionay b to dict a
            print ('ok5')
        ll_duties.append(dictduties)
        l_all.append(ll_duties)
        #print (ll_duties)
        ll_duties = []

    
            
    
    print (l_all)
    
    A_all = [] # take (lenght of invoice  positions  minus 1) th  item of list l_all
    for i in range(len(l_all)):
        A_all.append(l_all[i])
        i=i+(z-1)
        
    print (A_all)
    
    dfab = pd.DataFrame(A_all)
    dfab.to_csv('F:\list.csv', index=False)
    
    #df1 = df1.assign(A=df1['A'].str.strip('\n').str.split('\n')).explode(A).reset_index(drop=True)
    #print (df1)
    
    dfab.to_csv("F:\custom_data.csv",index=False)     
    #print (list_duties)
    #df = pd.DataFrame(data=list_duties)
    #convert into excel
    dfab.to_excel("F:\custom_data.xlsx", index=False)


    # process further  ------------------------------ 
    df =pd.read_csv('F:\custom_data.csv')

    if os.path.exists('F:\cdata.csv'):
      os.remove('F:\cdata.csv')
      
    for i in range(0,len(A_all)):
        a=df.iat[i,0]
        print(a)
        x1 = a.replace("'","")
        
        #print (x1)
        dfa = pd.read_csv(StringIO(x1), header=None)
        #print (dfa)
        dft=dfa.T
        #df1 = pd.DataFrame(data=a)
        dft.columns=['A']
        print (dft)
        dfN= dft["A"].str.split(":", n = 1, expand = True)
        dfN.columns=['A','B'+str(i)]
        dfN.set_index(['A'])
        dfm=dfN.T
        
        dfm.to_csv('F:\cdata.csv', mode='a', index = False, header=None)
       
    dff = pd.read_csv("F:\cdata.csv")
    #df = df[df["A"] != 0]
    dff = dff[dff['{invsno'] != '{invsno' ] # removes mutiple header files with starting as invsno
    dff.to_csv('F:\customdata.csv', mode='w', index = False, header=True) 
       ##dff = pd.read_csv('F:\customdataa.csv')
    #df = df[df["A"] != 0]
    #dff = dff[dff['{invsno'] != '{invsno' ] # removes mutiple header files with starting as invsno
    #dff.to_csv('F:\customdataa.csv', mode='w', index = False, header=True) 
    ##    [{invsno	 itemsno	 cth	 5.ITEM DESCRIPTION	 11.upi	 15.S.QTY	 16.S.UQC	 17.SCH	 14.ass val	 30. total duty	 5.igst(lnotesrno)
    ##     5.igst(lnotificno)	 1.bcd	 BCD_AMOUNT	 3.sws rate	 sws_amount	 igst_rate	 igst_amount	 igst_duty_fg	 2.ACD rate	 2.ACD amount
    ##     4.SAD RATE	 4.SAD amount	 5.caidc	 caidc_duty_fg	 3.LIC SLNO	 4.LIC NO	 5.LIC DATE	 6.CODE	 7.PORT	 8.DEBIT VALUE
    ##     9.QTY	 10.UQC	 11.DEBIT DUTY	 "1.BUYERS NAME & ADDRESS"	 3.SUPPLIER NAME & ADDRESS	 LTD.	 2.FREIGHT	 3.INSURANCE
    ##     10.SVB CH	 11.SVB NO	 1. BCD	 3.SWS	 14.Cur\nUSD	 A.\nINVOICE	 A.\nINVOICEdate	 BUYER ADDRESS	 3.SUPPLIER ADDRESS
    ##     Port Code	 BE No	 BE Date	 BE Type	 INV	 ITEM	 CONT	 1.IGM NO	 3.INW DATE	 18.TOT.ASS VAL	 7.IGST
    ##     6.MAWB NO	 7.DATE	 8.HAWB NO	 9.DATE	 15.INT	 16.PNLTY	 17.FINE	 19.TOT. AMOUNT	 2.LCL/	 4.SEAL	 5.CONTAINER NUMBER
    ##     4.CUR	 EXCHANGE RATE	 2.MODE	 IEC/Br	 PKG	 G.WT(KGS)	 14.COUNTRY OF CONSIGNMENT	 16.PORT OF SHIPMENT	 2.CB NAME	 AD CODE
    ##     GSTIN/TYPE	 CB CODE	 13.COUNTRY OF ORIGIN	 15.PORT OF LOADING	 OOC NO.	 OOC DATE]
    ##    
    ##print (dff.columns)
    dfff=dff[[' 3.SUPPLIER NAME & ADDRESS',' BE No',' BE Date',' 5.ITEM DESCRIPTION',' 6.MAWB NO',' 7.DATE',' A.\\nINVOICE',' EXCHANGE RATE',' cth',' 15.S.QTY',' 16.S.UQC',
    ' 11.upi',' 14.ass val',' itemsno',' 1.bcd',' BCD_AMOUNT',' 3.sws rate',' sws_amount',' igst_rate',' igst_amount',' igst_duty_fg',' 30. total duty',
    ' 2.CB NAME',' CB CODE',' AD CODE',' BE Type',' 2.MODE',' Port Code',' 14.COUNTRY OF CONSIGNMENT',' 16.PORT OF SHIPMENT',' 13.COUNTRY OF ORIGIN',' 15.PORT OF LOADING',
    ' "1.BUYERS NAME & ADDRESS"',' BUYER ADDRESS',' IEC/Br',' 8.HAWB NO',' 9.DATE',' PKG',' G.WT(KGS)',' 1.IGM NO',' 3.INW DATE',' 5.CONTAINER NUMBER',' 4.SEAL',' 2.LCL/',
    ' A.\\nINVOICEdate',' 4.CUR',' 14.Cur\\nUSD',' 2.FREIGHT',' 14.Cur\\nUSD',' 3.INSURANCE',' 3.SUPPLIER ADDRESS',' 4.LIC NO',' 5.LIC DATE',' 7.PORT',' 9.QTY',' 11.DEBIT DUTY',
    ' 2.ACD rate',' 2.ACD amount',' 4.SAD RATE',' 4.SAD amount',' GSTIN/TYPE',' 5.igst(lnotificno)',' 5.igst(lnotesrno)',' 16.PNLTY',' 17.FINE',' 15.INT',' 10.SVB CH',' 11.SVB NO',
    ' 18.TOT.ASS VAL',' OOC NO.',' OOC DATE',' 7.IGST',' 19.TOT. AMOUNT'
    ]]
    dfff.to_csv('F:\customdata.csv', mode='w', index = False, header=True) 


    
    dfff[' A.\\nINVOICE']=dfff[' A.\\nINVOICE'].str[4:]
    #print  (dfff[' A.\\nINVOICE'])
    A = dfff[' EXCHANGE RATE'].str.find('=')
    B = dfff[' EXCHANGE RATE'].str.find('INR')
    #print (A)
    dfff[' EXCHANGE RATE'] = dfff[' EXCHANGE RATE'].str[A+1:B]
    
    dfff.to_csv('F:\customdata.csv', mode='w', index = False, header=True) 

    #df['col'] = df['col'].str.slice(0, 9)     

if __name__ == '__main__':
    try:
        main()
    except SystemExit as e:
        ctypes.windll.user32.MessageBoxW(0,u"Error",u"DataExtractor",48)
        print ('Error!', e)
        print ('Press enter to exit (and fix the problem)')
        raw_input()





