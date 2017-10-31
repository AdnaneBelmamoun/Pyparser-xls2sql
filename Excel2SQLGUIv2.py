# -*- coding: utf-8 -*-
#title           : Excel2SQLGUI.py
#description     : Excel2SQL GUI
#author          : Adnane Belmamoun
#date            : Copyright (C) 2011-2013 Adnane Belmamoun
#version         : 2.2
#usage           : Excel2SQ GUI Main Window 
#notes           :
#python_version  : 2.7.19  and above
#==============================================================================
import wx
import os
import time
import shutil
import tkFileDialog
import pyExcelerator
import sys
import MySQLdb as mdb
import sys
import io
import codecs
from pyExcelerator import *
import sys
import getopt
import csv
import re


__author__ = "Adnane Belmamoun"
__copyright__ = "Copyright 2011-2013, The Pyparser Project"
__credits__ = ["Adnane Belmamoun"]
__license__ = "GPL"
__version__ = "2.2.1"
__maintainer__ = "Adnane Belmamoun"
__email__ = "adn.belm@gmail.com"
__status__ = "Production"



# Set up some button ID's for the menu

ID_ABOUT=101
ID_OPEN=102
ID_SAVE=103
ID_BUTTON1=300
ID_EXIT=200
defaultoutputDBname=""

class MainWindow(wx.Frame):
    def __init__(self,parent,title):
        wx.Frame.__init__(self,parent,wx.ID_ANY, title)

        # adding the text editor and the status bar 
        self.control = wx.TextCtrl(self, 1, size=(500,400), style=wx.TE_MULTILINE)
        self.CreateStatusBar() # A Statusbar in the bottom of the window

        # adding and setting up the Menu
        filemenu= wx.Menu()
        # use ID_ for future easy reference
        Openmenu = filemenu.Append(ID_OPEN, "&Open"," Select an Excel File")
        filemenu.AppendSeparator()
        Savemenu = filemenu.Append(ID_SAVE, "&Save"," Save log")
        filemenu.AppendSeparator()
        Aboutmenu = filemenu.Append(ID_ABOUT, "&About"," Informations About Excel2SQL Tool")
        filemenu.AppendSeparator()
        Exitmenu = filemenu.Append(ID_EXIT,"&Quit"," Exit Excel2SQL GUI ")

        # Creating the menu bar.
        menuBar = wx.MenuBar()
        menuBar.Append(filemenu,"&Start") # Adding the "filemenu" to the MenuBar
        self.SetMenuBar(menuBar)  # Adding the MenuBar to the Frame content.
        
        #self.defaultoutputDBname = ""
        # Define the code to be run when a menu option is selected
        #wx.EVT_MENU(self, ID_ABOUT, self.OnAbout) # unused deprecated Call from V1.1
        self.Bind(wx.EVT_MENU, self.OnAbout, Aboutmenu)
        #wx.EVT_MENU(self, ID_EXIT, self.OnExit)   # unused deprecated Call from V1.1
        self.Bind(wx.EVT_MENU, self.OnExit, Exitmenu)   
        #wx.EVT_MENU(self, ID_OPEN, self.OnOpen)   # unused deprecated Call from V1.1
        self.Bind(wx.EVT_MENU, self.OnOpen, Openmenu)   
        #wx.EVT_MENU(self, ID_SAVE, self.OnSave)   # unused deprecated Call from V1.1
        self.Bind(wx.EVT_MENU, self.OnSave, Savemenu)   
		
        # Set up buttons horizontally
        self.sizer2 = wx.BoxSizer(wx.HORIZONTAL)

        self.label_nom_bd = wx.StaticText(self,-1,label='SQL DB Name !')
        self.label_nom_bd.SetBackgroundColour(wx.BLUE)
        self.label_nom_bd.SetForegroundColour(wx.WHITE)
        self.sizer2.Add( self.label_nom_bd,1, wx.EXPAND )

        self.entry_nom_db = wx.TextCtrl(self,-1,value="db name")
        self.sizer2.Add(self.entry_nom_db,1,wx.EXPAND)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnPressEnter, self.entry_nom_db)

        self.label_login = wx.StaticText(self,-1,label='MySQL Login :')
        self.label_login.SetBackgroundColour(wx.BLUE)
        self.label_login.SetForegroundColour(wx.WHITE)
        self.sizer2.Add( self.label_login, 1, wx.EXPAND )
        
        self.entry_login = wx.TextCtrl(self,-1,value="root")
        self.sizer2.Add(self.entry_login,1,wx.EXPAND)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnPressEnter, self.entry_login)

        self.label_pass = wx.StaticText(self,-1,label='MySQL Password :')
        self.label_pass.SetBackgroundColour(wx.BLUE)
        self.label_pass.SetForegroundColour(wx.WHITE)
        self.sizer2.Add( self.label_pass, 1, wx.EXPAND )

        self.entry_pass = wx.TextCtrl(self,-1,value="",style=wx.TE_PASSWORD )
        self.sizer2.Add(self.entry_pass,1,wx.EXPAND)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnPressEnter, self.entry_pass)

        self.label_ip = wx.StaticText(self,-1,label=' IP/URL DB :')
        self.label_ip.SetBackgroundColour(wx.BLUE)
        self.label_ip.SetForegroundColour(wx.WHITE)
        self.sizer2.Add( self.label_ip, 1, wx.EXPAND )
        
        self.entry_ip = wx.TextCtrl(self,-1,value="127.0.0.1")
        self.sizer2.Add(self.entry_ip,1,wx.EXPAND)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnPressEnter, self.entry_ip)

        boutonconvertir = wx.Button(self,-1,label="Parse Data")
        self.sizer2.Add(boutonconvertir,1)
        self.Bind(wx.EVT_BUTTON, self.Clickparsing, boutonconvertir)

        boutonarrosage = wx.Button(self,-1,label="Send Data To DB")
        self.sizer2.Add(boutonarrosage,1)
        self.Bind(wx.EVT_BUTTON, self.OnArrosageDBClick, boutonarrosage)


        boutoneffacerlog = wx.Button(self,-1,label="Clear Log")
        self.sizer2.Add(boutoneffacerlog,1)
        self.Bind(wx.EVT_BUTTON, self.effacerlog, boutoneffacerlog)


        self.sizer=wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.control,1,wx.EXPAND)
        self.sizer.Add(self.sizer2,0,wx.EXPAND)

        self.SetSizer(self.sizer)
        self.SetAutoLayout(1)
        self.sizer.Fit(self)

        self.Show(1)

        # definition de l'element menu Apropos affichant un message apropos
        # de l'application
        self.aboutme = wx.MessageDialog( self, " Excel2SQL GUI \n"
                            "Excel2SQL tool is free tool designed and developped \n"
                            " Entirely written in python 2.4 and upgraded with python 2.7.19 \n"
                            "The software GUI is easy to use \n"
                            "Select an XLS file, then Convert it \n"
                            "and finally send it as a table to your MySQL DB.\n"
                            "Enjoy....\n"
                            "Author email: adn.belm@gmail.com\n"
                            "Author name : Adnane Belmamoun\n"							
                            "Copyright (C) 2011-2013 \n"
                            ,"About Excel2SQL V.1.2 GUI", wx.OK)
        self.doiexit = wx.MessageDialog( self, "Thank you for using Excel2SQLGUI By Adnan. Belmamoun\n ",
                        "Please confirm exiting the software.\n", wx.YES_NO)

        self.dirname = ''
        self.dbconn = None# mdb.connect(dbip,dblogin,dbpassw);

    def convertcsv2sql(self,csvinpath):
        # add procedure to ignore the first dots(.) in the path only file xtension dot is left:
        pathparts = csvinpath.split('.')
        tmp =""
        for ii in range(len(pathparts)-1):
            tmp += pathparts[ii]+"."

        #print("temporary csvinpath  :   " + tmp)

        #csvinpath = tmp

        self.startcsv2sql(csvinpath)
        #self.outsqlpath = csvinpath.split('.')[0]+".sql"
        self.outsqlpath = str(tmp+"sql")
        #print("\n temporary self.outsqlpath  :   " + self.outsqlpath +"\n")
        #self.outsqlpath = self.dirname+"/"+(self.filename).rsplit(".xls")[0]+"/"+(self.filename).rsplit(".xls")[0]+".sql"
        #if '/' in csvinpath:
        #     self.outsqlpath = str(self.outsqlpath.split('/')[-1:][0])+".sql"

        output = open(self.outsqlpath,'w+')
        result = self.generateparsedsql()
        #print("resultsql: "+result)
        print >> output, result
        output.close()
        return result
         
    def convertxls2csv(self,inxlspath):
        self.csvoutputpath =self.dirname+"/"+(self.filename).rsplit(".xls")[0]+"/"+(self.filename).rsplit(".xls")[0]+".csv"
        if not os.path.exists(self.csvoutputpath):
            os.makedirs(self.dirname+"/"+(self.filename).rsplit(".xls")[0])
            self.csvwritefilehdl=open(self.csvoutputpath,'w+')
        else:
            self.csvoutputpath = self.dirname+"/"+(self.filename).rsplit(".xls")[0]+"/"+(self.filename).rsplit(".xls")[0]+"_2"+".csv"
            self.csvwritefilehdl=open(self.csvoutputpath,'w+')
            
        # Pour chaque nom de feuille et ca valeure rencontree
        # dans le tableau recupére a la sortie de la fonction
        # parse_xls(chemin fichier xls = argument arg).
        for sheet_name, values in parse_xls(inxlspath, 'cp1251'): # parse_xls(arg) -- default encoding
            # on initialise  une matrice
            matrix = [[]]
            # Et on affiche le nom de la feuille 
            sheetnamerow = 'Sheet Name = \"%s\"' % sheet_name.encode('cp866', 'backslashreplace') 
            #print(sheetnamerow) #'Sheet Name = \"%s\"' % sheet_name.encode('cp866', 'backslashreplace')
            print >> self.csvwritefilehdl, sheetnamerow
            # Et pour chaque valeur indexee par son indice de ligne et celui de colomne
            # d'un tableau ordonnee recuperee a la sortie de la fonction keys()
            # appliquee au tableau values recupere avant par parse_xls()
            for row_idx, col_idx in sorted(values.keys()):
                # On recuper cette valeur
                v = values[(row_idx, col_idx)]
                # si c'est une instance unicode
                if isinstance(v, unicode):
                    # On se doit de la re-encoder
                    v = v.encode('cp866', 'backslashreplace')
                else:
                    # Sinon on la garde comme elle est 
                    v = `v`
                # On annule les espaces de fin et debut des lignes    
                v = ('"%s"' % v.strip())
                #on recuper les les limites des indexes
                last_row, last_col = len(matrix), len(matrix[-1])
                # tant qu'on est dans les bons indexes
                while last_row <= row_idx:
                    # On donne a notre matrice la taille de la feuille
                    matrix.extend([[]])
                    last_row = len(matrix)

                while last_col < col_idx:
                    matrix[-1].extend([''])
                    last_col = len(matrix[-1])
                #on met la valeur dans la matrice a partir de la fin
                matrix[-1].extend([v])
            # pour chaque vecteur ligne dans la matrice        
            for row in matrix:
                # on commence a structurer notre csv
                # on joinant les vecteur lignes par des ","
                csv_row = ', '.join(row)
                # puis on affiche le vecteur ligne csv resultant
                # dans la sortie standard
                #print(csv_row)
                print >> self.csvwritefilehdl, csv_row
        self.csvwritefilehdl.close()
        return 	self.csvoutputpath
         
    def parsingJob(self):  #chemin_complet_fichier, nom_fichier, chemin_absolu):
        #ici je commence mon traitement
        self.xlsinputpath = self.dirname+"/"+self.filename
        incsvpath = self.convertxls2csv(self.xlsinputpath)
        self.control.WriteText(incsvpath)
        self.control.WriteText("\n CSV File Created .... \n")
		
        fichiercsvinread=open(incsvpath,'r')
        self.control.WriteText(fichiercsvinread.read())
        self.control.WriteText("\n")
        fichiercsvinread.close()
				
        res = self.convertcsv2sql(incsvpath)
        fichiersql = open(self.outsqlpath, 'r')
        self.control.WriteText(fichiersql.read())
        self.control.WriteText("\n")
        fichiersql.close()

        # comme dernier etape du traitement viens l'arrosage de la DB SQL par
        # le fichier .SQL obtenu dans l'etape precedente.
        # Remarque: pour cela une connexion a la DB devra etre prévu avant d'entamer
        # cette partie d'arrosage de la DB SQL.
                  
                              

    def Clickparsing(self,event):
        self.control.WriteText("\n The Parsing of the Excel File  : "+self.filename+"  Is In Progress...... \n")
        #chemin_complet = self.dirname+"/"+self.filename
        self.parsingJob()

    
        
    def OnAbout(self,e):
       # if e.GetMenu()==self.OnAbout:
        self.aboutme.ShowModal() # Shows the about menu


    def OnArrosageDBClick(self,event):
        # ici je dois recuperer les valeur entree par l'utilisateur
        # pour pouvir arroser la DB MySQL par le contenu du fichier
        # .SQL resultant de la phase de conversion
        # attention ici j'utilise une fonction d'arrosage
        self.control.WriteText("\n The execution of SQL Queries on the Server is In progress.....\n")
        self.control.WriteText(self.entry_nom_db.GetValue()+"\n")
        self.control.WriteText(self.entry_login.GetValue()+"\n")
        self.control.WriteText(self.entry_pass.GetValue()+"\n")
        self.control.WriteText(self.entry_ip.GetValue()+"\n")
        
        self.arrosage()
        

    def OnPressEnter(self,event):
        self.control.WriteText("")

        
    def OnExit(self,e):
        # A modal with an "are you sure" check - we don't want to exit
        # unless the user confirms the selection in this case ;-)
        igot = self.doiexit.ShowModal() # Shows it
        if igot == wx.ID_YES:
            self.Close(True)  # Closes out this simple application

    def OnOpen(self,e):
        passok = 0
        self.dlg = wx.FileDialog(self, "Select an XLS File", self.dirname, "", "*.*", wx.FD_OPEN)
        while passok== 0:
            #self.dlg = wx.FileDialog(self, "Select an XLS File", self.dirname, "", "*.*", wx.FD_OPEN)
            if self.dlg.ShowModal() == wx.ID_OK:
                self.filename = self.dlg.GetFilename()
                self.dirname = self.dlg.GetDirectory()
                destpath = str(self.dlg.GetDirectory()+"/"+self.dlg.GetFilename()).replace(".xls","/")
                if not os.path.exists(destpath):
                    passok = 1
                else:
                    wx.MessageBox("The Output directory and files with the same Path  :\n"+self.dirname+"/"+self.filename+"\n Already Exist in your output directory \n Please Choose a different Path Or change your input name \n if you you wan to stay in the same directory "
                                  , "Info: Existing files", wx.OK | wx.ICON_INFORMATION)
                    self.dlg.Destroy()
                    passok = 2
            else:
                self.dlg.Destroy()
                passok = 2

                
        if passok== 1:
            # ouvérture du fichier lecture du contenu du fichier puis édition 
            # dans le Textcontrol()
            #filehandle=open(os.path.join(self.dirname, self.filename),'r')
            self.control.WriteText("Opening the XLS File : "+self.dirname+"/"+self.filename+"\n")
            pathname = self.dlg.GetPath()
            try:
                with open(pathname, 'r') as file:
                    #self.control.WriteText(file.read()+"\n")
                    #self.doLoadDataOrWhatever(file)
                    parentdirectory = os.path.dirname(pathname)
                    if not os.path.exists(parentdirectory):
                        os.makedirs(parentdirectory)
                
                    #os.makedirs(self.dirname+"/"+(self.filename).rsplit(".xls")[0]+"_2", 'w+')
                    file.close()
                    # here i remove all the unwanted chars in the DB name
                    self.defaultoutputDBname = self.filename.replace(" ","")#[" ",".",",","\'","\""],["","","","",""])
                    # putting the new DB name in the textctrl field
                    self.entry_nom_db.SetValue(str(self.defaultoutputDBname.split('.')[0]))
            except IOError:
                wx.LogError("Cannot open the XLS file '%s'." % newfile)


	    # maintenant je dois afficher le Fichier 
            self.control.WriteText("The XLS file is : "+self.dirname+"/"+self.filename+"  Is Ready For Parsing to SQl format \n")

            # Report du nom du dérnier fichier lu
            self.SetTitle("Editing the Data from \n : "+self.filename)
            # tout le texte editer par le moteur de converssion peut etre changer
            # c-a-d qu'on peut lui ajouter des observations puis le sauveguarder...
            # tout en appuyant sur enregistrer..
            # Enfin on peux détruire l'instance courante
            #self.dlg.Destroy()

    def OnSave(self,e):
        # enregistrement du texte édité par moteur de converssion
        dlg = wx.FileDialog(self, "Select an Output Path", self.dirname, "", "*.*", \
                wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if dlg.ShowModal() == wx.ID_OK:
            # récupèration du contenu a enregistrer
            itcontains = self.control.GetValue()

            # Ouverture du fichier en ecriture,-->écriture-->enregistrement
            self.savefilename=dlg.GetFilename()
            self.savedirname=dlg.GetDirectory()
            filehandle=open(os.path.join(self.savedirname, self.savefilename),'w+')
            print >> filehandle, itcontains
            #filehandle.Write(itcontains)
            filehandle.close()
        # destrcution de la boite de dialogue
        dlg.Destroy()

    def effacerlog(self,event):
        self.control.SetValue("")
 
    def createDB(self,dbip,dblogin, dbpassw,dbname):
       self.dbconn = mdb.connect(dbip,dblogin,dbpassw);
       dbcreationquery = "CREATE DATABASE IF NOT EXISTS "+str(dbname)+";"
       dbcreationcursor = self.dbconn.cursor()
       dbcreationcursor.execute(dbcreationquery)
       self.dbconn.commit()
       dbcreationcursor.close()
       #print("The Data Base "+dbname+" was successfuly created ..... ")
       self.control.WriteText("\n ----> The Data Base : "+str(dbname)+"   was successfuly created ..... : \n")
       self.control.WriteText("\n Data Base Name : "+str(dbname)+" \n  Data Base IP :  "+str(dbip)+" \n  Data Base Login :  "+str(dblogin) +" \n  Data Base Password :  "+str(dbpassw))
       self.control.WriteText("\n ---->  Proceeding to create the tables on the DB " + str(dbname))
       self.dbconn.close()

    def executequeryfileonDB(self,sqlfilepath):
       sqlfile = open(sqlfilepath, 'r')
       #Lecture des ligne et Compilation et execution des requetes
       insertquerieslist = filter(None, (line.rstrip() for line in sqlfile))
       #insertquerieslist= sqlfile.readlines()
       #print insertquerieslist
       #self.control.WriteText("Data successfully sent to the Data base ----> Please Check your DB")
       try:
           insertquerycursor = self.dbconn.cursor()
           for insertquery in insertquerieslist:
               #print "Query to execute :  "+str(insertquery)
               if str(insertquery):
                  insertquerycursor.execute(str(insertquery))
           #self.control.WriteText("Data successfully sent to the Data base ----> Please Check your DB")
           self.dbconn.commit()
           insertquerycursor.close()
       except mdb.Error as e:
           #print("Query Execution Error %d: %s" % (e.args[0],e.args[1]))
           self.control.WriteText("\n Query Execution Error %d: %s" % (e.args[0],e.args[1]))
           #self.dbconn.close()
       #print(" Data successfully sent to the Data base")
       #print("----> Please Check The existence of your created DataBase on the server ")
       self.control.WriteText("\n Data successfully sent to the Data base \n ----> Please Check The existence of your created DataBase on the server \n")
    def arrosage(self):

        nom_db= self.entry_nom_db.GetValue()#parsedtest
        login = self.entry_login.GetValue()#Adnane
        passw = self.entry_pass.GetValue()# pass
        ip_db = self.entry_ip.GetValue() #127.0.0.1
        
        try:
            self.dbconn = mdb.connect(ip_db, login, passw, nom_db)
        except mdb.Error as e:
            print("Compilation Error %d: %s" % (e.args[0],e.args[1]))
            #self.dbconn.close()			
            if "Unknown database" in str(e.args[1]):#=="Unknown database":

                 #print("The Data Base "+nom_db+" does not Exist on your SQL Server....! \n -----> Proceeding to the creation of the DB with the name :  "+nom_db)
                 self.control.WriteText("The Data Base "+nom_db+" does not Exist on your SQL Server....! \n -----> Proceeding to the creation of the DB with the name :  "+nom_db)
                 self.createDB(ip_db, login, passw,nom_db)

        try:
            self.dbconn = mdb.connect(ip_db, login, passw, nom_db)
        except mdb.Error as e:
            #print("Compilation Error %d: %s" % (e.args[0],e.args[1]))
            self.control.WriteText(" Compilation Error %d: %s" % (e.args[0],e.args[1]))
            
        #print("Executing the set of queries on th SQL server........... \n")
        self.control.WriteText("Executing the set of queries on th SQL server........... \n")
        pathtosqlfile= self.dirname+"/"+((self.filename).split('.'))[0]+"/"+((self.filename).split('.'))[0]+'.sql'		 
        self.executequeryfileonDB(pathtosqlfile)
        #self.control.WriteText("Data successfully sent to the Data base ----> Please Check your DB")
        #print("Now Closing the Connexion with the SQL Data Base \n")
        #print("\t \t <---- Thank you for trying this tool ---->")
        self.control.WriteText("Now Closing the Connexion with the SQL Data Base \n  \t \t <---- Thank you for trying this tool ---->")
        self.dbconn.close()
            #sys.exit(1)
			
#on crée d'abord la classe csv2mysql héritante de la classe object
#class csv2sql(object):
  # ici je défini la fonction d'initialisation de la classe (le constructeur)
    def startcsv2sql(self, filename):
       # appel du constructeur de object
       #super(csv2sql, self).__init__()
       # je defini ic la variable d'instance courante self.filename
       self.filename = os.path.basename(filename)
       # temp est une variable temporaire contenant le chemin du fichier sans extension
       # add procedure to ignore the first dots(.) in the path only file xtension dot is left:
       pathparts = os.path.basename(filename).split('.')
       tmp =""
       for ii in range(len(pathparts)-1):
           tmp += pathparts[ii]+"."
           if '/' in tmp:
            # ici j'elimne tout ce qui reste pour ne laisser que la premiére position
            # en commancant de la fin
            tmp = tmp.split('/')[-1:][0]
       # self.tablename devra alors contenir le nom du fchier seulement (sans extension)  
       #self.tablename = tmp
       #self.DBname = tmp
       # et donc on peut commencer par ouvrir notre fichier csv
       #self.csvfile = open(filename)
       self.csvf = open(filename, 'r')
       #with open(filename, 'rb') as self.csvf:
       self.csvreader = csv.reader(self.csvf)
       #print self.csvreader
       self.prr = csv.reader(self.csvf, delimiter=',', quotechar='|')
       #for row in self.prr:
        # print "-------->   "+str(row)


    def generateheaders(self,tabn,nbheader):
       sqltabcreate = "CREATE TABLE IF NOT EXISTS `"+tabn+"` ("
       for ind in range(nbheader):
             self.headersdef = 'header'+str(ind+1)+" VARCHAR(255)"
             sqltabcreate += self.headersdef
             if ind<nbheader-1:
                 sqltabcreate = sqltabcreate +","
             else:
                 sqltabcreate = sqltabcreate +")"
       #print sqltabcreate
       return sqltabcreate

    def generaterowinserts(self,tabn,freshrow):
      sqlinsertrow = "INSERT INTO `"+tabn+"` VALUES ("+freshrow+")"
      return sqlinsertrow
	
    def generateparsedsql(self):
      self.nbrtab = 0
      self.nbrcollist = list()
      self.indcurrenttab = 0
      self.tabnamelst = list()
      parsedsql = ""
      tabname = ""
      for row in self.prr:
         #print "-------->   "+str(','.join(row))
         #tabname = str(((','.join(row)).split('=')[1])).replace("\"", "").replace(" ", "")
         if str(row).find("=")!= -1:
             tag = str(((','.join(row)).split('=')[0]))
             tabname = str(((','.join(row)).split('=')[1])).replace("\"", "").replace(" ", "")
             self.nbrtab=self.nbrtab+1
             self.tabnamelst.append(tabname)
         else:
             tabinsrow = ','.join(row)
             if self.nbrtab > self.indcurrenttab:
                 self.nbrcollist.append(len(tabinsrow.split(',')))
                 parsedsql += self.generateheaders(str(self.tabnamelst[self.indcurrenttab]),len(tabinsrow.split(',')))+"\n"
             self.indcurrenttab = self.nbrtab
             parsedsql += self.generaterowinserts(str(self.tabnamelst[self.indcurrenttab-1]),tabinsrow)+"\n"
      return parsedsql

   

# Set up a window based app, and create a main window in it
app = wx.App()
view = MainWindow(None, " EXCEL2SQL GUI <<Designed and developed by: Belmamoun Adnane Copyright (c) 2011-2013>>")
# Enter event loop
app.MainLoop()
