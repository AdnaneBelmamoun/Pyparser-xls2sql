#!/usr/bin/env python
# encoding: utf-8
#title           :csv2mysql.py
#description     :Parsing csv 2 postsql file.
#author          : Adnane Belmamoun
#date            :Copyright (C) 2011-2013 Adnane Belmamoun
#version         :2.1
#usage           :function used upon internal call back 
#notes           :
#python_version  :2.7.19  
#==============================================================================
import sys
import getopt
import csv
import os
import re

help_message = '''
csv2mysql is python coded tool developped By Belmamoun Adnane 2011
This tool parse CSV file to SQL Query file to send and execute on any SQL DB server '''

class Usage(Exception):
  def __init__(self, msg):
    self.msg = msg

#on crée d'abord la classe csv2mysql héritante de la classe object
class csv2sql(object):
  # ici je défini la fonction d'initialisation de la classe (le constructeur)
  def __init__(self, filename):
    # appel du constructeur de object
    super(csv2sql, self).__init__()
    # je defini ic la variable d'instance courante self.filename
    self.filename = os.path.basename(filename)
    # temp est une variable temporaire contenant le chemin du fichier sans extension
    tmp = os.path.basename(filename).split('.')[0]
    if '/' in tmp:
      # ici j'elimne tout ce qui reste pour ne laisser que la premiére position
      # en commancant de la fin
      tmp = tmp.split('/')[-1:][0]
    # self.tablename devra alors contenir le nom du fchier seulement (sans extension)  
    #self.tablename = tmp
    self.DBname = tmp
    # et donc on peut commencer par ouvrir notre fichier csv
    self.csvfile = open(filename)
    self.csvf = open(filename, 'rb')
    #with open(filename, 'rb') as self.csvf:
    self.csvreader = csv.reader(self.csvf)
    self.parsedrowreader = csv.reader(self.csvf, delimiter=',', quotechar='|')
  # ici je defini une fonction qui test si un x est un entier retournant boolean

  def generateheaders(self,tabn,nbheader):
    sqltabcreate = "CREATE TABLE IF NOT EXISTS `"+tabn+"` ("
    for ind in range(nbheader):
         self.headersdef = 'header'+str(ind+1)+" VARCHAR(255)"
         sqltabcreate = sqltabcreate+self.headersdef
         if ind<nbheader-1:
             sqltabcreate = sqltabcreate +","
         else:
             sqltabcreate = sqltabcreate +")"
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
    for row in self.parsedrowreader:
     #print "-------->   "+str(row)
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

  def isInteger(self, x):
    if x == '' or x == None:
      return True
    try:
      foo = int(x)
      return True
    except:
      return False

# The main Function start here
def main(argv=None):
  if argv is None:
    argv = sys.argv

  try:
    try:
      opts, args = getopt.getopt(argv[2:], 
                                  "hv", 
                                  ["help"])
    except getopt.error, msg:
      raise Usage(msg)
  
    for option, value in opts:
      if option == "-v":
        verbose = True
      if option in ("-h", "--help"):
        raise Usage(help_message)
  
  except Usage, err:
    print >> help_message
    print >> sys.stderr, sys.argv[0].split("/")[-1] + ": " + str(err.msg)
    return 2
  
  try:
    try:
      fname = argv[1]
    except IndexError:
      raise Usage("\n You have to specify the CSV input file..! ")
      return 2
    try:      
      builder = csv2sql(fname)
    except IOError:
      raise Usage(" The file  '"+fname+"' doesn\'t exist or is not a valid CSV file.")
      return 2
    
    oname = fname.split('.')[0]
    if '/' in fname:
      oname = oname.split('/')[-1:][0]
    output = open(oname+'.sql','w+')
    result = builder.generateparsedsql()
#    print result
    print >> output, result
    
  except Usage, err:
    print >> sys.stderr, sys.argv[0].split("/")[-1] + ": " + str(err.msg)
    return 2    
  
if __name__ == "__main__":
  sys.exit(main())
