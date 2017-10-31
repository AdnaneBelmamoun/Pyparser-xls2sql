#!/usr/bin/env python
# -*- coding: UTF-8 -*-
#title           :xls2csv.py
#description     :xls file Parser to csv file.
#author          : Adnane Belmamoun
#date            : Copyright (C) 2011-2013 Adnane Belmamoun
#version         :2.1
#usage           :function used upon internal call back 
#notes           :
#python_version  :2.7.19  
#==============================================================================

#importation du module py excelerator pour traitement des document Excel(XLS)

from pyExcelerator import *
import sys
# programation de la ligne de commande xls2csv
# recuperation de ce que l'utilisateur a entree en ligne de commande
me, args = sys.argv[0], sys.argv[1:]
#si les arguments sont non vide on les mets dans un tableau args[]
if args:
    #  pour tout argument dans ce tableau
    for arg in args:
        # On affiche sur sortie standard cet argument
        print >>sys.stderr, 'Extraction des donnees a partir de :', arg
        # Pour chaque nom de feuille et ca valeure rencontree
        # dans le tableau recupére a la sortie de la fonction
        # parse_xls(chemin fichier xls = argument arg).
        for sheet_name, values in parse_xls(arg, 'cp1251'): # parse_xls(arg) -- default encoding
            # on initialise  une matrice
            matrix = [[]]
            # Et on affiche le nom de la feuille 
            print 'Feuille = \"%s\"' % sheet_name.encode('cp866', 'backslashreplace')
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
                v = '"%s"' % v.strip()
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
                print csv_row
#NB: afin de recuperer les resultats dans un fichier CSV,il convient
#    de faire une redirection de la sortie standard vers un fiichier .csv
#    sous windows la redirection se fait grace a " > " comme dans l'exemple qui
#    suit :
#
#         avec python :
#               python xls2csv.py nom_fichier_xls.xls > sortie_csv.csv
#
#        avec ligne de commande executable sous windows:
#                xls2csv.exe chemin_fichier/nom_fichier_xls > sortie_csv.csv

# ici je gere l'exception de fichier excel non trouvee.
else:
    print 'Probleme avec: %s \n (Le Fichier Excel (.XLS) en Entree est introuvable)' % me

