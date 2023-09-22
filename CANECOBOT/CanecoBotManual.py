# -*- coding: utf-8 -*-
"""
Created on Wed Aug  9 08:42:31 2023

@author: axel.streiff
"""
#using pywinauto and PyAutoGUI
# some libraries need to be installed (via pip or conda)

import pywinauto as pwa
from pywinauto.application import Application
import pyautogui as pag 
import time
import pyperclip
import numpy as np 
import os 
import ctypes 
import pandas as pd     
import shutil 

# wait_cpu_usage_lower(threshold=2.5, timeout=None, usage_interval=None)  #??

# all click position depend on a 1920x1080 px screen 

#%% MANUAL PROCESS SUR VERSION READER
 
# THESE ARE THE GLOBAL VARIABLES THAT NEED TO BE PRELOADED

screensize = pag.size()
centerpos = (screensize[0]/2 ,  screensize[1]/2)

specs_labels = ['Repère','Style','Contenu','Nb','Consommation','Longueur',  
             'Protection','Prot Cl','JdB','D.Origine','Alimentation',
             'Lieu','Désignation','Type Câble','Ame','Pôle','Pose',
             'D 1er App','delta_u Max','K Temp.','K Prox.','K Complem.',
             'K Utilisation','K Foisonnement','Cl. Conso.','K Simultanéité',
             'cos_phi']

# target path without parenthesis:     
targetpath = r"C:\Users\axel.streiff\OneDrive - ETE RESEAUX\Documents\Python Scripts\NDC IRVE - DALKIA BORDEROUGEcopy.afr"
folderpath = r"C:\Users\axel.streiff\OneDrive - ETE RESEAUX\Documents\Python Scripts"

# filename parameters
prefix = "NDC IRVE - "
filename = "TESTNDC"

# import variables ->
# Data frames initialisation

#reading mapping excel then storing it into an array
mapFile = pd.read_excel(r"C:\Users\axel.streiff\OneDrive - ETE RESEAUX\Documents\Python Scripts\mapping.xlsx")
mapFile = mapFile[['CLES_DICT','CANECO_SPEC']] #removing extra info outside the frame
mapFile = np.asarray(mapFile)

canImportLabels = 'Repère	Style	Contenu	Nb récepteurs	Consommation	JdB Amont	Alimentation	Désignation	Longueur	Longueur max	L. Chemint.	Mode de pose	dU maxi	Classe conso.	KSimult	K Température	K Proximité	K Complémentaire	K Symétrie fs	Neutre chargé	Circuit Verrouillé	Repère Câble	IB	Etat circuit	Repère aval	Jdb Aval	Ind. Révision	Désignation complément.	Texte1	Texte2	Texte3	Texte4	Texte5	Texte6	Texte7	Texte8	Affectation des phases forcé	Affectation des phases	Amont	Récepteur	Amont	Repère	Longueur	Désignation	Type de câble	Ame	Pôle	Mode de pose	Repère Câble	Nb câbles multi	Câble	Neutre	PE ou PEN	IB	IZ	L. Chemint.	Prix Liaison	Longueur max	Largeur	Hauteur	Poids Liaison'
canImportLabels = canImportLabels.split('	')
canImportFrame = np.asarray(canImportLabels).reshape(1,len(canImportLabels))

# reading ColDefaut excel file for the filling of the import frame        
defaultFile = pd.read_excel(r"C:\Users\axel.streiff\OneDrive - ETE RESEAUX\Documents\Python Scripts\ColDefaut.xlsx")
defaultFile = defaultFile[['COL_NAME','TYPE','DEFAULT_GENERAL','TGBT_SPECIFIC','TD_SPECIFIC','IRVE_SPECIFIC']] #removing extra info outside the frame
defaultFile = np.asarray(defaultFile)

# THIS IS THE INPUT DICTIONNARY EXAMPLE WITH THE RIGHT ARCHITECTURE 

INPUT_DICT = {
    "ID":{
        "TitreNDC": 'TESTNDC',
        "Societe": 'MACDONALDS PORTE ITALIE',
        "AddrPostale": '3 Impasse du Kremlin, 31100, Toulouse, France',
        "Date": '21/09/2023',
        "Indice": 'A',
        "Avancement": 'APD',
        },
    "SOURCE":{
        "Type": "autre",
        "Regime": "IT",
        "Courant": "Mono",
        "PuissanceDispo": 1000,
        "Repere": 'SOURCE',
        },

    "TGBTS":[
        
            {
             "Repere": 'TGBT1',
             "Existant": 'None',
             "Disjoncteur": None,
             "Intersectionneur": None,
             "DIFF":{"Repère":'DIFF1'},
             "TDS":[
                 {
                  "Repere": 'TD1',
                  "Existant": None,
                  "DistCables": None,
                  "ModePose": None,
                  "DoublageCable": None,
                  "IRVES":[
                      {
                       "Repere": 'IRVE1',
                       "Existant": None,
                       "Courant": None,
                       "PuissanceDefaut": None,
                       "ModePose": None,
                       "DistCables": None,
                       "ConsPrevue": None,
                       "CosPhi": None,
                       "CoeffFoisonnement": None,
                       "CoeffProximite": None,
                       "NbPointsCharge": None,
                       "PROTECTION":{"Repère":'PROT1'}
                      },
                      {
                       "Repere": 'IRVE2',
                       "Existant": None,
                       "Courant": None,
                       "PuissanceDefaut": None,
                       "ModePose": None,
                       "DistCables": None,
                       "ConsPrevue": None,
                       "CosPhi": None,
                       "CoeffFoisonnement": None,
                       "CoeffProximite": None,
                       "NbPointsCharge": None,
                      }
                      ]
                 },
                 {
                  "Repere": 'TD2',
                  "Existant": None,
                  "DistCables": None,
                  "ModePose": None,
                  "DoublageCable": None,
                  "IRVES":[
                      {
                       "Repere": 'IRVE3',
                       "Existant": None,
                       "Courant": None,
                       "PuissanceDefaut": None,
                       "ModePose": None,
                       "DistCables": None,
                       "ConsPrevue": None,
                       "CosPhi": None,
                       "CoeffFoisonnement": None,
                       "CoeffProximite": None,
                       "NbPointsCharge": None,
                      },
                      {
                       "Repere": 'IRVE4',
                       "Existant": None,
                       "Courant": None,
                       "PuissanceDefaut": None,
                       "ModePose": None,
                       "DistCables": None,
                       "ConsPrevue": None,
                       "CosPhi": None,
                       "CoeffFoisonnement": None,
                       "CoeffProximite": None,
                       "NbPointsCharge": None,
                      }
                      ]
                 }
                 
                 ]
            }
        
        ]
    }


"""
    The logos of ENSIO and the client are in a Logos folder at the location of the bot (folderpath)
    They're then copied to the destination folder from which they can be imported by Caneco
    The Ensio logo image must be named 'LogoEtude.png'
    The client logo image must be named 'LogoClient.png'
    
"""

destFolder = r'C:\ProgramData\ALPI\Caneco BT\5.12\Labels'   # this is the only folder where Caneco can pick up the logos 
print("copying logos to Caneco folder...")
# removing the logos from the previous iteration on the Caneco LABELS folder
if os.path.isfile(os.path.join(destFolder,'LogoClient.png')):
    os.remove(os.path.join(destFolder,'LogoClient.png'))
if os.path.isfile(os.path.join(destFolder,'LogoEtude.png')):
    os.remove(os.path.join(destFolder,'LogoEtude.png'))

# adding the logos to the Caneco LABELS folder
shutil.copy(os.path.join(folderpath,'Logos','LogoClient.png'),destFolder)
shutil.copy(os.path.join(folderpath,'Logos','LogoEtude.png'),destFolder)

OBJ_INDEX = {'TRANSFO_INT':88,'Tableau':5,'DIV_IRVE TRI':78}   # This dictionnary stores the index of the objects in the Nouveau circuit pane 

#%% Tri fichier cible  

#C:\Users\axel.streiff\OneDrive - ETE RESEAUX\Documents\Python Scripts\NDC IRVE - DALKIA BORDEROUGE(copy).afr
#r"Z:\11- Bassens\06- CFO CFA\04 - IRVE\12 - IZIVIA IG\DALKIA\DALKIA BORDEROUGE\03 - DOCS TECHNIQUES\NDC\NDC IRVE - DALKIA BORDEROUGE.afr"

pathList = ['']
listinc = 0

for inc in range (len(targetpath)):
    if targetpath[inc] == "\\":
        listinc += 1 
        pathList.append('')
    else:    
        pathList[listinc] += targetpath[inc]
        
targetfile = pathList[-1]  #os.path.basename()      


#%% Useful functions  

#deactivating caps lock if activated

def capsLockCheck():
    keystate = ctypes.WinDLL('User32.dll').GetKeyState(0x14)
    
    if keystate == 1:
        pag.press('capslock')
    
#capsLockCheck()

# Putting the software window in fullscreen

def winFull(win):
    # on the main window
    if win.is_maximized() == False:
        win.TitleBar.AgrandirButton.click()  #fullscreen in manual
        #win.maximize()

# pressing any button with pyautogui 

def buttonVisualClick(name,xoffset=0,yoffset=0,speedup = True):
    button = pag.locateCenterOnScreen(os.path.join(folderpath,"ImgButtons",name+'.PNG'),grayscale=speedup)
    
    if button != None:
        pag.click(x = button[0] + xoffset,y = button[1] + yoffset) 
    else:
        print(f"Button {name} not found or already activated")
      
#buttonClick('Tableur')

# file browsing interface

def browsePath(targetpath,winpopup):
    
    fileroot,filepath = os.path.splitdrive(targetpath)
    win['Ce PCSplitButton'].click_input() # win popup openfile
    
    # root
    if fileroot == 'Z:':
        winpopup['AZURE (Z:)TreeItem'].click_input()
        winpopup['Nom du fichierEntry'].click_input()
        pag.write(filepath)
        pag.press('enter')  
        
    if fileroot == 'C:':
        winpopup['Acer (C:)TreeItem'].click_input()
        winpopup['Nom du fichierEntry'].click_input()
        pag.write(filepath)
        pag.press('enter')

# reading a numpy array and writing it properly in a text file (you can close it or not at the end)
        
def array2txt(file,array,replacechar = '---',colwidth = [30,30,300,20],closecond = False):
    rows,cols = np.shape(array)
    
    for linenum in range(rows):
        for cellnum in range(cols):
            try:
                if array[linenum,cellnum].strip() != 'n/a':
                    file.write(array[linenum,cellnum].strip())
                    buffer = colwidth[cellnum] - len(repr(array[linenum,cellnum].strip().replace('"','~')))  # the " chararcter is replaced with ~ so it can be counted
                else: 
                    file.write(replacechar)
                    buffer = colwidth[cellnum] - len(repr(replacechar))
            except:
                file.write(replacechar)
                buffer = colwidth[cellnum] - len(repr(replacechar))
             
            if buffer >= 0:
                for space in range(buffer):
                    if linenum == 0:
                        file.write(' ')
                    else:
                        file.write('-')              
            else:
                raise Exception("Columns given width is not sufficient for array content...")
            
        if linenum == 0:
            file.write('\n')
            sep = "".join(["_"for char in range(sum(colwidth))])
            file.write(sep)
            file.write('\n')
          
        file.write('\n')
        
    file.write(sep)
        
    if closecond == True:
        file.close()

# function for adding a new row to the import frame from input dictionnary 

def fillrow(parent,mapFile,canImportFrame,group,prevobj = None):
    newrow = np.empty((1,len(canImportFrame[0,:])),dtype= object) 
    newrow.fill(np.NaN)
    extrarow = np.empty((1,len(canImportFrame[0,:])),dtype= object) 
    extrarow.fill(np.NaN)   
    add = False
    
    for stat in parent.keys():
        try: 
            #print(stat)
            rownum = list(mapFile[:,0]).index(stat+'_'+group)
            #print(rownum)
            try:
                can_specs = mapFile[rownum,1].split(';')
                #print(can_specs, parent['Repere'])
                
                for spec in can_specs:
                    for colnum in range(len(canImportFrame[0,:])):
                        if canImportFrame[0,colnum] == spec:
                            newrow[0,colnum] = parent[stat]
                        elif canImportFrame[0,colnum] == 'Amont':
                            newrow[0,colnum] = prevobj
            except:
                pass
            
        # special cases when a new type of device is put in between  (the settings of the device must be directly the same as the ones present in the import frame)
        except:
            if stat not in ['TGBTS','TDS','IRVES']:
                print('I found a special case : ',parent[stat])
                add = True
                
                for spec in parent[stat]:
                    print(spec)
                    for colnum in range(len(canImportFrame[0,:])):
                        if canImportFrame[0,colnum] == spec:
                            extrarow[0,colnum] = parent[stat][spec]
                        elif canImportFrame[0,colnum] == 'Amont':
                            extrarow[0,colnum] = prevobj
                            
                #print("extrarow = ",extrarow)
                
    return newrow,extrarow,add

#%% Application startup

caneco = Application(backend='uia')
caneco.start(r"C:\Program Files (x86)\ALPI\Caneco BT\5.12\Caneco5.exe")


# bypassing key warning
win = caneco.top_window()

if win.AvertissementDialog.exists() == True:
    
    win.click_input()
    win.OuiButton.click()
    raise Warning("MISSING LICENSE KEY: Launched reader version...") 
    
else: 
    print("Succesfully launched Caneco BT classic")

time.sleep(10)
win = caneco.top_window()
winFull(win) # checking for fullscreen 
    
#%% Reconnecting to Caneco when already opened

caneco = Application(backend='uia')
caneco.connect(path=r"C:\Program Files (x86)\ALPI\Caneco BT\5.12\Caneco5.exe")
win = caneco.top_window()
print("--RECONNECTED TO CANECO--")

winFull(win) # checking for fullscreen 

#%% Filling the caneco import frame from a dictionnary

"""
this is from import.py and works on the architecture of the INPUT_DICT preloaded above 
"""

# Reading the input dictionnary and filling the import frame

# prise en compte de l'objet AMONT !!!

parent = INPUT_DICT
for tgbt in range(len(parent['TGBTS'])):
    parent = INPUT_DICT['TGBTS'][tgbt]
    print('TGBT: ',parent.keys())
    print('')
    newrow,extrarow,add = fillrow(parent,mapFile,canImportFrame,'TGBT',prevobj = 'SOURCE')
    
    # uploading the new line down in the last row of the import frame
    canImportFrame = np.concatenate((canImportFrame,newrow))
    
    # the line for the exceptional objects are only added if they're not empty 
    if add == True:
        canImportFrame = np.concatenate((canImportFrame,extrarow))
    
    for td in range(len(parent['TDS'])):
        prevobj = parent['Repere']
        parent = parent['TDS'][td]
        print('TD: ',parent.keys())
        print('')
        
        newrow,extrarow,add = fillrow(parent,mapFile,canImportFrame,'TD',prevobj = prevobj)
        canImportFrame = np.concatenate((canImportFrame,newrow))
        if add == True:
            canImportFrame = np.concatenate((canImportFrame,extrarow))
        
        for irve in range(len(parent['IRVES'])):
            prevobj = parent['Repere']
            print('IRVE: ',parent['IRVES'][irve].keys())
            print('')
            
            newrow,extrarow,add = fillrow(parent['IRVES'][irve],mapFile,canImportFrame,'IRVE',prevobj = prevobj)
            canImportFrame = np.concatenate((canImportFrame,newrow))
            if add == True:
                canImportFrame = np.concatenate((canImportFrame,extrarow))
        
        parent = INPUT_DICT['TGBTS'][tgbt]
     
    parent = INPUT_DICT


#  Filling the remaining columns with default values (this section is optional and might alter the import)
objlist = ['TGBT','TD','IRVE']

for defcol in range(len(defaultFile[:,0])):
    for colnum in range(len(canImportFrame[0,:])):
        if defaultFile[defcol,0] == canImportFrame[0,colnum]:
            
            if defaultFile[defcol,1] == 'general':
                for cellnum in range(1,len(canImportFrame[:,colnum])):
                    if pd.isnull(canImportFrame[cellnum,colnum]) == True :
                        canImportFrame[cellnum,colnum] = defaultFile[defcol,2] 
                        
            elif defaultFile[defcol,1] == 'specific':
                dta = list(defaultFile[defcol,3:])
                for cellnum in range(1,len(canImportFrame[:,colnum])):
                    for objectnum in range(len(objlist)):
                        if canImportFrame[cellnum,0].find(objlist[objectnum]) != -1:
                            canImportFrame[cellnum,colnum] = dta[objectnum]
                            
# Transforming to csv for export                
                
# deleting previous iteration if the csv file exists
if os.path.exists(r'C:\Users\axel.streiff\OneDrive - ETE RESEAUX\Documents\Python Scripts\CanecoImport.csv'):
  os.remove(r'C:\Users\axel.streiff\OneDrive - ETE RESEAUX\Documents\Python Scripts\CanecoImport.csv')

# saving new csv file
canImportStripped = canImportFrame[1:,:]
canImportFrame = pd.DataFrame(canImportStripped,columns= canImportLabels)

canImportFrame.to_csv(r'C:\Users\axel.streiff\OneDrive - ETE RESEAUX\Documents\Python Scripts\CanecoImport.csv',sep=';', index=False, encoding = 'utf-8-sig')

#%% Import/export FROM CSV TO CANECO

custom = True
preset = "TestVideo"
presetnum = 7

SETTINGS_labels = ['Nom','Type','Opération','Matériels génériques (Pseudo Generic Object)','Informations générales',
                   'Sources','Tableaux et transformateurs','ASI','Canalisations préfabriquées','Circuits','Paramètres de calcul','Pertes Joule']

SETTINGS = {'Nom':'ConfigName','Type':'Transfert CSV','Matériels génériques (Pseudo Generic Object)':'test',
            'Opération':'Import','Circuits':'testCircuits','Sources':'testSources','Tableaux et transformateurs':'testTableaux'}

capsLockCheck()

# opening the import popup
buttonVisualClick('AccueilPanel')
buttonVisualClick('ImportTexte')
time.sleep(1)

#%%
# Selecting the desired preset 
win['Pane2'].Pane.click_input()

findpreset = True
test = 1
while findpreset:
    win['Pane26'].click_input()
    pag.hotkey('ctrl','a')
    pag.hotkey('ctrl','c')
    
    for tries in range(2):
        presetstr = str(pyperclip.paste())

    if presetstr == preset:
        findpreset = True
    else:
        win['Pane2'].Pane.click_input()
        pag.press('up',presses= test)
        test += 1
    
    if test >= presetnum:
        raise Warning('COULD NOT FIND THE DESIGNATED IMPORT PRESET')


#%%

pag.press('up',presses=4)

# Customise the preset
if custom == True:
    for key in SETTINGS.keys():
        if key in SETTINGS_labels:
            
            indkey = SETTINGS_labels.index(key)
           
            # exceptions in label names
            if key == 'Nom':
                pass
            
            elif key == 'Type':
                print(SETTINGS[key])
                win['Pane25'].click_input()
                pag.write(SETTINGS[key])
                pag.press('enter')
                #time.sleep(2)
                
            elif key == 'Opération':
                print(SETTINGS[key])
                win[SETTINGS[key]+'RadioButton'].click()
                
            else:
                print(SETTINGS[key])
                win[key+'Pane'].click_input()
                #pag.press('space')
                
                if key != 'Matériels génériques (Pseudo Generic Object)':
                    win['Edit' + str(indkey-4)].click_input()
                    pag.hotkey('ctrl','a')
                    pag.press('delete')
                    pag.write(SETTINGS[key])
                    pag.press('enter')
                    #time.sleep(2)
        else:
            raise Warning('INVALID SETTINGS LABEL FOUND...')

#win['ExécuterButton'].click()


#%% Browsing for a file  !!!process un peu long!!!!

try:
    capsLockCheck()
    
    buttonVisualClick('OpenFile')

    browsePath(targetpath,winpopup = win)     # add top_window in case bug 
        
    time.sleep(10)
    # conversion de format popup 
    if win["Conversion de format d'affaireDialog"].exists() == True:           
        win["L'affaire est en cours d'étudeRadioButton"].click_input()
        win["ConvertirButton"].click_input()
        
    buttonVisualClick('AgrandirWinAffaire')
        
except: 
    raise Exception('FILE COULD NOT OPEN...')
    
#%% Display control (Enter table or unifilaire  + choosing between tree items: TGBT/BAT B/TD IRVE) 

# THIS METHOD IS SENSIBLE TO THE SOFTWARE VERSION AND BE BROKEN IF THE BUTTONS LOCATION AND STYLE ARE CHANGED

def displayMod(display,treeitem = None):

    buttonVisualClick("BasseTensionPanel") # making sure the right panel is on display
    time.sleep(0.5)
    buttonVisualClick(display)
    time.sleep(0.5)
    buttonVisualClick("Distribution",xoffset=60)
    
    if treeitem != None:
        pag.write(treeitem)
        pag.press('enter')
    else:
        pass
    
#%% Adding objects to the model 

def newTreeObject(obj,parent_device,number,tag_circuit=None,tag_distri=None):  # use spaces instead of underscores in the tag_circuit and tag_distri parameters

    capsLockCheck()
    
    pressnum = OBJ_INDEX[obj]
    
    # accessing the panel 
    displayMod(display='Unifilaire général2')           
    win.type_keys(parent_device)
    pag.press('enter')
    
    time.sleep(0.5)
    buttonVisualClick('NouveauCircuit')
    
    win['Nouveau circuitDialog'].click_input()
    pag.press(['p','c','up'])   #reset cursor
    
    pag.press('down',presses=pressnum)
    
    time.sleep(4)
    
    pag.hotkey('ctrl','c')
    selection = str(pyperclip.paste())
    
    
    if selection.find(obj) == -1:
        win['AnnulerButton'].click()
        raise Exception('INVALID OBJECT INDEX...')
    
    win['Pane8'].click_input()
    mousepos=pag.position()
    pag.click(mousepos[0],mousepos[1])
    
    if tag_circuit != None:
        pag.hotkey('ctrl','a')
        pag.press('delete')
        pag.write(tag_circuit)
        
    pag.press('tab')
    
    if tag_distri != None:
        pag.hotkey('ctrl','a')
        pag.press('delete')
        pag.write(tag_distri)     
        
    pag.press('tab')
    
    pag.press('up',presses=number-1)  # setting the number of devices
    
    pag.press('enter')

# 2 examples in a row  
newTreeObject(obj ='Tableau',parent_device = 'T1',number = 1,tag_circuit = 'TD1',tag_distri = 'TD1')  
newTreeObject(obj ='DIV_IRVE TRI',parent_device = 'TD1',number = 6,tag_circuit = None,tag_distri = None)  
#%% Abonnement Setup

# le type d'abonnement doit arriver sous une des formes presentes dans tariff_optn
tariff_optn = ['jaune','bleu','vert','autre']
tariff = (INPUT_DICT['SOURCE']['Type']).lower()                     # inserer valeurs dict 
power = str(INPUT_DICT['SOURCE']['PuissanceDispo'])
config = INPUT_DICT['SOURCE']['Regime']

if tariff in tariff_optn:
    
    if tariff == 'vert' or tariff == 'autre':   # opening the new blank project pane
        win.click_input()
        buttonVisualClick('Nouveau')
        
        time.sleep(5)
        try:
            buttonVisualClick('Avertissement')
            time.sleep(0.5)
            buttonVisualClick('OKButton')
            
        except:
            pass
        
        time.sleep(0.5)
        win.OKButton.click()
        
        time.sleep(3)
        pag.write(power)
        time.sleep(0.5)
        pag.press('tab')
        
        if tariff == 'autre':
            win['TNPane'].type_keys('IT sans N')
            pag.press('tab')
        
        if tariff == 'vert':
            win['TNPane'].type_keys('TN')
            pag.press('tab')
            
        win['CalculerButton'].click()
        
        while pag.locateCenterOnScreen(os.path.join(folderpath,'ImgButtons','OKUniversal.PNG'),grayscale = True) != None:  #looping through popups
            win.OKButton.click()
            time.sleep(1)
        
        # adding the necessary transformator for special cases
        if tariff == 'autre':
            
            newTreeObject(obj ='TRANSFO_INT',parent_device = 'TGBT',number = 1,tag_circuit = "TRANSFO IT-TN",tag_distri = 'T1')
            time.sleep(0.5)
            
            # changing form IT to TN 
            buttonVisualClick("Distribution",xoffset=60)
            win.type_keys('T1')
            pag.press('enter')
            
            pag.click(centerpos[0],centerpos[1])
            pag.press('up')
            pag.press('enter')
            
            time.sleep(2)
            win['Transformateur avalTreeItem'].click_input()
            win['IT sans NPane'].type_keys('TN')
            time.sleep(0.5)
            pag.press('tab')
            win.OKButton.click()

    else:                                       # opening the new project from model pane 
        buttonVisualClick('AccueilPanel')
        time.sleep(0.5)
        buttonVisualClick('NouvelleAffaire')
        time.sleep(0.5)
        
        if tariff == 'jaune':
            win['Source P. surveilléeTreeItem'].click_input()
            win['Tarif Jaune Dimensionnent NF C15100ListItem'].click_input()
            mousepos = pag.position()
            pag.click(mousepos[0],mousepos[1],clicks=2)
             
        else:                   
            win['Source P. limitéeTreeItem'].click_input()                  #bleu
            win['Puissance limitée (ex-Tarif Bleu)ListItem'].click_input()
            mousepos = pag.position()
            pag.click(mousepos[0],mousepos[1],clicks=2)
            
        time.sleep(4)
        try:
            buttonVisualClick('Avertissement')
            buttonVisualClick('OKButton')
            
        except:
            pass
        
        # checking for special cases 
        
        if tariff == 'bleu' and config == 'mono':
            print('specialcase')
            displayMod(display='Unifilaire général2')   #FINIR CLICK 230 PN
           
            
else:
    raise Exception('INVALID TARIFF NAME...')

#%% Model Setup (Filling up ID)

capsLockCheck()
INPUTS = INPUT_DICT
NDCID = INPUT_DICT['ID']

# opening the affaire info popup
buttonVisualClick('AccueilPanel')
time.sleep(0.5)
buttonVisualClick('InformationsAffaire')
time.sleep(0.5)


ETUDE_DEF = {'Societe':'ENSIO',
             'AddrPostale':'ZI du Chapitre, 7 Chemin des Silos, 31100, Toulouse, France'}    # adapted for list format 

#REMOVE LATER
#NDCID = {'TitreNDC': ['TESTNDC'], 'Societe': ['MACDONALDS PORTE ITALIE'], 'AddrPostale': ['3 Impasse du Kremlin, 31100, Toulouse, France'], 'Date': ['21/09/2023'], 'Indice': ['A'], 'Avancement': ['APD']}

# the lines and pixels offsets corresponding to each data are vulnerable to changes in the interface, you need to update the numbers here

PANES_TABS = {"Généralités":{'Indice':[4],'Date':[5],'Avancement':[6]},
              "Client":{'Societe':[2],'AddrPostale':[4,5,6,7,12]},
              "Etude":{'Societe':[2],'AddrPostale':[4,5,6,7,12]}}

PANES_POS_OFFSET = {"Généralités":200,
                    "Client":140,
                    "Etude":110}

# filling up the name
win["C1510020Pane"].click_input()
pag.hotkey('shift','tab')
pag.hotkey('ctrl','a')
pag.press('delete')
pag.write(prefix + NDCID['TitreNDC'])  # here the name is a list element !!

for pane in PANES_POS_OFFSET.keys():
    
    win['Pane4'].click_input()
    mousepos = pag.position()
    print(pane)
    pag.click(mousepos[0],mousepos[1]-PANES_POS_OFFSET[pane])
    #time.sleep(3)
    
    if pane == 'Etude':
        filldict = ETUDE_DEF
    else:
        filldict = NDCID
        
    for entry in PANES_TABS[pane].keys():
        print(entry)
        
        if entry == 'AddrPostale':
            filldict['AddrPostale'] = filldict['AddrPostale'].split(',')   # working fr a list format from dict add or remove [0]
            
            for ele in range(len(filldict['AddrPostale'])):
                filldict['AddrPostale'][ele] = filldict['AddrPostale'][ele].strip(' ')
            #ite1 = False
            
            if len(filldict['AddrPostale'])<5:
                filldict['AddrPostale'].insert(1,' ')
                
            print(filldict['AddrPostale'])
        
        for lineInd in range(len(PANES_TABS[pane][entry])):
            win["C1510020Pane"].click_input()
            pag.press('tab',presses=PANES_TABS[pane][entry][lineInd])
            pag.write(filldict[entry])
            pag.press('tab')
            #time.sleep(3)


    if pane in ['Etude','Client']:
        
        # clicking in the logo section 
        win["C1510020Pane"].click_input()
        mousepos = pag.position()
        pag.click(mousepos[0]+120,mousepos[1]+90,clicks=2)  # position of the logo section might change and offsets might need to be updated
        
        # selecting the image 
        win['Nom du fichier\xa0:ComboBox'].click_input()
        pag.hotkey('ctrl','a')
        pag.press('delete')
        pag.write('Logo'+pane+'.png')
        pag.press('enter')
        
win['OKButton'].click()


#%% Reading table info + storing in array   

try: 
    
    #bring back cursor 
    pag.moveTo(centerpos[0],centerpos[1],0) 
    pag.click()
    
    pag.mouseDown(button='middle')
    time.sleep(1)
    pag.mouseUp(button='middle',x=centerpos[0]-400,y=centerpos[1]-400)
    
    # copy & paste full table in a string
    buttonVisualClick('SelectFullTableur')
    
    time.sleep(1)
    pag.hotkey('ctrl','c')
    fulltablestr = str(pyperclip.paste())
    
    # parsing
    tablesplit = fulltablestr.split('\n')
    tablesplit.pop(-1)  # removing excess elements
    
    
    fulltableArr = []
    for line in tablesplit:
        
        # we make sure to remove the extra SJB lines from the copy & paste (useless)
        if line.find('Associé Jdb') == -1:
            linesplit = line.split('\t')
            fulltableArr.append(linesplit)
      
    # transformming to numpy array
    fulltableArr = np.array(fulltableArr)
    
    # clean up array data
    for j in range(np.shape(fulltableArr)[1]):
        for i in range(np.shape(fulltableArr)[0]):
            try:
                if fulltableArr[i,j][0] == '"':
                    fulltableArr[i,j] = fulltableArr[i,j].strip('"')
                    print(fulltableArr[i,j])
            except:
                pass
    
    print(fulltableArr)
    pag.press('esc')
    print("Reading successful")
   
except:
    pag.press('esc')
    raise Exception("TABLE READING FAILED...")

#%% Insert values in table / Changing one or multiple cases in one call

try: 
    # target parameters 
    TARGET_SPECS = ['Longueur','Consommation','Ame']       
    
    TARGET_NAMES = [['BORNE 3', 'BORNE 4', 'BORNE 8', 'BORNE 2'],
                   ['DIFF - BORNE 2', 'BORNE 5', 'DIFF - BORNE 5', 'BORNE 1'],
                   ['DIFF - BORNE 2', 'BORNE 5', 'DIFF - BORNE 5', 'BORNE 1', 'DIFF - BORNE 8', 'DIFF - BORNE 3']]
    
    TARGET_LINES = [[6, 8, 16, 4],
                    [3, 10, 9, 2],
                    [3, 10, 9, 2, 15, 5]]
    
    DATA = [['101','102','103','104'],
            ['conso1', 'conso2','conso3', 'conso4'],
            ['Al', 'Al', 'Al', 'Al', 'Al', 'Al']]                      
    
    capsLockCheck()
    
    #bring back cursor 
    pag.moveTo(centerpos[0],centerpos[1],0) 
    pag.click()
    
    pag.mouseDown(button='middle')
    time.sleep(1)
    pag.mouseUp(button='middle',x=centerpos[0]-400,y=centerpos[1]-400)
    
    # entry point table
    buttonVisualClick('SelectFullTableur')
    
    # cursor starts moving according to the offsets
    mod = 'line' #this mode allows for changing line numbers or line name parameters
    cursor_pos = (0,0)
    namelist = list(fulltableArr[:,0])
    
    for col in range(len(TARGET_SPECS)):
        for line in range(len(TARGET_LINES[col])):
            
            if mod == 'line':
                offset = specs_labels.index(TARGET_SPECS[col]) - cursor_pos[0] , TARGET_LINES[col][line] -1 - cursor_pos[1]
            if mod == 'name': 
                offset = specs_labels.index(TARGET_SPECS[col]) - cursor_pos[0] , namelist.index(TARGET_NAMES[col][line]) - cursor_pos[1]
           
            # default moving directions
            moveX = 'right'
            moveY = 'down'
            
            # if not sorted the directions are inverted
            if offset[0] <= 0:
                moveX = 'left'
            if offset[1] <= 0:
                moveY = 'up'
                
            # moving keyboard commands
            pag.press(moveX, presses= abs(offset[0]))
            pag.press(moveY, presses= abs(offset[1]))
            
            # fill up the target
            pag.write(DATA[col][line])
            pag.press('enter')
            
            cursor_pos = specs_labels.index(TARGET_SPECS[col]), TARGET_LINES[col][line] -1
            #time.sleep(2)
    
except:
    raise Exception("TABLE INSERT FAILED...")

#%% Calculate the model   

sourcename = 'TGBT'
displayMod(display='Tableur2',treeitem=sourcename)

#center ------ f8 to display the calculations pop-up 
pag.moveTo(centerpos[0],centerpos[1],0) 
pag.click(button='left')
pag.press('f8')
time.sleep(1)

try:
    win.CalculerButton.click()
    win.wait('visible')

    # skipping all warning windows until execution with a loop that waits until top window is Calcul global
    counter = 0
    
    popupcond = True
    while popupcond == True:
        try:
            win.click_input()
            win.OKButton.click()
            time.sleep(1)
        except:
            popupcond = False
       
            
    # clicking on the last popup windows before closing
    print('passed')
    
    popupcond = True
    while popupcond == True:
        try:
            win.FermerButton.click() 
            popupcond = False
        except:
            win.click_input()
            win.OKButton.click()
            time.sleep(1)
            
except:
    raise Exception("AUTOMATIC CALCULATIONS OF MODEL FAILED")
      
#%% Read output messages from the caneco console       

# make sure to run a calcul automatique before this function 

capsLockCheck()

pag.click(x=int(centerpos[0]),y=int(centerpos[1]+400))
pag.hotkey('ctrl','a')
time.sleep(1)
pag.hotkey('ctrl','c')
consoletxt_str = str(pyperclip.paste())

#parsing

consoletxt_bycalc = consoletxt_str.split('------------')
consoletxt_bycalc.pop(-1)
consoletxt_lastcalc = consoletxt_bycalc[-1]         
consoletxt_lines = consoletxt_lastcalc.split('\r\n')

indexlist = []
for line in range(len(consoletxt_lines)):
    if consoletxt_lines[line] == '':
        indexlist.append(line)
indexlist.reverse()

for line in indexlist:
    consoletxt_lines.pop(line) 

for line in range(len(consoletxt_lines)):
    consoletxt_lines[line] = consoletxt_lines[line].split('\t')
    while len(consoletxt_lines[line]) < 4:
        consoletxt_lines[line].append('n/a')
    
date,timestamp = consoletxt_lines[0][0].split(' ')[0],consoletxt_lines[0][0].split(' ')[1]
consoletxt_lines.pop(0)

# putting in an array

consoletxt_arr = np.concatenate((np.array([['GROUPS'],['LABELS'],['MESSAGE'],['PRIORITY']]).T,np.asarray(consoletxt_lines)),axis=0)


# creating a new text file containing the full console text

pag.click(x=int(centerpos[0]),y=int(centerpos[1]+400))
pag.hotkey('ctrl','a')
pag.click(button='right',x=int(centerpos[0]),y=int(centerpos[1]+400))
pag.press('down',presses=4)
pag.press('enter')    

browsePath(r"C:\Users\axel.streiff\OneDrive - ETE RESEAUX\Documents\Python Scripts",winpopup = win)   #might bug here !!!

pag.hotkey('ctrl','a')
pag.press('delete')

time.sleep(5)
pag.write('CONSOLE-RAW')
pag.press(['tab','down','enter','enter'])

# -----------------------------------------COLOUR RECOGNITION-------------------------------------------

time.sleep(1)
consolepath = os.path.join(folderpath,"CONSOLE-RAW.txt")
consoleRAW = open(consolepath,'r') 
consoleRAWSTR = consoleRAW.read()
consoleRAW.close()

# deleting the raw text file after it's been used
if os.path.exists(consolepath):
  os.remove(consolepath)

consoleRAWLIST = consoleRAWSTR.split('------------') 
consoleRAW_lastcalc = consoleRAWLIST[-2]
last_calc_list = consoleRAW_lastcalc.split('\par')

print(last_calc_list) 

# removing intermediary empty lines
indexlist = []
for line in range(len(last_calc_list)):
    if last_calc_list[line] == '' or last_calc_list[line] == '\n' or last_calc_list[line] == '\n\n' or last_calc_list[line] == '\n\\cf1' or last_calc_list[line] == '\n\\cf2':
        indexlist.append(line)
indexlist.reverse()

for line in indexlist:
    last_calc_list.pop(line) 
    
# removing date/timestamp and leftovers at the top  
indexlist = []
for line in range(len(last_calc_list)):
    if last_calc_list[line] != '\nSOURCE':
        indexlist.append(line)
    else:
        break
indexlist.reverse()

for line in indexlist:
    last_calc_list.pop(line) 
    
# and the bottom of the list 
last_calc_list.pop(-1)


# finding and storing the lines where there is a color switch
txtcolor = 'RED'
txtcolorlist = list()

# the colorstamps characters depend on the presence of an error
if consoleRAW_lastcalc.find('\cf4') == -1:    #NO ERRORS
    RED = '\cf3'
    BLUE = '\cf2'
    BLACK = '\cf1' 
else:                                         #ERRORS
    RED = '\cf4'
    BLUE = '\cf3'
    BLACK = '\cf2'
                                                    # PREVOIR COULEUR VIOLET 
for line in range(len(last_calc_list)):
    if last_calc_list[line].find(RED) != -1:
        txtcolor = 'RED'
        txtcolorlist.append([line,txtcolor])
        
    if last_calc_list[line].find(BLUE) != -1:
        txtcolor = 'BLUE'
        txtcolorlist.append([line,txtcolor])
        
    if last_calc_list[line].find(BLACK) != -1:
        txtcolor = 'BLACK'
        txtcolorlist.append([line,txtcolor])

# filling up the last column of consoletxt_arr for priority
inc = 0
colorstamp = 'BLACK'  # default font color

for line in range(1,np.shape(consoletxt_arr)[0]):
    
    if line == txtcolorlist[inc][0]+1:
        colorstamp = txtcolorlist[inc][1]
        
        if inc < len(txtcolorlist)-1:
            inc += 1 
    
    consoletxt_arr[line,3] = colorstamp  #filling up the last column
    
#%% OUTPUTS console reading SPYDER + ErrorLog folder

# THE NOTEPAD FONT MUST BE SET TO "Consolas" IN ORDER TO VISUALISE THE DATA ALIGNED

# Updating the output csv for the error log and the console warnings to .txt in the ErrorLog folder

# deleting previous iteration if the csv file exists
if os.path.exists(os.path.join(folderpath,'ErrorLog',"ConsoleOutput.csv")):
  os.remove(os.path.join(folderpath,'ErrorLog',"ConsoleOutput.csv"))

# saving new csv file
consoleFrame = pd.DataFrame(consoletxt_arr[1:,:],columns= consoletxt_arr[0,:])
consoleFrame.to_csv(os.path.join(folderpath,'ErrorLog',"ConsoleOutput.csv"),sep=';', index=False, encoding = 'utf-8-sig')

# saving the warnings in a .txt file

# deleting previous iteration if the text file exists
if os.path.exists(os.path.join(folderpath,'ErrorLog',"CONSOLE-REPORT.txt")):
  os.remove(os.path.join(folderpath,'ErrorLog',"CONSOLE-REPORT.txt"))

# writing in the text file
reportFileTXT = open(os.path.join(folderpath,'ErrorLog',"CONSOLE-REPORT.txt"),"w+")

reportFileTXT.write('////////////////////////////////////////CONSOLE OUTPUT FRAME////////////////////////////////////////\n\n')
array2txt(reportFileTXT,consoletxt_arr)
reportFileTXT.write('\n')

# sorted warnings
count = 0
warnings = list()
for line in range(1,np.shape(consoletxt_arr)[0]): 
    if consoletxt_arr[line,3] == 'RED':
        count += 1
        warnings.append(consoletxt_arr[line,2])
        
reportFileTXT.write(f'\n////////////////////////////////////////THE ROBOT FOUND {count} WARNING MESSAGES////////////////////////////////////////\n\n')
for line in warnings:
    reportFileTXT.write(f'(!)   ---   {line}')
    reportFileTXT.write('\n')
    
reportFileTXT.close()

# deleting previous iteration if the HTML file exists
if os.path.exists(os.path.join(folderpath,'ErrorLog',"CONSOLE-REPORT.html")):
  os.remove(os.path.join(folderpath,'ErrorLog',"CONSOLE-REPORT.html"))

# creating the html format 
txtfilecontent = open(os.path.join(folderpath,'ErrorLog',"CONSOLE-REPORT.txt"),"r")
reportFileHTML = open(os.path.join(folderpath,'ErrorLog',"CONSOLE-REPORT.html"),"w+")

for line in txtfilecontent.readlines():
    reportFileHTML.write("<pre>" + line + "</pre>\n")


reportFileHTML.close()
txtfilecontent.close()
    
#%% Manual closing of file before rerun 

reportFileTXT.close()
reportFileHTML.close()
txtfilecontent.close()

#%% Extract NDC in pdf version in NDCBOT folder

capsLockCheck()

if os.path.exists(os.path.join(folderpath,'NDCBOT',prefix + filename + '.pdf')):
  os.remove(os.path.join(folderpath,'NDCBOT',prefix + filename + '.pdf'))
  print("Deleted previous version of file : ",prefix + filename + '.pdf')
  
buttonVisualClick('ImpressionPanel')
time.sleep(0.5)
buttonVisualClick('MiseEnPage')

time.sleep(3)
win = caneco.top_window()
win['DOSSIER IRVETreeItem'].click_input()

win['Aperçu ...Button'].click()
time.sleep(6)
buttonVisualClick('Imprimer')
time.sleep(6)
win = caneco.top_window()
win['Nom :ComboBox'].click_input()
time.sleep(0.5)
win['Microsoft Print to PDF'].click_input()
time.sleep(0.5)
win.OKButton.click()
time.sleep(3)
win = caneco.top_window()
browsePath(os.path.join(folderpath,'NDCBOT'),winpopup = win)  #browsing for the NDCBOT folder 
pag.write(prefix + filename)
time.sleep(1)
pag.press('enter')

search = True
while search:
    if pag.locateCenterOnScreen(os.path.join(folderpath,'ImgButtons','ApercuFermer.PNG'),grayscale = True) != None:
      buttonVisualClick('ApercuFermer')
      search = False
    
#%% Exit the model (save or not) 

save = True

# erase previous version if it exists 
if os.path.exists(os.path.join(folderpath,'NDCBOT',prefix + filename + '.afr')):
  os.remove(os.path.join(folderpath,'NDCBOT',prefix + filename + '.afr'))
  print("- Deleted previous version of file : ",prefix + filename + ".afr")
if os.path.exists(os.path.join(folderpath,'NDCBOT',prefix + filename + '.rap')):
  os.remove(os.path.join(folderpath,'NDCBOT',prefix + filename + '.rap'))
  print("- Deleted previous version of file : ",prefix + filename + ".rap")
  
pag.click(x=1892, y=42)                 # exact position of the quit button

# win popup confirm exit and save
if save == True:
    pag.press('enter',presses=3,interval=2)
    print('- Model saved and closed')
else:
    pag.press('tab')
    pag.press('enter')
    print('- Model closed without saving')

time.sleep(1)
# ends up in C:\Users\axel.streiff\OneDrive - ETE RESEAUX\Documents\Caneco BT and must be moved to the NDCBOT folder
sourceFolder = r'C:\Users\axel.streiff\OneDrive - ETE RESEAUX\Documents\Caneco BT'
destFolder = os.path.join(folderpath,'NDCBOT')

try:
    shutil.copy(os.path.join(sourceFolder,prefix + filename +'.afr'),destFolder)
    shutil.copy(os.path.join(sourceFolder,prefix + filename +'.rap'),destFolder)
    print("- Migration of the .afr and .rap files : done")
except:
    raise Exception("Migration of the .afr and .rap files : failed")


#removing the files from the caneco default folder 
time.sleep(1)
try:
    os.remove(os.path.join(sourceFolder,prefix + filename + '.afr'))
    os.remove(os.path.join(sourceFolder,prefix + filename + '.rap'))
    print("- Removing of the .afr and .rap files : done")
except:
    raise Exception("Removing of the .afr and .rap files : failed")

#%% Quit Caneco

#Force quit without saving
caneco.kill()
print("Caneco was killed")

# or   win['FermerButton'].click()

#%% Mouse position test
while True:
    print(pag.position())
    time.sleep(10)

#%% 
win = caneco.top_window()
win.print_control_identifiers()

