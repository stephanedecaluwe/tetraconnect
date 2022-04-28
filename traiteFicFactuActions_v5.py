import glob
import locale
import os
import re
import sys
from datetime import datetime
import time
import tablib
import logging
import xlrd
import shutil
from typing import ClassVar
import colorlog
from dataclasses import dataclass,field,asdict,astuple

############## A MODIFIER
# Les terminaux de AVEM_CAPS dans le parc "Campagne Ingenico" ne sont pas facturés (mise à jour vulnérabilité T2)
# A partir de Q4 2021, il faudra modifier la facturation pour la formule basic
ignoreList = ['ingAvantVentes','ingAvantVentes2','ingTests','QA','PGC_SG','autotestING','LaPoste','DEMO','ING','ING2','BETA_TEST_VALENCE','CMCIC','CONECS']

pathFichierTarif =  r'C:\Users\sdecaluwe\Desktop\factuTC\TARIFICATION_TETRA_CONNECT.xlsx'

dossierClients   =  r'C:\Users\sdecaluwe\Desktop\factuTC\factuQ1-2022'

annee = 2022
trimestre = 1

############## FIN A MODIFIER

locale.setlocale(locale.LC_ALL, '') #set local en Français

RegexActionFileName = re.compile(r"Actions_([^.]*)_202.-[^.]*\.xlsx",re.IGNORECASE )
RegexTPEFileName    = re.compile(r"([^.]*)_202.-[^.]*\.xlsx",re.IGNORECASE )


traducNomAction = { 'init':{'n':'Initialisation'},
                    'init_mass':{'n':'Initialisation en masse'},
                    'blockpay':{'n':'Blocage/deblocage paiement'}, 
                    'tms_params': {'n':'Paramétrage TMS','free':True}, #Action paramétrage TMS gratuite Planet..
                    'set_coms':{'n':'Paramétrage moyens de com'},
                    'download_order':{'n':'Ordre téléchargement','free':True},#gratuit pour toutes les formules
                    'french_config':{'n':'Paramétrage passerelle/con caisse'},
                    'message': {'n':'Message','free':True}, #sauf en Basic où c'est payant
                    'screen_saver':{'n':'Economiseur écran'}, 
                    'del_instances': {'n':'Suppression instances'},
                    'delete_components':{'n':'Suppression applications','free':True},
                    'writeFileHost':{'n':'Ecriture fichier paramétrage'},
                    'createSSL':{'n':'Création SSL Lyra/Paybox'},
                    'writeBinFile':{'n':'Ecriture paramètre automate' },
                    'launchTLP': {'n':'Lancement téléparamétrages' },
                    'printAmountToTransmit': {'n':'Impression totaux à transmettre' },
                    'testIPportConnexion' :  {'n':'Test connexion IPport' }
                    }

# traducNomAction = { 'init':{'n':'Initialisation'},
                    # 'init_mass':{'n':'Initialisation en masse'},
                    # 'blockpay':{'n':'Blocage/deblocage paiement'}, 
                    # 'tms_params': {'n':'Paramétrage TMS','free':True}, #Action paramétrage TMS gratuite Planet..
                    # 'set_coms':{'n':'Paramétrage moyens de com'},
                    # 'download_order':{'n':'Ordre téléchargement','free':True},
                    # 'french_config':{'n':'Paramétrage passerelle/con caisse'},
                    # 'message': {'n':'Message','free':True}, #sauf en Basic où c'est payant
                    # 'screen_saver':{'n':'Economiseur écran'}, #screensaver temporairement comme TCONNECT supporte pas
                    # 'del_instances': {'n':'Suppression instances'},
                    # 'delete_components':{'n':'Suppression applications','free':True},
                    # 'writeFileHost':{'n':'Ecriture fichier paramétrage'},
                    # 'createSSL':{'n':'Création SSL Lyra/Paybox'},
                    # 'writeBinFile':{'n':'Ecriture paramètre automate'},
                    # 'testIPportConnexion':{'n':'Test port IP','free':True},
                    # 'launchTLP':{'n':'launchTLP','free':True}
                    # }
                    
#blockpay
#createSSL
#del_instances
#delete_components
#download_order
#french_config
#init
#message
#screen_saver
#set_coms
#testIPportConnexion
#tms_params

                    

def GetExcelFileLines(filePath, sheetName:str=None):
    """Lit un fichier Excel et retourne un tableau des objets lignes.
    """
    logging.info(f'Lecture fichier {os.path.basename(filePath)}')
    myBook = xlrd.open_workbook(filePath)

    if sheetName is None:
        sheet = myBook.sheet_by_index(0)

    if sheet.nrows <2:
        return None

    dic ={num: col.value for num,col in enumerate(sheet.row(0) ,start=0)}

    return [{dic[num]:col.value for num,col in enumerate(sheet.row(iLigne) ,start=0)} for iLigne in range(1,sheet.nrows)]

def SetLogging(logToConsole=True):
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    
    logging.getLogger("requests").setLevel(logging.WARNING)
    logging.getLogger('urllib3').setLevel(logging.WARNING)
    #Console Logger
    if logToConsole:
        consoleHandler = logging.StreamHandler()
        consoleFormatter = colorlog.ColoredFormatter(
            "%(log_color)   s%(message)s",
            datefmt=None,
            reset=True,
            log_colors={
                'DEBUG':    'cyan',
                'INFO':     'green',
                'WARNING':  'yellow',
                'ERROR':    'red',
                'CRITICAL': 'red',
            }
        )
        consoleHandler.setFormatter(consoleFormatter)
        consoleHandler.setLevel(logging.INFO)
        
        if (logger.hasHandlers()):
            logger.handlers.clear()

        logger.addHandler(consoleHandler)

    #File Logger
    parentDir =os.path.abspath(os.path.join(os.path.dirname( __file__), os.pardir))

    dossierLogs = os.path.join(parentDir, 'RES','logs')
    os.makedirs(dossierLogs,exist_ok=True)

    logFilePath= os.path.join( dossierLogs, f'{time.strftime("%Y%m%d_%Hh%M")}_logs.log')

    fileHandler =logging.FileHandler(filename=logFilePath, mode='a', encoding="utf-8", delay=False)
    fileHandler.setFormatter( logging.Formatter('[%(levelname)s]: %(message)s') )
    fileHandler.setLevel(logging.DEBUG)
    logger.addHandler(fileHandler)
    return logFilePath

@dataclass
class Tarif:
    NomClient:str
    NomSAP:str
    CodeSAP:str
    EstGrossiste: bool
    Gratuit: bool
    SAPToPrix: dict = field(default_factory=dict,init=False)

@dataclass
class TPE:
    SN: str
    codeParc: str
    nomParc:str
    formule:str
    frequence:str
    formFreqChange: str
    PN: str
    LastSyncDate: object #datetime

    @property
    def XlLine(self):
        return {"SN":self.SN, "PN":self.PN, "CodeParc":self.codeParc, "NomParc":self.nomParc, "Formule":self.formule,"Frequence":self.frequence,"DernièreSynchro":self.LastSyncDate}

@dataclass
class Action:
    SN: str
    nomAction: str
    nomParc:str
    #dateCreation: datetime inutile
    Resultat:str
    formule:str
    dateRealisation: datetime
    contenu: str

    ssActPayantes: list     = field(default_factory=list,init=False)
    ssActGratuites: list    = field(default_factory=list,init=False)

    @staticmethod
    def traduireListeActions(listAtraduire):
        return  [traducNomAction[ elt.strip('"') ]['n'] for elt in listAtraduire]

    @property
    def XlLine(self):
        return {"SN":self.SN,"nomAction":self.nomAction, "nomParc":self.nomParc, "Résultat":self.Resultat,"Formule":self.formule,
        "dateRéalisation":self.dateRealisation,"SsActGratuites":self.ContGratuitAction, "SsActPayantes":self.ContPayantAction,
        "NbSsActPayantes":self.NbssActPayantes,"NbSsActFacturees":self.NbssActFacturees}

    @property
    def EstPayante(self):
        return self.Resultat == "SUCCESS" and self.NbssActPayantes>0

    @property
    def ContGratuitAction(self):
        return "-".join(Action.traduireListeActions(self.ssActGratuites))

    @property
    def ContPayantAction(self):
        return "-".join(Action.traduireListeActions(self.ssActPayantes))

    @property
    def NbssActGratuites(self):
        return len(self.ssActGratuites)

    @property
    def NbssActPayantes(self):
        return len(self.ssActPayantes)

    @property
    def NbssActFacturees(self):
        nb =  self.NbssActPayantes
        if not nb:
            return 0

        return min(3,nb )

    def __post_init__(self):
        assert self.formule in ['Basic','Prémium','Ultimate']

        listActions = self.contenu.replace('"','').split(',')
        if len(listActions) == 0:
            logging.error("Action vide")
            sys.exit()
            
        # message,screensaver, init, init_mass payant qu'en formule basic nouveauté Q4 2021
        traducNomAction['message']['free']      = (self.formule != 'Basic')
        traducNomAction['init']['free']         = (self.formule != 'Basic')
        traducNomAction['init_mass']['free']    = (self.formule != 'Basic')
        traducNomAction['screen_saver']['free'] = (self.formule != 'Basic')

        if self.formule == 'Ultimate': #En Ultimate toutes les actions sont gratuites
            self.ssActGratuites = listActions.copy()
        else:
            for act in listActions:
                prop = traducNomAction[act]

                if prop.get('free',False):                 
                    self.ssActGratuites.append(act)                    
                else:
                    self.ssActPayantes.append(act)

@dataclass
class DataGrossiste:
    Nom:str
    Tarif: Tarif
    ParentAllData:  object #AllData

    SsClients: dict = field(default_factory=dict,init=False)
    DataCeGrossite: object = field(default=None,init=False)
    DossierCeGrossiste:str = field(default="",init=False)
    DossierDetailsCeGrossiste:str = field(default="",init=False)
    
    @property
    def DossierCetteFactu(self):
        return self.ParentAllData.DossierCetteFactu

    @property
    def NomGr(self):
        return self.Nom +"_grossiste"

    def __post_init__(self):
        self.DataCeGrossite = DataUnClient(self.NomGr,self.Tarif  ,self.ParentAllData, IsGrossiste=True)

        self.DossierCeGrossiste = os.path.join(self.DossierCetteFactu,self.NomGr )
        os.makedirs(self.DossierCeGrossiste, exist_ok=True )

        self.DossierDetailsCeGrossiste = os.path.join(self.DossierCeGrossiste,"Détails" )
        os.makedirs(self.DossierDetailsCeGrossiste, exist_ok=True )

    def getFactuLines(self,includeSAP= False):
        return [l.getLineFactu(includeSAP) for l in self.SsClients.values() ] + [self.getFactuLine(includeSAP)]

    def getFactuLine(self, includeSAP= False):
        return self.DataCeGrossite.getLineFactu(includeSAP)

    def ajouterSsClient(self,clientName):
        dataCeClient = DataUnClient(clientName, self.Tarif, self.ParentAllData, IsGrossiste=False)
        dataCeClient.DataParentGrossiste = self
        self.SsClients[clientName] = dataCeClient

    def EcrireFichiersExcel(self):
        #fichier global grossiste
        logging.info(f"Génération Excel global '{self.NomGr}'")
        excelFilePath = os.path.join(self.DossierCeGrossiste,f"{self.ParentAllData.nowString}_{self.NomGr}_factuGlobale_{self.ParentAllData.TrimAnneeStr}.xlsx")
        ecrireFichierExcel(excelFilePath, self.getFactuLines() , f"{self.NomGr}_Globale")

        #fichier Détails sous clients
        for ssclient in self.SsClients.values():
            ssclient.generateExcelDetails()

        #Genere 1 zip ce grossiste
        shutil.make_archive(os.path.join(self.DossierCetteFactu,f"{self.ParentAllData.TrimAnneeStr}_{self.NomGr}"),format="zip",root_dir= self.DossierCeGrossiste)

@dataclass
class DataUnClient:
    NomClient: str
    Tarif: Tarif
    ParentAllData:  object #AllData

    IsGrossiste:bool            = field(default=False)

    UniqueSNs: set              = field(default_factory=set,init=False)
    TPEs:list                   = field(default_factory=list,init=False)
    NbTpesParFormFreq:dict      = field(default_factory=dict,init=False)
    
    Actions:list                        = field(default_factory=list,init=False)
    totalNbActionsFacturees:int             = field(default=0, init=False)
    totalNbActionsPayantesAvantPlafond:int  = field(default=0, init=False)
    NbActionsEchouees:int                   = field(default=0, init=False)

    DataParentGrossiste: object   = field(default=None, init=False)
    _DossierDetails:str            = field(default="",init=False)


    FormFreqToSAPCode:ClassVar[dict]= {
                ("Basic","Semestrielle"):"TC_BASIC_MPE",
                ("Basic","Mensuelle"):"TC_BASIC_MPE",
                ("Prémium","Mensuelle"):"TC_PRE_MOIS_MPE",
                ("Prémium","Hebdomadaire"):"TC_PRE_SEM_MPE",
                ("Prémium","Quotidien"):"TC_PRE_JOUR_MPE",
                ("Ultimate","Mensuelle"):"TC_ULT_MOIS_MPE",
                ("Ultimate","Hebdomadaire"):"TC_ULT_SEM_MPE",
                ("Ultimate","Quotidien"):"TC_ULT_JOUR_MPE"
            }

    def __post_init__(self):
        for (form,freq) in DataUnClient.FormFreqToSAPCode.keys():
            self.NbTpesParFormFreq[(form,freq) ] = 0

    @property
    def DossierDetails(self):
        if not self._DossierDetails:
            if self.DataParentGrossiste:
                self._DossierDetails = self.DataParentGrossiste.DossierDetailsCeGrossiste
            else:
                self._DossierDetails = os.path.join( self.DossierCetteFactu,"Détails")
                os.makedirs(self._DossierDetails,exist_ok=True )

        return self._DossierDetails

    @property
    def DossierCetteFactu(self):
        return self.ParentAllData.DossierCetteFactu

    def getLineFactu(self,includeSAP=False):
        res = {"Nom_Client":self.NomClient }
        if includeSAP:
            res["Nom_SAP"] = ""
            res["Code_SAP"] = ""

        res["TC_ACTION_ACTE"] = self.totalNbActionsFacturees

        for ff,v in self.NbTpesParFormFreq.items():
            if DataUnClient.FormFreqToSAPCode[ff] in res:
                res[DataUnClient.FormFreqToSAPCode[ff]] += v
            else: 
                res[DataUnClient.FormFreqToSAPCode[ff]] = v

        return res

    def generateExcelDetails(self):
        if self.DataParentGrossiste:
            logging.info(f"Génération Excel Détails {self.DataParentGrossiste.Nom}\{self.NomClient}")
        else:
            logging.info(f"Génération Excel Détails {self.NomClient}")

        books = tablib.Databook()

        syntheseUsed = self.getLineFactu()

        for ch in ['Nom_SAP','Code_SAP',"A_facturer_[EUR]","Est_grossiste"]:
            if ch in syntheseUsed:
                del syntheseUsed[ch]

        tabSynthese = tablib.Dataset(title='Synthèse', headers= syntheseUsed.keys() )
        tabSynthese.append( syntheseUsed.values() )
        books.add_sheet(tabSynthese)

        if self.TPEs:
            tpes = None
            for t in self.TPEs:
                if not tpes:
                    tpes = tablib.Dataset(title='Terminaux',headers=t.XlLine.keys() )

                tpes.append(t.XlLine.values())

            books.add_sheet(tpes)
        else:
            logging.warning(f"Pas de fichier TPEs sur {self.NomClient}")
            

        if self.Actions:
            tabActions = None
            for a in self.Actions:
                if not tabActions:
                    tabActions = tablib.Dataset(title='Actions', headers= a.XlLine.keys() )

                tabActions.append(a.XlLine.values())

            books.add_sheet(tabActions)
        else:
            logging.warning(f"Pas de fichier d'action sur {self.NomClient}")
 
        filePath = os.path.join( self.DossierDetails, f"{self.ParentAllData.nowString}_Détails_{self.NomClient}_{self.ParentAllData.TrimAnneeStr}.xlsx")
        try:
            ecrireFichierGlobal(books, filePath)
        except PermissionError:
            input(f"Fermer le fichier {filePath} et appuyer sur O")
            ecrireFichierGlobal(books, filePath)

    def ajouteTPE(self,tpe):
        if not tpe.formule:
            logging.error(f"Formule vide {tpe.SN} {self.NomClient}")
        elif self.NomClient == "AVT":
            logging.warning(f"TPE skippé car AVT (test)")
        else:
            self.UniqueSNs.add(tpe.SN)
            self.TPEs.append(tpe)
            self.NbTpesParFormFreq[(tpe.formule,tpe.frequence)] += 1
            if self.DataParentGrossiste:
                self.DataParentGrossiste.DataCeGrossite.NbTpesParFormFreq[(tpe.formule,tpe.frequence)] += 1

    def ajouteAction(self,action):        
        if self.NomClient == "AVT":
            logging.warning(f"Action skippée car AVT (test)")
        elif action.Resultat != "SUCCESS":
            self.NbActionsEchouees += 1
        else:
            if not self.TPEs:
                logging.error(f"Pas de tpes pour client {self.NomClient} alors que actions presentes")
                if input("x pour quitter").lower() == 'x':
                    sys.exit()
            
            if action.SN not in self.UniqueSNs:
                if action.EstPayante:
                    logging.error(f'TPE {action.SN} present dans action mais pas dans TPE factures pour client {self.NomClient}')
                    if input("x pour quitter").lower() == 'x':
                        sys.exit()
                else:
                    return #pas grave, l'action n'est pas payante

            self.Actions.append(action)

            self.totalNbActionsFacturees            += action.NbssActFacturees
            self.totalNbActionsPayantesAvantPlafond += action.NbssActPayantes

            if self.DataParentGrossiste:
                self.DataParentGrossiste.DataCeGrossite.totalNbActionsFacturees            += action.NbssActFacturees
                self.DataParentGrossiste.DataCeGrossite.totalNbActionsPayantesAvantPlafond += action.NbssActPayantes

@dataclass
class AllData:
    DossierBaseResFactu:str
    Annee:int
    Trimestre:int

    DataParClientGrossiste:dict     = field(default_factory=dict,init=False)
    DataParClientNonGrossiste:dict  = field(default_factory=dict,init=False)
    DossierCetteFactu: str          = field(default="",         init=False)
    
    def __post_init__(self):
        assert 1<=self.Trimestre <=4
        self.DossierCetteFactu = os.path.join(self.DossierBaseResFactu, f"{self.nowString}_factu_Q{trimestre}-{annee}" )
        os.makedirs(self.DossierCetteFactu, exist_ok=True )
    
    @property
    def debutTrim(self):
        return datetime(self.Annee,1+(self.Trimestre-1)*3 , 1,0,0,0,0)
    
    @property
    def finTrim(self):        
        return datetime(self.Annee,self.Trimestre*3 ,31 if (self.Trimestre == 1 or self.Trimestre==4) else 30 ,hour=23, minute=59, second=59, microsecond=999999)

    @property
    def nowString(self):
        return datetime.now().strftime("%Y%m%d_%Hh%M") 

    @property
    def TrimAnneeStr(self):
        return f"Q{self.Trimestre}_{self.Annee}"

    def lireFichierTarif(self):
        """ Lit le fichier tarif
            :param pathFichierTarif: Chemin fichier tarif
            :return: {'dicNomClientVersTarif_Grossistes': {'tarif':dicNomClientVersTarif_Grossistes},'dicNomClientVersTarif_NonGrossistes':{'tarif': dicNomClientVersTarif_NonGrossistes}}
        """
        logging.info('Lecture fichier tarif')

        # ouverture du fichier Excel
        with open(pathFichierTarif,'rb') as fh:
            dataTarif = tablib.Dataset().load(fh,'xlsx')

        if not dataTarif:
            return logging.error(f'fichier vide: {pathFichierTarif}')

        for vals in dataTarif.dict:
            nomClient   = vals['NomClientSousTetraConnect'].strip()
            isGrossiste = vals['Est_grossiste'].lower().strip() == 'oui'

            if nomClient in ignoreList:
                continue
            
            tarif = Tarif(nomClient,vals['Nom_SAP'],vals['Code_SAP'],isGrossiste,vals['Gratuit'].lower().strip() == 'oui' )

            for k,v in vals.items():
                if k and k.startswith('TC_'):
                    tarif.SAPToPrix[k] = v

            if isGrossiste:
                self.DataParClientGrossiste[nomClient] = DataGrossiste(nomClient,tarif,self)
            else:
                self.DataParClientNonGrossiste[nomClient] = DataUnClient(nomClient,tarif,self)

    def getDataForClient(self,clientName):
        if not self.DataParClientNonGrossiste:
            self.lireFichierTarif()

        for nameGrossiste in self.DataParClientGrossiste.keys():
            if clientName.startswith(nameGrossiste):
                dataGr = self.DataParClientGrossiste[nameGrossiste]

                if clientName not in dataGr.SsClients:
                    dataGr.ajouterSsClient(clientName)

                return dataGr.SsClients[clientName]

        if clientName not in self.DataParClientNonGrossiste:
            logging.error(f'Client {clientName} non présent dans les tarifs finaux')
            sortirSiConfirme()

            self.DataParClientNonGrossiste[clientName] = DataUnClient(clientName, Tarif(clientName, "","", False,False),self)

        return self.DataParClientNonGrossiste[clientName]

    def generateAllFactu(self):
        listForGlobalXl = []

            #Grossistes
        for gr in dataCetteFactu.DataParClientGrossiste.values():
            gr.EcrireFichiersExcel()
            listForGlobalXl.append( gr.getFactuLine(includeSAP=True) )

            #Non grossistes
        for cl in dataCetteFactu.DataParClientNonGrossiste.values():
            cl.generateExcelDetails()
            listForGlobalXl.append( cl.getLineFactu(includeSAP=True) )

        filePath = os.path.join(self.DossierCetteFactu,f"{self.nowString}_factuGlobale_{self.TrimAnneeStr}.xlsx")
        ecrireFichierExcel(filePath,listForGlobalXl,title=self.TrimAnneeStr )
        
    def litFichier(self, filePath, isTerminal):
        NomDuClient = ""
        
        fileNameOnly = os.path.basename(filePath)

        if isTerminal:
            m = RegexTPEFileName.match(fileNameOnly) #   re.search('[^.]*',fileNameOnly)
        else:
            m = RegexActionFileName.match(fileNameOnly)
         
        if not m:
            logging.error(f'Pas de matching nom fichier {fileNameOnly}')
            sys.exit()

        NomDuClient = m.group(1).strip()
        if NomDuClient in ignoreList:
            logging.warning(f'ignoré {fileNameOnly}')
            return None

        dataCeClient = self.getDataForClient(NomDuClient)
        nbHorsTrim =0
        nbSkipCampagneIng =0

        lines = GetExcelFileLines(filePath)

        if( not lines):
            return logging.error(f'fichier vide: {filePath}')

        formatDateExcel = "%d/%m/%Y %H:%M"
        trimestreString = f"{dataCetteFactu.Annee} Trimestre {dataCetteFactu.Trimestre}"

        if isTerminal:
            for line in lines:
                # if line['Trimestre'] != trimestreString:
                #     logging.error(f"Mauvais trimestre fichier terminal {line} {fileNameOnly}")
                #     sys.exit()

                if NomDuClient=='AVEM_CAPS' and line['nom du Parc'] =='Campagne Ingenico': #Ne pas facturer campagne Ingetrust AVEM
                    nbSkipCampagneIng +=1
                    continue

                assert line['Nom du client'].strip() == NomDuClient,f"Noms différents '{line['Nom du client']}' / '{NomDuClient}'"

                lastSync = datetime.strptime(line['Dernière Synchronisation du trimestre'], formatDateExcel)
                
                # if lastSync < self.debutTrim or lastSync > self.finTrim: #En dehors du trimestre
                #     #logging.warning(f"Date last sync en dehors du trimestre")
                #     nbHorsTrim +=1
                    
                dataCeClient.ajouteTPE( TPE(line['Identifiant du terminal'],line['Code Parc'], line['nom du Parc'],line['Formule courante'],line['Fréquence courante'], line['formfreq changes'],line['PN'], lastSync) )      
        else:
            for l in lines:
                if l["Nom du Parc"]=="Campagne Ingenico" and NomDuClient =="AVEM_CAPS": #Pas de factu Ingetrust
                    nbSkipCampagneIng +=1
                    continue

                if not l['Formule']:
                    logging.warning(f"Formule vide action {l}")
                    sys.exit()
                    continue

                resuAction = l["Résultat de l'action"]
                assert resuAction in ["SUCCESS", "ERROR"],f"Mauvais statut action {resuAction}" #3 derniers statuts cf ticket #950
                #assert resuAction in ["SUCCESS","AUTO_CLOSED","ERROR", "OUTDATED","PLANNED","PENDING"],f"Mauvais statut action {resuAction}" #3 derniers statuts cf ticket #950

                dateRéa = l["Date de réalisation utc"]

                if not dateRéa:
                    logging.warning("Pas de date de réalisation de l'action")
                    sys.exit()
                    continue

                dateRead = datetime.strptime(dateRéa, formatDateExcel)
                action = Action(l['SN du terminal'],l['name'],l["Nom du Parc"],resuAction,l["Formule"],dateRead,l["Contenu de l'action"] )
                # if dateRead < self.debutTrim or dateRead > self.finTrim: #En dehors du trimestre
                #     #logging.warning(f"Date réa action en dehors du trimestre {l}")
                #     if action.EstPayante:
                #         nbHorsTrim +=1
                #     continue

                dataCeClient.ajouteAction(action )
                
        if nbHorsTrim:
            logging.warning(f"Nb hors Trim => {nbHorsTrim}")

        if nbSkipCampagneIng:
            logging.warning(f"Nb skip camp Ingenico => {nbSkipCampagneIng}")

def sortirSiConfirme():
    if input("Taper x pour quitter").lower() == 'x':
        sys.exit()

def traitementFactu(dataCetteFactu: AllData):
    for startingAct in [False,True]:
        if startingAct:
            fichiersTRouvés = glob.glob( os.path.join(dataCetteFactu.DossierBaseResFactu, "Actions_*.xlsx") )
        else:
            fichiersTRouvés = glob.glob( os.path.join(dataCetteFactu.DossierBaseResFactu, "*.xlsx") )

        for fic in fichiersTRouvés:
            fileNameOnly = os.path.basename(fic)
           
            if fileNameOnly.startswith('~') or (not startingAct and fileNameOnly.startswith('Actions_') ):
                continue

            dataCetteFactu.litFichier(fic, not startingAct)

    #on a lu toutes les données, exploitation commence:
    dataCetteFactu.generateAllFactu()

def ecrireFichierGlobal(books, excelFilePath):
    if books.size ==0:
        logging.error("Pas d onglets Excel pour " + excelFilePath)
        input("pause")
        return

    with open(excelFilePath, 'wb') as f:  #PermissionError si déjà ouvert
        f.write(books.export('xlsx'))

def ecrireFichierExcel(excelFilePath, objList, title):
    tab = tablib.Dataset(headers=objList[0].keys(), title=title)

    for row in objList:
        tab.append(row.values())

    with open(excelFilePath, 'wb') as f:  #PermissionError si déjà ouvert
        f.write(tab.export('xlsx'))

if __name__ == '__main__':
    dataCetteFactu = AllData(dossierClients, annee, trimestre)

    SetLogging(logToConsole=True)
    traitementFactu(dataCetteFactu)
