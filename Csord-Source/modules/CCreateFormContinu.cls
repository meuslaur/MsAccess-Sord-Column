VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCreateFormContinu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name:     CCreateFormContinu
' Kind:     Module de classe
' Purpose:  Création d'un formulaire pour la classe CsordFormColumn
' Author:   Laurent
' Date:     27/04/2022
' DateMod:  11/05/2022-20:46
' ------------------------------------------------------
Option Compare Database
Option Explicit

'//::::::::::::::::::::::::::::::::::    VARIABLES      ::::::::::::::::::::::::::::::::::
    Private Const MOD_CLASS     As String = "CSordFormColumn"   '// Nom du module de classe à importer.

'// Options defaut.
    Private Const CLS_VAR       As String = "m_CSordForm"   '// Name variable for create class 'CsordFormColumn'
    Private Const PROC_NAME     As String = "SordColumn"    '// Nom de la function pour lancer le tri.
    
    Private Const CODE_FRM      As String = "CodeForm.cls"  '// Fichiers pour import code dans le formulaire et la classe.
    Private Const CODE_CLASS    As String = "CodeClasse.cls"

    Private Const PREF_CMB      As String = "cmb_"      '// Préfixe pour le commandbutton.
    Private Const SUF_CMB       As String = "_cmb"      '// Suffixe pour le commandbutton.
    Private Const PREF_TXTBOX   As String = "txt_"      '// Préfixe pour la textebox.
    Private Const SUF_TXTBOX    As String = "_txt"      '// Suffixe pour la textebox.

    Private m_oMsApp            As Access.Application

    '// Options définies dans le formulaire.
    Private m_oMsBase           As DAO.Database
    Private m_sBaseFullName     As String   '// Sauvegarde du chemin\nom de la base ouverte.
    Private m_bMsAppIsUp        As Boolean
    Private m_bMsBaseIsOpen     As Boolean

    Private m_sOptFormName      As String
    Private m_sOptFormSource    As String
    Private m_sFormFields()     As String   '// Stock les champs à utiliser dans le form.
    Private m_lFieldIndex       As Long     '// Index du tableau des champs.

    Private m_sOptVarClasse     As String   '// Nom de la variable pour initialiser la classe.
    Private m_sOptFunctionName  As String

    Private m_sOptPictureFolder As String
    Private m_sOptPictureAsc    As String
    Private m_sOptPictureDesc   As String

    Private m_sOptCmbPrefix     As String
    Private m_sOptCmbSuffix     As String

    Private m_sOptTextBoxSuffix As String
    Private m_sOptTextBoxPrefix As String

<<<<<<< HEAD
    Public m_cUtils As New CUtilitaires     '// FSO
=======
    Private m_cUtils As New CUtilitaires
>>>>>>> 90b235b329fe03c39da9bc499914478247f26c39
'//:::::::::::::::::::::::::::::::::: END VARIABLES ::::::::::::::::::::::::::::::::::::::
    
'//----------------------------------       INIT/TERM       ----------------------------------
Private Sub Class_Initialize()

    Dim CsordFrm As CsordFormColumn

    '// Défini les options par défaut.
    m_sOptVarClasse = CLS_VAR
    m_sOptFunctionName = PROC_NAME

    '// Extraire les info par défaut de la classe 'CsordFormColumn'
    Set CsordFrm = New CsordFormColumn
    m_sOptPictureFolder = CsordFrm.PicturePath  '/MODE:
    m_sOptPictureAsc = CsordFrm.PictureASC
    m_sOptPictureDesc = CsordFrm.PictureDESC
    Set CsordFrm = Nothing

    m_sOptCmbPrefix = PREF_CMB
    m_sOptCmbSuffix = SUF_CMB

    m_sOptTextBoxPrefix = PREF_TXTBOX
    m_sOptTextBoxSuffix = SUF_TXTBOX

    m_lFieldIndex = -1      '// Indique tableau non initialiser.

End Sub

Private Sub class_Terminate()

    If (MsBaseIsOpen And (Not m_oMsBase Is Nothing)) Then m_oMsBase.Close
    If (MsAppIsUp And (Not m_oMsApp Is Nothing)) Then m_oMsApp.Quit
    DoEvents

    Set m_oMsBase = Nothing
    Set m_oMsApp = Nothing
    Set m_cUtils = Nothing

End Sub
'//---------------------------------------------------------------------------------------

'//==================================       PROP        ==================================

Public Property Get objMsBase() As DAO.Database
    Set objMsBase = m_oMsApp.CurrentDb
End Property

Public Property Get MsAppIsUp() As Boolean
    MsAppIsUp = m_bMsAppIsUp
End Property
Public Property Let MsAppIsUp(bNewVal As Boolean)
    m_bMsAppIsUp = IIf(m_oMsApp Is Nothing, False, bNewVal)
End Property

Public Property Get MsBaseIsOpen() As Boolean
    MsBaseIsOpen = m_bMsBaseIsOpen
End Property
Public Property Let MsBaseIsOpen(bNewVal As Boolean)
    m_bMsBaseIsOpen = IIf((m_oMsApp.CurrentProject.AccessConnection Is Nothing), False, bNewVal)
End Property

Public Property Get GetBaseFullName() As String
    GetBaseFullName = m_sBaseFullName
End Property

Public Property Get OptFormName() As String
    OptFormName = m_sOptFormName
End Property
Public Property Let OptFormName(NewValue As String)
    m_sOptFormName = NewValue
End Property

Public Property Get OptFormSource() As String
    OptFormSource = m_sOptFormSource
End Property
Public Property Let OptFormSource(sNewValue As String)
    m_sOptFormSource = sNewValue
End Property

Public Property Get GetFields() As String()
    '// Retourne le tableau des Fields sélectionnés.
    GetFields = m_sFormFields
End Property

'//     m_sFormFields()    '// Stock les champs à utiliser dans le from.
'//     m_iFieldIndex      '// Index(démarre à -1 donc 0) du tableau des champs.
Public Property Let AddField(sFieldName As String)
    If (Len(sFieldName) = 0) Then Exit Property

    If (m_lFieldIndex <= UBound(m_sFormFields)) Then
        m_lFieldIndex = m_lFieldIndex + 1
        ReDim Preserve m_sFormFields(0 To m_lFieldIndex) As String
    End If
    m_sFormFields(m_lFieldIndex) = sFieldName

End Property

Public Property Get ClasseModuleName() As String
    ClasseModuleName = MOD_CLASS
End Property
Public Property Get OptVarClasse() As String
    OptVarClasse = m_sOptVarClasse
End Property
Public Property Let OptVarClasse(sNewValue As String)
    m_sOptVarClasse = sNewValue
End Property

Public Property Get OptFunctionName() As String
    OptFunctionName = m_sOptFunctionName
End Property
Public Property Let OptFunctionName(ByVal sNewValue As String)
    m_sOptFunctionName = sNewValue
End Property

Public Property Get OptPictureFolder() As String
    OptPictureFolder = m_sOptPictureFolder
End Property
Public Property Let OptPictureFolder(ByVal sNewValue As String)
    m_sOptPictureFolder = sNewValue
End Property

Public Property Get OptPictureAsc() As String
    OptPictureAsc = m_sOptPictureAsc
End Property
Public Property Let OptPictureAsc(sNewValue As String)
    m_sOptPictureAsc = sNewValue
End Property

Public Property Get OptPictureDesc() As String
    OptPictureDesc = m_sOptPictureDesc
End Property
Public Property Let OptPictureDesc(sNewValue As String)
    m_sOptPictureDesc = sNewValue
End Property

Public Property Get OptCmbPrefix() As String
    OptCmbPrefix = m_sOptCmbPrefix
End Property
Public Property Let OptCmbPrefix(sNewValue As String)
    m_sOptCmbPrefix = sNewValue
End Property

Public Property Get OptCmbSuffix() As String
    OptCmbSuffix = m_sOptCmbSuffix
End Property
Public Property Let OptCmbSuffix(sNewValue As String)
    m_sOptCmbSuffix = sNewValue
End Property

Public Property Get OptTextBoxSuffix() As String
    OptTextBoxSuffix = m_sOptTextBoxSuffix
End Property
Public Property Let OptTextBoxSuffix(sNewValue As String)
    m_sOptTextBoxSuffix = sNewValue
End Property

Public Property Get OptTextBoxPrefix() As String
    OptTextBoxPrefix = m_sOptTextBoxPrefix
End Property
Public Property Let OptTextBoxPrefix(sNewValue As String)
    m_sOptTextBoxPrefix = sNewValue
End Property
'//====================================== END PROP =======================================

'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PUBLIC SUB/FUNC   \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function OpenMsApp() As Boolean
    On Error GoTo ERR_OpenMsApp

    If Me.MsAppIsUp Then
        MsgBox "Un objet Access.Application à déjà était créer.", vbInformation + vbOKOnly, "OpenMsApp"
        Exit Function
    End If

    '// Création de l'objet Access.Application.
    Set m_oMsApp = New Access.Application
<<<<<<< HEAD
=======
'TODO:
    m_oMsApp.Visible = False
>>>>>>> 90b235b329fe03c39da9bc499914478247f26c39
    DoEvents
    m_oMsApp.Visible = False
    m_bMsAppIsUp = True
    OpenMsApp = True

SORTIE_OpenMsApp:
    Exit Function

ERR_OpenMsApp:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.CCreateFormContinu.OpenMsApp, ligne " & Erl & "."
    Resume SORTIE_OpenMsApp
End Function

' ----------------------------------------------------------------
' Procedure Nom:    OpenMsBase
' Sujet:            Ouvre la base.
' Procedure Kind:   Function
' Procedure Access: Public
' Références:       Ouvre la base.
'
'=== Paramètres ===
' sBaseFullName (String): Chemin et nom de la base à ouvrir.
'==================
'
' Return Type: Boolean TRUE si pas de problème.
'
' Author: Laurent
' Date:    30/04/2022 - 18:11
' DateMod: 07/05/2022 - 12:42
'
' !Use! : CheckFileExist, HasAutoexec, GetStartUpForm
' ----------------------------------------------------------------
Public Function OpenMsBase(sBaseFullName As String) As Boolean
    On Error GoTo ERR_OpenMsBase

    If Me.MsBaseIsOpen Then
        MsgBox "Un objet DAO.Database est déjà ouvert, veuillez fermer cette base svp.", vbInformation + vbOKOnly, "OpenMsBase"
        Exit Function
    End If

    Dim sObj    As String
    Dim bRep    As Boolean
    Dim bAutoEx As Boolean

    '// Voir si la base déjà ouverte (fichier .laccdb)...
    bRep = CheckFileExist(sBaseFullName, "laccdb")
    If bRep Then
        MsgBox "La base " & sBaseFullName & vbCrLf & "est déjà ouverte, opération annuler.", vbExclamation, "OpenMsBase"
        OpenMsBase = False
        Exit Function
    End If

    '// Ouverture de la base (sBaseName).
    m_oMsApp.OpenCurrentDatabase sBaseFullName, True
    DoEvents
    Set m_oMsBase = m_oMsApp.CurrentDb
    '// Vérifie si bien initialisée.
    If (m_oMsBase Is Nothing) Then
        MsgBox "Impossible d'ouvrir la base " & sBaseFullName, vbCritical, "OpenMsBase"
        OpenMsBase = False
        Exit Function
    End If

    '// Vérifier si macro Autoexec...
    bAutoEx = HasAutoexec(m_oMsBase)
    If bAutoEx Then
        MsgBox "La base utilise une macro 'Autoexec', désactivez la avant d'utiliser cet outil." & vbCrLf & _
               "Opération annulée.", vbExclamation, "OpenMsBase"
        Me.MsBaseIsOpen = True
        Exit Function
    End If

    '// Vérifier si un form ouvert au démarrage...
    sObj = GetStartUpForm(m_oMsBase)
    If (sObj <> vbNullString) Then

        MsgBox "Le Formulaire '" & sObj & "' est ouvert automatiquement au lancement de la base," & vbCrLf & _
               "il doit être fermer pour continuer." & vbCrLf & vbCrLf & _
               "Note: il peux contenir du code qui vas fermer base ou l'application.", vbExclamation, "OpenMsBase"
        '// On ferme l'objet.
        m_oMsApp.DoCmd.Close acForm, sObj
    End If

    '// Vérification si la base est toujours ouverte.
    If (m_oMsApp.CurrentProject.AccessConnection Is Nothing) Then
        MsgBox "La base à était fermée par l'objet " & sObj, vbExclamation, "OpenMsBase"
        m_bMsBaseIsOpen = True      '// Evite message 'base déjà fermée".
        Exit Function
    End If

    m_oMsApp.Visible = False
    Me.MsBaseIsOpen = True              '// Flag base ouverte.
    m_sBaseFullName = sBaseFullName     '// Stock le nom.
    ReDim m_sFormFields(0)              '// RaZ du tableau des champs.
    m_lFieldIndex = -1                  '// RaZ index tableau.

    OpenMsBase = True

SORTIE_OpenMsBase:
    Exit Function

ERR_OpenMsBase:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.CCreateFormContinu.OpenMsBase, ligne " & Erl & "."
    Resume SORTIE_OpenMsBase
End Function

Public Function CloseMsBase(Optional CloseMsApp As Boolean = False) As Boolean
    On Error GoTo ERR_CloseMsBase

'TODO:    m_oMsApp.Visible = True

    If (Me.MsAppIsUp = False) Or (Me.MsBaseIsOpen = False) Then
        MsgBox "La base est déjà fermée.", vbInformation + vbOKOnly, "OpenMsBase"
        Exit Function
    End If

    '// Ouverture de la base (sBaseName).
    If (Not m_oMsBase Is Nothing) Then
        If (Not m_oMsApp.CurrentProject.AccessConnection Is Nothing) Then m_oMsApp.CloseCurrentDatabase
    End If
    DoEvents

    m_bMsBaseIsOpen = False
    Set m_oMsBase = Nothing

    '// On ferme aussi l'application ?
    If CloseMsApp Then
        If (MsAppIsUp And (Not m_oMsApp Is Nothing)) Then m_oMsApp.Quit
        MsAppIsUp = False
        DoEvents
        Set m_oMsApp = Nothing
    End If

    CloseMsBase = True

SORTIE_CloseMsBase:
    Exit Function

ERR_CloseMsBase:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.CCreateFormContinu.CloseMsBase, ligne " & Erl & "."
    Resume SORTIE_CloseMsBase
End Function

Public Function LanceCreation() As Boolean
On Error GoTo ERR_LanceCreation

    Dim sDosTemp    As String
    Dim sFile       As String
    Dim sFileMod    As String
    Dim bRep        As Boolean  '// Retour Function.

    '// Extraire le code de la table dans des fichiers temp...
    sDosTemp = Environ("Temp") & "\"
    bRep = ExtrairePiecesJointes("~T_Info", "pjCode", sDosTemp)
    If (bRep = False) Then Exit Function

    '// Modification des valeurs dans le code...
    sFile = sDosTemp & CODE_CLASS
    sFileMod = RemplaceCode(sFile)
    If (sFileMod = vbNullString) Then Exit Function

    '// Création du module de Classe...
    '// ------------------------------
    bRep = CreateModuleClass(sFileMod)
    If (bRep = False) Then Exit Function

    '// Création du formulaire...
    '// ------------------------------
    LanceCreation = CreateForm
    
SORTIE_LanceCreation:
    Exit Function

ERR_LanceCreation:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  CSord.CCreateFormContinu.LanceCreation, ligne " & Erl & "."
    Resume SORTIE_LanceCreation
End Function
'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ END PUB. SUB/FUNC \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'// ######################## PRIVATE SUB/FUNC ###########################################
'
' NOTE: cmb 1 - Left 566 - Top 170 - Height 405 - Width 1125
'       twips 567 twips par centimètre / header 1700 twips

' ----------------------------------------------------------------
' Procedure Nom:    CreateForm
' Sujet:            Création du formulaire en continu pour la classe 'CsordFormColumn'
' Procedure Kind:   Function
' Procedure Access: Private
'
'=== Paramètres ===
'==================
'
' Return Type: Boolean True si pas d'erreur.
'
' Author:  Laurent
' Date:    03/05/2022 - 15:41
' DateMod: 23/04/2022 - 20:01
'
' !Use! : RemplaceCode
' ----------------------------------------------------------------
Private Function CreateForm() As Boolean
    On Error GoTo ERR_CreateForm

    Dim oObj        As Object
    Dim oForm       As Access.Form
    Dim oModule     As Access.module
    Dim oCtrCmb     As Access.CommandButton
    Dim oCtrTxtB    As Access.TextBox
    Dim oDb         As DAO.Database
    Dim oRst        As DAO.Recordset
    Dim oRstField   As DAO.Field

    Const EVE_CLICK     As String = "Click"     '// Event proc sur click button.

    Const CTR_WIDTH     As Long = 2835  '// Width  TxtB/Cmb
    Const CMB_HEIGHT    As Long = 405   '// Height  CommandButton
    Const TXB_HEIGHT    As Long = 315   '// Height  TextBox
    Const CMB_TOP       As Long = 1303  '// Top CommandButton (en bas de la section détail)

    Dim lLeft       As Long     '// Position gauche TxtB/Cmb 566

    Dim sFirstField As String   '// First field in rst (for .SetFocus)
    Dim lLigne      As Long     '// Ligne en cour dans le code.
    Dim sCode       As String   '// Code insérer dans le formulaire.
    Dim lIndex      As Long     '// Index pour le tableau des champs.

    Dim sDos        As String   '// Temp folder.
    Dim sFile       As String   '// source code.
    Dim sFileMod    As String   '// New code.

    Dim lRep    As Long     '// MsgBox.
    Dim sMsg    As String   '// MsgBox.

    '// Vérification si le formulaire existe déjà.
    For Each oObj In m_oMsApp.CurrentProject.AllForms

        If oObj.Name = Me.OptFormName Then
            sMsg = "Le formulaire '" & Me.OptFormName & "' Existe déjà dans le projet source." & vbNewLine & vbNewLine & "Voulez-vous le remplacer ?"
            Exit For
        End If
    Next

    '// si le formulaire existe, le supprimer ?
    If (sMsg <> vbNullString) Then
        lRep = MsgBox(sMsg, vbDefaultButton2 + vbYesNo, "Remplacer le formulaire")
        If (lRep = vbYes) Then
            m_oMsApp.DoCmd.DeleteObject acForm, Me.OptFormName
        Else
            Exit Function
        End If
    End If

'// ------------ Création Form, TextBox, CommandButton ----------
    
    Set oForm = m_oMsApp.CreateForm()

    '// Paramètrages.
    oForm.RecordSource = Me.OptFormSource
    oForm.DefaultView = 1                       '// Mode continu.
    oForm.Width = 17000
    oForm.ScrollBars = 2                        '// Verticale seulement.
    oForm.Caption = Me.OptFormName
    m_oMsApp.DoCmd.RunCommand acCmdFormHdrFtr   '// Création de la section header/footer.
    oForm.Section(acHeader).Height = 1708
    oForm.Section(acDetail).Height = 315

    sCode = oForm.Name                                      '// Stock temporairement le nom du form.
    m_oMsApp.DoCmd.Close acForm, sCode, acSaveYes
    m_oMsApp.DoCmd.Rename Me.OptFormName, acForm, sCode     '// Renome le form.
    m_oMsApp.DoCmd.OpenForm Me.OptFormName, acDesign, , , , acHidden

    Set oForm = m_oMsApp.Forms(Me.OptFormName)
    Set oModule = oForm.module

    '// Ouvre la source du form, pour définir le nom des commandButton et TextBox.
    Set oDb = m_oMsApp.CurrentDb
    Set oRst = oDb.OpenRecordset(Me.OptFormSource, dbOpenDynaset, dbReadOnly + dbForwardOnly)

    lLeft = 566 '// point de depart position left pour des boutons/Textebox.

    For Each oRstField In oRst.Fields
        '// Le champs fait partie de ceux sélectionnés ?
        If (oRstField.Name = m_sFormFields(lIndex)) Then
        
            '// Création des CommandButton (in Header).
            Set oCtrCmb = m_oMsApp.CreateControl(Me.OptFormName, acCommandButton, acHeader, , , lLeft, CMB_TOP, CTR_WIDTH, CMB_HEIGHT)
            oCtrCmb.Name = Me.OptCmbPrefix & oRstField.Name & Me.OptCmbSuffix
            oCtrCmb.Caption = oRstField.Name
            oCtrCmb.TabStop = False

            '// Détermine le control qui aura le focus, on prend le 1er champ trouver (Me." & sFirstField & ".SetFocus")
            If (lIndex = 0) Then sFirstField = Me.OptTextBoxPrefix & oRstField.Name & Me.OptTextBoxSuffix

            '// Création des TexteBox (in Detail).
            Set oCtrTxtB = m_oMsApp.CreateControl(oForm.Name, acTextBox, acDetail, , , lLeft, 0, CTR_WIDTH, TXB_HEIGHT)
            oCtrTxtB.Name = Me.OptTextBoxPrefix & oRstField.Name & Me.OptTextBoxSuffix
            oCtrTxtB.ControlSource = oRstField.Name

            '// Création des Procédure event pour les CommandButton.
            lLigne = oModule.CreateEventProc(EVE_CLICK, oCtrCmb.Name)
            oModule.InsertLines lLigne + 1, vbTab & Me.OptFunctionName
            oModule.DeleteLines lLigne + 2, 1       '// Supprime les lignes vide.

            lLeft = lLeft + CTR_WIDTH               '// Incrémente position gauche des txtb/cmb.
            
            '// Tous les champs sélectionnés ont était créer, on sort.
            lIndex = lIndex + 1
            If (lIndex > UBound(m_sFormFields)) Then Exit For

        End If
    Next
'// --------- FIN Création Form, TextBox, CommandButton ---------

    '// Insertion du code Function dans le form pour utilisation de la classe 'CsordFormColumn'...
    sDos = Environ("Temp") & "\"
    sFile = sDos & CODE_FRM
    sFileMod = RemplaceCode(sFile, sFirstField)
    If (sFileMod = vbNullString) Then Exit Function
    oModule.AddFromFile sFileMod

    oRst.Close
    oDb.Close

    '// Enregistrement du form.
    m_oMsApp.DoCmd.Close acForm, Me.OptFormName, acSaveYes

    CreateForm = True

SORTIE_CreateForm:
    On Error Resume Next
    Set oForm = Nothing
    Set oModule = Nothing
    Set oCtrCmb = Nothing
    Set oCtrTxtB = Nothing
    Set oRst = Nothing
    Set oRstField = Nothing
    Set oDb = Nothing

    Exit Function

ERR_CreateForm:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  CCreateFormContinu.MD_CreateForm.CreateForm, ligne " & Erl & "."
    Resume SORTIE_CreateForm
End Function

Public Function RemplaceCode(ByVal sFile As String, Optional sCtrFocus As String) As String
    On Error GoTo ERR_RemplaceCode

    Dim oFSO        As Object
    Dim oStreamIn   As Object
    Dim oStreamOut  As Object
    Dim sLigne      As String
    Dim sRemp       As String

    Set oFSO = m_cUtils.GetoFSO

    '// paramètres OpenTextFile :  ForReading 1, ForWriting  2
    '//                            TristateUseDefault -2 = valeur système par défaut.
    Set oStreamIn = oFSO.OpenTextFile(sFile, 1)             '// Lecture CodeForm.cls 'CodeForm.cls
    Set oStreamOut = oFSO.CreateTextFile(sFile & "tmp", 2)  '// Ecriture

    '// Boucle, modif texte, enregistre ligne.
    If InStr(sFile, CODE_FRM) Then
    
        '// Code Form : VARCLS FUNCNAME LENPREF LENSUFF CTRFOCUS
        Do While Not oStreamIn.AtEndOfStream
            sLigne = oStreamIn.ReadLine

            sRemp = Replace(sLigne, "VARCLS", Me.OptVarClasse): sLigne = sRemp
            sRemp = Replace(sLigne, "FUNCNAME", Me.OptFunctionName): sLigne = sRemp
            sRemp = Replace(sLigne, "LENPREF", Len(Me.OptTextBoxPrefix)): sLigne = sRemp
            sRemp = Replace(sLigne, "LENSUFF", Len(Me.OptTextBoxSuffix)): sLigne = sRemp
            sRemp = Replace(sLigne, "CTRFOCUS", sCtrFocus): sLigne = sRemp
            oStreamOut.WriteLine sLigne
        Loop
    End If

    If InStr(sFile, CODE_CLASS) Then

        '// Module de classe : PICDEFFOLDER PICDEFIMGASC PICDEFIMGDESC
        Do While Not oStreamIn.AtEndOfStream
            sLigne = oStreamIn.ReadLine
    
            sRemp = Replace(sLigne, "PICDEFFOLDER", Me.OptPictureFolder): sLigne = sRemp
            sRemp = Replace(sLigne, "PICDEFIMGASC", Me.OptPictureAsc): sLigne = sRemp
            sRemp = Replace(sLigne, "PICDEFIMGDESC", Me.OptPictureDesc): sLigne = sRemp
            oStreamOut.WriteLine sLigne
        Loop
    End If

    oStreamIn.Close
    oStreamOut.Close

    '// Remvoie le nom du fichier modifié.
    RemplaceCode = sFile & "tmp"
    
SORTIE_RemplaceCode:
    Set oStreamIn = Nothing
    Set oStreamOut = Nothing
    Set oFSO = Nothing
    Exit Function

ERR_RemplaceCode:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.CCreateFormContinu.RemplaceCode, ligne " & Erl & "."
    Resume SORTIE_RemplaceCode
End Function

Private Function CreateModuleClass(sFileForInser As String) As Boolean
    On Error GoTo ERR_CreateModuleClass

    Dim oMd     As VBIDE.VBComponent
    Dim sMsg    As String   '// MsgBox.
    Dim lRep    As Long     '// MsgBox.

    '// Le module existe déjà ? ...
    If (ModuleExiste(MOD_CLASS, m_oMsApp.VBE.ActiveVBProject)) Then
        sMsg = "Le module " & MOD_CLASS & " existe déjà dans le projet source." & vbNewLine & vbNewLine & "Voulez-vous le remplacer ?"
        lRep = MsgBox(sMsg, vbDefaultButton2 + vbYesNo, "Remplacer le module")
        If (lRep = vbNo) Then Exit Function
        '// Supprime le module.
        With m_oMsApp.VBE.ActiveVBProject.VBComponents
            .Remove .Item(MOD_CLASS)
        End With
    End If

    '// Création du module.
    Set oMd = m_oMsApp.VBE.ActiveVBProject.VBComponents.Add(vbext_ct_ClassModule)
    'm_oMsApp.VBE.MainWindow.Visible = True
    'm_oMsApp.VBE.ActiveVBProject.VBComponents(oMd.Name).Activate

    '// Renome.
    m_oMsApp.DoCmd.Close acModule, oMd.Name, acSaveYes
    m_oMsApp.DoCmd.Rename MOD_CLASS, acModule, oMd.Name

    'DoCmd.OpenModule MOD_CLASS

    '// Insert le fichier.
    oMd.CodeModule.AddFromFile sFileForInser

    '// Sauvegarde
    m_oMsApp.DoCmd.Close acModule, MOD_CLASS, acSaveYes
    DoEvents

    CreateModuleClass = True

SORTIE_CreateModuleClass:
    Exit Function

ERR_CreateModuleClass:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  CSord.CCreateFormContinu.CreateModuleClass, ligne " & Erl & "."
    Resume SORTIE_CreateModuleClass
End Function
'// ################################# END PRIV. SUB/FUNC #################################
