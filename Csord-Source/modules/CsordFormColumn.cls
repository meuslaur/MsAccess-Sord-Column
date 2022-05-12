VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsordFormColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name:     CSordFormColumn
' Kind:     Module de classe
' Purpose:  Tri sur les colonnes d'un form en continu.
' Author:   Laurent
' Date:     18/04/2022
' DateMod:  02/05/2022-10:04
'
'   Utilisation dans le Formulaire :
'   ==============================
'
'   Private CSordForm As CsordFormColumn
'
'   Private Sub Form_Close()
'       Set CSordForm = Nothing      '// Recommended, clean mem.
'   End Sub
'
'
'   Utilisation dans le Controle :
'   ============================
'
'Private Sub ID_Supplier_cmd_Click()
'   if sufix (len 4) is defined, the classe use 'ID_Supplier' for Field name.
'   SordColumn
'End Sub
'
'       OR
'
'Private Sub CompanyTest_NameTest_cmd_Click()
'   Pass Field name, beacause the classe can't use this name for field
'   SordColumn "Company_Name", OptionOn
'End Sub
'
'   Créate this function or sub in your form
'   ========================================
'
'Private Function SordColumn(Optional eActiveImage As T_OnOff = OptionOff, _
'                            Optional eActiveTexte As T_OnOff = OptionOn, _
'                            Optional sPicAsc As String = vbNullString, _
'                            Optional sPicDesc As String = vbNullString, _
'                            Optional sFieldName As String = vbNullString) As Boolean
'
'    Dim bRet As Boolean
'
'    '// Initialisation de la classe, on peut indiquer, si besoin, le préfixe et/ou le suffixe (nb de car).
'    '// Init class and defined suffix (the class cuts automatically the button name for extact field name)
'    If (m_CSordForm Is Nothing) Then
'        Set m_CSordForm = New CsordFormColumn
'        m_CSordForm.FieldPrefixLen = 0
'        m_CSordForm.FieldSuffixLen = 4
'    End If
'
'    With m_CSordForm
'        .PictureOn = eActiveImage   '// Optional (Default Off) si On, active les images par defaut.
'        .TexteOn = eActiveTexte     '// Optional (Default On)
'        '.TexteColor = 2366701       '// Optional (Default see const 'TXT_COULSORD' in class)
'        '.PicturePath = "\Res\"      '// Optional (Sous dossier des images, Default CurrentProject.Path + Const 'IMG_FOLDER' in class)
'        .PictureASC = sPicAsc       '// Optional (see const 'IMG_ASC' in class)
'        .PictureDESC = sPicDesc     '// Optional (see const 'IMG_DESC' in class)
'
'        bRet = .SordNow(sFieldName)     '// Execute le tri, retour TRUE if ok, 'sFieldName' is optional.
'
'    End With
'
'    Me.ID_Supplier.SetFocus
'
'    If (bRet = False) Then
'        '// Your code here
'        '// Your code here
'    End If
'
'End Function
' ------------------------------------------------------
Option Compare Database
Option Explicit

    '// Defaut paramètres.
    Private Const DEF_TXTFONTWEIGHT As Long = 700               '// Bold
    Private Const DEF_TXTCOLOR      As Long = 1137349           '// Couleur with sord  (m_objButton)
'VERIF:
    Private Const DEF_IMGFOLDER     As String = "\assets\"             '// Default folder for picture (currentProject.path + DEF_IMGFOLDER)
    Private Const DEF_IMGASC        As String = "ic_SordASC16.png"  '// (m_objButton)
    Private Const DEF_IMGDESC       As String = "ic_SordDESC16.png" '// (m_objButton)

    Private Const SORD_ASC          As String = " ASC"
    Private Const SORD_DESC         As String = " DESC"

    '// ON/OFF picture and/or caption texte.
    Public Enum T_OnOff
        OptionOff = False
        OptionOn = True
    End Enum
    Private m_eTexteOn          As T_OnOff      '// Indique si mettre color et bold sur le texte du bouton (on par defaut).
    Private m_ePictureOn        As T_OnOff      '// Indique si il faut afficher une image dans le bouton (off par defaut).

    Private Enum T_Sord
        eAsc = 1
        eDESC = 2
    End Enum
    Private m_eCurrentSord      As T_Sord           '// Sens du tri en cours (1 asc, 2 desc).

    Private m_bCheckOk          As Boolean          '// Flag if error occured.

    Private m_objForm           As Form             '// Form parent du control(sur lequel appliquer le tri).

    Private m_objButton         As CommandButton    '// Control bouton cliquer à utiliser.
    Private m_sButtonOldName    As String           '// Nom du control précedement modifier.

    Private m_sFieldName        As String           '// Champ sur lequel appliquer le tri.
    Private m_iFieldPrefixLen   As Integer          '// Préfix name for button.name
    Private m_iFieldSuffixLen   As Integer          '// Suffix name for button.name

    '// Paramètre pour m_objButton.
    Private m_lTxtColor         As Long             '// Indique la couleur à utiliser.
    Private m_sPicturePath      As String
    Private m_sPictureAsc       As String           '// Image tri pour ASC.
    Private m_sPictureDesc      As String           '// Image tri pour DESC.
    Private m_sPictureSord      As String           '// Image a afficher suivant le tri en cours (ASC ou DESC).

    '// Sauvegarde des propriétés du form/button avant modification.
    Private m_sSvFrmSord        As String           '// Sauvegarde les paramètres de tri d'origine.
    Private m_bSvFrmSortOn      As Boolean          '//             ""

    Private m_lSvTxtColor       As Long
    Private m_lSvTxtFontW       As Long

    Private m_eSvPicArr         As AcPictureCaptionArrangement
    Private m_SvPicType         As Byte             '// 0 incorporée, 1 liée
    Private m_vSvPicData        As Variant          '// Pour PicType = 0 (pictureData)
    Private m_sSvPicFile        As String           '// Pour PicType = 1 (Path and picture file)

'//----------------------------------       INIT/TERM       ----------------------------------
Private Sub Class_Initialize()

    '// TOUT est OK par défaut.
    m_bCheckOk = True
    m_sFieldName = "*"
'VERIF:
    m_sPicturePath = CurrentProject.Path & DEF_IMGFOLDER
    m_sPictureAsc = DEF_IMGASC
    m_sPictureDesc = DEF_IMGDESC
    
    m_lTxtColor = DEF_TXTCOLOR

End Sub

Private Sub class_Terminate()

    '// Restaure les filtres du form.
    If (Not m_objForm Is Nothing) Then
        m_objForm.OrderBy = m_sSvFrmSord
        m_objForm.OrderByOn = m_bSvFrmSortOn
    End If

    Set m_objButton = Nothing
    Set m_objForm = Nothing

End Sub
'//---------------------------------------------------------------------------------------

'//==================================       PROP        ==================================
Public Property Get FieldPrefixLen() As Integer
    FieldPrefixLen = m_iFieldPrefixLen
End Property
Public Property Let FieldPrefixLen(NewFieldPrefixLen As Integer)
    m_iFieldPrefixLen = 0
    If (NewFieldPrefixLen >= 0) Then m_iFieldPrefixLen = NewFieldPrefixLen
End Property

Public Property Get FieldSuffixLen() As Integer
    FieldSuffixLen = m_iFieldSuffixLen
End Property
Public Property Let FieldSuffixLen(NewFieldSuffixLen As Integer)
    m_iFieldSuffixLen = 0
    If (NewFieldSuffixLen >= 0) Then m_iFieldSuffixLen = NewFieldSuffixLen
End Property

'// Texte On
Public Property Get TexteOn() As T_OnOff
    TexteOn = m_eTexteOn
End Property
Public Property Let TexteOn(ActiveTexte As T_OnOff)
    m_eTexteOn = IIf(m_bCheckOk, ActiveTexte, OptionOff)
End Property

'// Texte Color
Public Property Get TexteColor() As Long
    TexteColor = m_lTxtColor
End Property
Public Property Let TexteColor(NewColor As Long)
    m_lTxtColor = IIf((NewColor < 1), DEF_TXTCOLOR, NewColor)
    Me.TexteOn = OptionOn
End Property

'// Picture On
Public Property Get PictureOn() As T_OnOff
    PictureOn = m_ePictureOn
End Property
Public Property Let PictureOn(ActivePicture As T_OnOff)
    '// Si pas d'erreur sur Dossier ou image, mettre à jour.
    m_ePictureOn = IIf(m_bCheckOk, ActivePicture, OptionOff)
End Property

'// Picture Path
Public Property Get PicturePath() As String
    PicturePath = m_sPicturePath
End Property
Public Property Let PicturePath(NewPath As String)

    '// Evite les messages redondants.
    If (m_bCheckOk = False) Then Exit Property

    '// Le dossier existe ?
    If (CheckFolderExists(NewPath) = False) Then
        MsgBox "Sous-Dossier '" & NewPath & "', non trouver", vbExclamation + vbOKOnly, "PicturePath"
        m_bCheckOk = False
        PictureOn = OptionOff
        Exit Property
    End If

    '// Vérifier si fini bien par '\'.
    If (Right$(NewPath, 1) <> "\") Then NewPath = NewPath & "\"

    m_sPicturePath = NewPath

End Property

'// Picture for sord asc
Public Property Get PictureASC() As String
    PictureASC = m_sPictureAsc
End Property
Public Property Let PictureASC(NewFileName As String)

    '// OFF, rien a faire on sort.
    '// Evite les messages répétitif d'érreur.
    If (m_ePictureOn = OptionOff) Then Exit Property

    '// Img vide, et img par défaut déjà appliquer, on sort.
    If (NewFileName = vbNullString) And (m_sPictureAsc = DEF_IMGASC) Then Exit Property

    '// Vide applique img par défaut.
    If (NewFileName = vbNullString) Then NewFileName = DEF_IMGASC

    '// Le fichier existe ?
    If (CheckFileExist(m_sPicturePath & NewFileName) = False) Then
        MsgBox "Fichier " & vbCrLf & NewFileName & vbCrLf & " non trouver.", vbExclamation, "PictureASC"
        m_bCheckOk = False
        m_ePictureOn = OptionOff
        Exit Property
    End If

    m_sPictureAsc = NewFileName
    m_ePictureOn = OptionOn

End Property

'// Picture for sord desc
Public Property Get PictureDESC() As String
    PictureDESC = m_sPictureDesc
End Property
Public Property Let PictureDESC(NewFileName As String)

    '// OFF, rien a faire on sort.
    '// Evite les messages répétitif d'érreur.
    If (m_ePictureOn = OptionOff) Then Exit Property

    '// Img vide, et img par défaut déjà appliquer, on sort.
    If (NewFileName = vbNullString) And (m_sPictureDesc = DEF_IMGDESC) Then Exit Property

    '// Vide applique img par défaut.
    If (NewFileName = vbNullString) Then NewFileName = DEF_IMGDESC

    '// Le fichier existe ?
    If (CheckFileExist(m_sPicturePath & NewFileName) = False) Then
        MsgBox "Fichier " & vbCrLf & NewFileName & vbCrLf & " non trouver.", vbExclamation, "PictureDESC"
        m_bCheckOk = False
        m_ePictureOn = OptionOff
        Exit Property
    End If

    m_sPictureDesc = NewFileName
    m_ePictureOn = OptionOn

End Property

Private Property Get GetFSO() As Object

    Static pFSO               As Object
    
    If pFSO Is Nothing Then
        Set pFSO = CreateObject("Scripting.FileSystemObject")
    End If

    Set GetFSO = pFSO

End Property

'//=======================================================================================

'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PUBLIC SUB/FUNC   \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Function SordNow(Optional sFieldName As String = vbNullString) As Boolean
    On Error GoTo ERR_SordNow

    DoCmd.Echo False    'TODO: DoCmd.Echo

    '// Initialise le form, le bouton, si besoin...
    InitObjets
    If (m_bCheckOk = False) Then GoTo SORTIE_SordNow     '// Erreur on sort.

    '// Défini le champs, si changement de CommandButton.
    If (m_sButtonOldName <> m_objButton.Name) Then
        SetFieldName sFieldName
        If (m_bCheckOk = False) Then GoTo SORTIE_SordNow      '// Erreur on sort.
    End If

    '// On inverse le tri.
    Dim sCurSord As String

    '// Determine le texte et l'image as utilisés suivant le sens du tri.
    Select Case m_eCurrentSord
        Case eAsc
            m_eCurrentSord = eDESC
            sCurSord = SORD_ASC
            m_sPictureSord = m_sPicturePath & m_sPictureAsc
        Case eDESC
            m_eCurrentSord = eAsc
            sCurSord = SORD_DESC
            m_sPictureSord = m_sPicturePath & m_sPictureDesc
    End Select

    '// Initialise le control cliquer...
    SavePropButton      '// Sauvegarde le bouton en cours...
    UpdatePropButton    '// MàJ du bouton en cours...

    '// Applique le tri.
    m_objForm.OrderBy = m_sFieldName & sCurSord
    m_objForm.OrderByOn = True

SORTIE_SordNow:
    SordNow = m_bCheckOk
    DoCmd.Echo True     'TODO:DoCmd.Echo
    Exit Function

ERR_SordNow:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.CSordFormColumn.SordNow, ligne " & Erl & "."
    m_bCheckOk = False
    Resume SORTIE_SordNow
End Function
'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'// ######################## PRIVATE SUB/FUNC ###########################################
' ----------------------------------------------------------------
' Procedure Nom:    SetFieldName
' Sujet:            Initialisation du champ de tri.
'                   Si null on utilise le nom du bouton, on vérifier si pref/suff.
'                   Si même nom que précedement, on fait rien.
'                   On vérifie si le nom est valide dans la table.
'                   On remet les options aux valeurs par défaut.
' Procedure Kind:   Sub
' Procedure Access: Private
'
'=== Paramètres ===
' NewFieldName (String): Nom du champs à utiliser.
'==================
'
' Author:  Laurent
' Date:    01/05/2022 - 14:42
' DateMod:
'
' !Use! : CheckPrefixSuffix, CheckFieldName, RestorePropButton.
' ----------------------------------------------------------------
Private Sub SetFieldName(NewFieldName As String)

    '// Meme champ, rien a faire on sort.
    If (m_sFieldName = NewFieldName) Then Exit Sub

    Dim sTmpField As String
    sTmpField = NewFieldName

    '// ...Pas de nom de champ passer, on utilise le nom du commandButton...
    If (NewFieldName = vbNullString) Then sTmpField = CheckPrefixSuffix(m_objButton.Name)

    '// ...Vérifier la validité du champ...
    CheckFieldName sTmpField
    If (m_bCheckOk = False) Then Exit Sub

    '// ...Nouveau champ, Restaure le commandButton précédent...
    RestorePropButton

    '// Place les options par défaut, remet le tri a DESC.
    m_eCurrentSord = eAsc

    '// On stock le nom du champ.
    m_sFieldName = sTmpField

End Sub

Private Sub InitObjets()

    '// ...Initialisation du form, si pas déjà fait...
    CheckForm
    If (m_bCheckOk = False) Then Exit Sub

    '// ...Initialise le bouton, si pas deja fait...
    CheckButton Screen.ActiveControl
    If (m_bCheckOk = False) Then Exit Sub

End Sub

Private Sub SavePropButton()

    '// Rien a faire, meme bouton, on sort.
    If (m_sButtonOldName = m_objButton.Name) Then Exit Sub

    '// Sauvegarde des propriétés, si nouveau bouton...
    m_lSvTxtFontW = m_objButton.FontWeight
    m_lSvTxtColor = m_objButton.ForeColor

    m_SvPicType = m_objButton.PictureType
    m_eSvPicArr = m_objButton.PictureCaptionArrangement
    Select Case m_SvPicType
        Case 0
            m_vSvPicData = m_objButton.PictureData
        Case Else
            m_sSvPicFile = m_objButton.Picture
    End Select

    m_sButtonOldName = m_objButton.Name

End Sub

Private Sub RestorePropButton()
    
    '// Rien a faire, 1er bouton ou même bouton, on sort.
    If ((m_sButtonOldName = vbNullString) Or (m_sButtonOldName = m_objButton.Name)) Then Exit Sub

    '// Texte.
    m_objForm(m_sButtonOldName).ForeColor = m_lSvTxtColor
    m_objForm(m_sButtonOldName).FontWeight = m_lSvTxtFontW

    '// Image.
    m_objForm(m_sButtonOldName).PictureType = m_SvPicType
    m_objForm(m_sButtonOldName).PictureCaptionArrangement = m_eSvPicArr

    Select Case m_SvPicType
        Case 0
            m_objForm(m_sButtonOldName).PictureData = m_vSvPicData
        Case Else
            m_objForm(m_sButtonOldName).Picture = m_sSvPicFile
    End Select

End Sub

Private Sub UpdatePropButton()

    '// Rien a faire, on sort.
    If (m_eTexteOn = OptionOff And m_ePictureOn = OptionOff) Then Exit Sub

    '// On Modifie le control.
    If m_eTexteOn Then
        m_objButton.FontWeight = DEF_TXTFONTWEIGHT
        m_objButton.ForeColor = m_lTxtColor
    End If

    If m_ePictureOn Then
        m_objButton.PictureCaptionArrangement = acLeft
        m_objButton.PictureType = 1
        m_objButton.Picture = m_sPictureSord
    End If
    
End Sub

' ----------------------------------------------------------------
' Procedure Nom:    CheckForm
' Sujet:            Vérifier si le form as bien était initialiser
' Procedure Kind:   Sub
' Procedure Access: Private
' Return Type:      Boolean TRUE si pas de problème.
' Author:           Laurent
' Date:             20/04/2022 - 06:37
' DateModif:        22/04/2022 - 09:04
' ----------------------------------------------------------------
Private Sub CheckForm()
    On Error GoTo ERR_CheckForm

    '// Initialise le form sur lequel appliquer le tri,
    '// Sauvegarde des filtres en cours.
    If (m_objForm Is Nothing) Then
        Set m_objForm = Screen.ActiveForm
        m_sSvFrmSord = m_objForm.OrderBy
        m_bSvFrmSortOn = m_objForm.OrderByOn
    End If

    m_bCheckOk = True

SORTIE_CheckForm:
    Exit Sub

ERR_CheckForm:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            "(" & Err.Description & ")" & vbCrLf & _
            "Dans  CSordFormColumn.CheckForm," & vbCrLf & _
            Switch(Erl = 0, vbNullString, Erl <> 0, vbCrLf & "Ligne No: " & Erl)
    m_bCheckOk = False
    Resume SORTIE_CheckForm
End Sub

' ----------------------------------------------------------------
' Procedure Nom:    CheckButton
' Sujet:            Vérifier si bouton défini suivant les options choisi.
' Procedure Kind:   Sub
' Procedure Access: Private
' Parameter NewButton (CommandButton):
' Return Type:      Boolean TRUE si pas de problème.
' Author:           Laurent
' Date:             20/04/2022 - 06:45
' DateModif:        22/04/2022 - 09:06
' ----------------------------------------------------------------
Private Sub CheckButton(NewButton As CommandButton)
    On Error GoTo ERR_CheckButton

    '// Initialise le bouton.
    If (m_objButton Is Nothing) Then
        Set m_objButton = NewButton
    ElseIf (m_objButton.Name <> NewButton.Name) Then
        Set m_objButton = NewButton
    End If

    m_bCheckOk = True

SORTIE_CheckButton:
    Exit Sub

ERR_CheckButton:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  SordFormCc.CSordFormColumn.CheckButton, ligne " & Erl & "."
    m_bCheckOk = False
    Resume SORTIE_CheckButton
End Sub

' ----------------------------------------------------------------
' Procedure Nom:    CheckFieldName
' Sujet:            Vérifier si la table contient des enrg et si nom du champ valide...
'                   Si tout vas bien m_bCheckOk est mis a True.
' Procedure Kind:   Sub
' Procedure Access: Private
'
'=== Paramètres ===
' sField (String): Nom du champ à vérifier.
'==================
'
' Author:  Laurent
' Date:    19/04/2022 - 11:35
' DateMod: 01/05/2022 - 13:4
'
' ----------------------------------------------------------------
Private Sub CheckFieldName(ByRef sField As String)
    On Error GoTo ERR_CheckFieldName

    m_bCheckOk = IIf((DCount(sField, m_objForm.RecordSource) = 0), False, True)

SORTIE_CheckFieldName:
    Exit Sub

ERR_CheckFieldName:
    If ((Err.Number = 3265) Or (Err.Number = 2471)) Then
        MsgBox "Le champ '" & sField & "' n'existe pas dans la source de donnée, vérifier la valeur de .FieldName", vbExclamation, "Classe SordFormColumn"
    Else
        MsgBox "Erreur " & Err.Number & " (" & Err.Description & ") Dans CheckFieldName." & Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Ligne No: " & Erl)
    End If

    m_bCheckOk = False
    Resume SORTIE_CheckFieldName
End Sub
' ----------------------------------------------------------------
' Procedure Nom:    CheckPrefixSuffix
' Sujet:            Vérifier si retirer un Préfixe/Suffixe sur le nom du commandButton pour extraire le nom du champs à utiliser.
' Procedure Kind:   Function
' Procedure Access: Private
'
'=== Paramètres ===
' sCmbName (String): Nom du commandButton
'==================
'
' Return Type: String Retourne le nom du champs à utiliser.
'
' Author:   Laurent
' Date:     22/04/2022 - 10:51
' DateMod:  01/05/2022 - 12:54
' ----------------------------------------------------------------
Private Function CheckPrefixSuffix(sCmbName As String) As String

    Dim sResult As String
    Dim lTemp   As Long

    lTemp = Len(sCmbName)

    '// Len valide ?
    If ((Len(m_iFieldSuffixLen) + (Len(m_iFieldPrefixLen)) >= lTemp)) Then
        m_iFieldSuffixLen = 0: m_iFieldPrefixLen = 0
    End If

    '// Applique préfixe.
    If (m_iFieldPrefixLen > 0) Then
        sResult = Right$(sCmbName, lTemp - m_iFieldPrefixLen)
        sCmbName = sResult
    End If

    '// Applique suffixe.
    If (m_iFieldSuffixLen > 0) Then
        lTemp = Len(sCmbName)
        sResult = Left$(sCmbName, lTemp - m_iFieldSuffixLen)
    End If

    CheckPrefixSuffix = sResult

End Function

Private Function CheckFileExist(ByVal sFullPathFile As String) As Boolean

    Dim oFSO            As Object

    Set oFSO = GetFSO

    CheckFileExist = oFSO.FileExists(sFullPathFile)

    Set oFSO = Nothing

End Function


'Returns True if the folder exists (and is accessible)
' - trailing backslash is completely optional
' - returns False if the full path to an existing file is passed
'   to the function (and not just the folder part)
Private Function CheckFolderExists(ByVal PathToFolder As String) As Boolean

    Dim oFSO As Object
    Dim bRes As Boolean

    Set oFSO = GetFSO
    bRes = oFSO.FolderExists(PathToFolder)

    Set oFSO = Nothing
    CheckFolderExists = bRes
    Set oFSO = Nothing

End Function

'// #####################################################################################
