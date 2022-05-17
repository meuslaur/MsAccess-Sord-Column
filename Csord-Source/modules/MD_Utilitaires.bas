Attribute VB_Name = "MD_Utilitaires"
' ------------------------------------------------------
' Name:    MD_Utilitaires
' Kind:    Module
' Purpose: Utilitaires divers
' Author:  Laurent
' Date:    30/04/2022 - 13:58
' DateMod:
' Use:     Class C_Utilitaires
' ------------------------------------------------------
Option Compare Database
Option Explicit

'//::::::::::::::::::::::::::::::::::    VARIABLES      ::::::::::::::::::::::::::::::::::

'// Project Types
Public Enum T_ObjectType
    Tables_Local = 1
    Tables_Linked_ODBC = 4
    Tables_Linked = 6
    QueriesType = 5
    FormsType = -32768
    ReportsType = -32764
    MacrosType = -32766
    ModulesType = -32761
End Enum

'// FileDialog type
Public Enum T_FileDialogType
    FD_TypeFilePicker = 3
    FD_TypeFolderPicker = 4
    FD_TypeFileOpen = 1
    FD_TypeFileSaveAs = 2
End Enum
Public Enum T_FileDialogView
    FD_ViewDetails = 2
    FD_ViewLargeIcons = 6
    FD_ViewList = 1
    FD_ViewPreview = 4
    FD_ViewProperties = 3
    FD_ViewSmallIcons = 7
    FD_ViewThumbnail = 5
    FD_ViewTiles = 9
    FD_ViewWebView = 8
End Enum

Private m_cUtil As New CUtilitaires
'//:::::::::::::::::::::::::::::::::: END VARIABLES ::::::::::::::::::::::::::::::::::::::


'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PUBLIC SUB/FUNC   \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' ----------------------------------------------------------------
' Procedure Nom:    ListObjects
' Sujet:            Retourne sous forme de chaîne SQL ou de liste de valeurs,
'                   les objets d'une base suivant le/les type indiquer(T_ObjectType).
' Procedure Kind:   Function
' Procedure Access: Public
'
'=== Paramètres ===
' eObjectType1 (T_ObjectType):  Filtre 1 sur type d'objet (voir Enum T_ObjectType).
' eObjectType2 (T_ObjectType):  Fitre 2 sur type d'objet (voir Enum T_ObjectType.
' eObjectType3 (T_ObjectType):  Fitre 3 sur type d'objet (voir Enum T_ObjectType.
' bListeVal (Boolean):          Si a TRUE retourne une liste de valeur (pour une ListBox, ComboBox ou autre),
'                                   sinon retourne la chaine SQL pour une source de données.
' oAutreBd (Database):
'==================
'
' Return Type:  String Liste de valeur(bListeVal =True) ou chaine SQL(bListeVal =False).
' Author:       Laurent
' Date:         27/04/2022 - 10:36
' DateMod:      28/04/2022 - 16:57
'
' !Use! : Enum T_ObjectType
' !Use! : Function ObjectTypeName
' ----------------------------------------------------------------
Public Function ListObjects(eObjectType1 As T_ObjectType, _
                            Optional bListeVal As Boolean = False, _
                            Optional eObjectType2 As T_ObjectType, _
                            Optional eObjectType3 As T_ObjectType, _
                            Optional ByRef oAutreBd As DAO.Database) As String
On Error GoTo ERR_ListObjects

    Dim sSql    As String
 
    '// Création de la chaine SQL.
    sSql = "SELECT MsysObjects.Type, MsysObjects.Name AS ObjectName FROM MsysObjects " & _
           "WHERE (((MsysObjects.Flags)>=0) AND ((MsysObjects.Type)=" & eObjectType1
           
    If (eObjectType2) Then sSql = sSql & " Or (MsysObjects.Type)=" & eObjectType2
    If (eObjectType3) Then sSql = sSql & " Or (MsysObjects.Type)=" & eObjectType3
           
    sSql = sSql & ") AND ((MsysObjects.Name) Not Like '~*' And (MsysObjects.Name) Not Like 'MSys*'))" & _
                  "ORDER BY MsysObjects.Type, MsysObjects.Name;"

    If (bListeVal = False) Then
        '// Retourne la chaine SQL, et on sort.
        ListObjects = sSql
        Exit Function
    End If

    Dim oDb     As DAO.Database
    Dim oRst    As DAO.Recordset

    Set oDb = IIf(oAutreBd Is Nothing, CurrentDb, oAutreBd)
    Set oRst = oDb.OpenRecordset(sSql, dbOpenSnapshot)
    sSql = vbNullString

    '// Boucle sur les objets de la table système.
    '// Création de la liste de valeur sur 2 colonnes (Col0: Type, Col1: Name, col2: Type en clair).
    With oRst
        If .RecordCount <> 0 Then
            Do While Not .EOF
                sSql = sSql & ![Type] & ";" & ![ObjectName] & ";" & ObjectTypeName(![Type]) & ";"
                .MoveNext
            Loop
        End If
    End With
 
    '// Retourne la liste de valeurs.
    ListObjects = sSql

SORTIE_ListObjects:
    On Error Resume Next
    If Not oRst Is Nothing Then
        oRst.Close
        Set oRst = Nothing
    End If
    If Not oDb Is Nothing Then Set oDb = Nothing
    Exit Function
 
ERR_ListObjects:
    MsgBox "L’erreur suivante s’est produite" & vbCrLf & vbCrLf & _
           "Erreur N°: " & Err.Number & vbCrLf & _
           "Source : ListObjects" & vbCrLf & _
           "Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl), _
           vbOKOnly + vbCritical, "Erreur survenue !"
    Resume SORTIE_ListObjects
End Function

' ----------------------------------------------------------------
' Procedure Nom:    ObjectFieldsToListVal
' Sujet:            Retourne les champs d'une table ou d'une requête sous forme de liste de valeurs.
' Procedure Kind:   Function
' Procedure Access: Public
' Références:       N/A
'
'=== Paramètres ===
' sObjectName (String): Nom de la table ou de la requête.
' lType (T_ObjectType): Indique le type de l'objet (Enum T_ObjectType).
' oAutreBd (Database):  Database à utiliser.
'==================
'
' Return Type:  String
' Author:       Laurent
' Date:         28/04/2022 - 16:49
' ----------------------------------------------------------------
Public Function ObjectFieldsToListVal(sObjectName As String, lType As T_ObjectType, Optional ByRef oAutreBd As DAO.Database) As String
    On Error GoTo ERR_ObjectFieldsToListVal

    Dim oBd     As DAO.Database
    Dim oTbDef  As DAO.TableDef
    Dim oQrDef  As DAO.QueryDef
    Dim oField  As Field
    Dim sLstVal As String
 
    Set oBd = IIf(oAutreBd Is Nothing, CurrentDb, oAutreBd)

    Select Case lType
        Case Tables_Local, Tables_Linked_ODBC, Tables_Linked
            '// Boucle sur les champs de la table.
            Set oTbDef = oBd.TableDefs(sObjectName)
            For Each oField In oTbDef.Fields
                sLstVal = sLstVal & oField.Name & ";"
            Next
        Case QueriesType
            '// Boucle sur les champs de la requête.
            Set oQrDef = oBd.QueryDefs(sObjectName)
            For Each oField In oQrDef.Fields
                sLstVal = sLstVal & oField.Name & ";"
            Next
        Case FormsType
        Case ReportsType
        Case MacrosType
        Case ModulesType
    End Select

    '// Retourne la liste de valeurs.
    ObjectFieldsToListVal = sLstVal

SORTIE_ObjectFieldsToListVal:
    Set oField = Nothing
    Set oTbDef = Nothing
    Set oQrDef = Nothing
    Set oBd = Nothing
    Exit Function

ERR_ObjectFieldsToListVal:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TirSurFormContinu.MD_VerifChampOblig.TableFieldsToListVal, ligne " & Erl & "."
    Resume SORTIE_ObjectFieldsToListVal
End Function

'Returns True if the folder exists (and is accessible)
' - trailing backslash is completely optional
' - returns False if the full path to an existing file is passed
'   to the function (and not just the folder part)
Public Function CheckFolderExists(ByVal PathToFolder As String) As Boolean

    Dim oFSO As Object
    Dim bRes As Boolean

    Set oFSO = m_cUtil.GetoFSO
    bRes = oFSO.FolderExists(PathToFolder)

    Set oFSO = Nothing
    CheckFolderExists = bRes
    Set oFSO = Nothing

End Function

Public Function CreateNewFolder(ByVal sPathToFolder As String) As Boolean
    On Error GoTo ERR_CreateNewFolder

    Dim oFSO As Object
    Dim bRes As Boolean

    Set oFSO = m_cUtil.GetoFSO
    bRes = oFSO.FolderExists(sPathToFolder)

    If (bRes = False) Then
        oFSO.CreateFolder (sPathToFolder)
        bRes = True
    End If

    Set oFSO = Nothing
    CreateNewFolder = bRes
    Set oFSO = Nothing
    
SORTIE_CreateNewFolder:
    Exit Function

ERR_CreateNewFolder:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.MD_Utilitaires.CreateNewFolder, ligne " & Erl & "."
    Resume SORTIE_CreateNewFolder
End Function

'Convenience function to avoid creating a File System Object
' This should be used in place of the Len(Dir()) construct because
' Dir() has terrible performance compared to FileExists, especially in
' certain use cases (e.g., checking for the existence of a single file
' in a very large (300,000+ files) UNC directory, such as G:\Photos\)
'
'Includes support for wild-card characters ("*" and "?")
'--== Project-wide find & replace ==--
'We can use the MZ-Tools Find & Replace RegEx mode to do a program wide change:
'
' Find: Len\(Dir\(([^,]+)\)\) > 0
' Find: Dir\(([^,]+)\) <> ""
' Repl: FileExists($1)
'
' Find: Len\(Dir\(([^,]+)\)\) = 0
' Find: Dir\(([^,]+)\) = ""
' Repl: Not FileExists($1)
'
' Note: Some existing code may have defined FileExists() properties or functions
' that will overlap and cause problems; to work around this, we can simply
' add "FileFunctions." to fully qualify the function call:
' FileFunctions.FileExists(PathToMyFile)
' ----------------------------------------------------------------
' Procedure Nom:    CheckFileExist
' Sujet:            Vérifier si le fichier existe
' Procedure Kind:   Function
' Procedure Access: Private
'
'=== Paramètres ===
' sFullPathFile (String):   Chemin complet et nom du fichier.
' sExtFile (String):        Extension a utiliser.
' ProcedureName (String):   Nom de la procédure appelante.
'==================
' Return Type:  Boolean, TRUE si le fichier existe.
' Author:       ?
' Date:         20/04/2022 - 06:21
' DateMod:      04/05/2022 - 17:5
'
' !Use! :   m_cUtil
' ----------------------------------------------------------------
Public Function CheckFileExist(ByVal sFullPathFile As String, Optional ByVal sExtFile As String) As Boolean

    Dim oFSO            As Object
    Dim sFolder         As String
    Dim sFile           As String
    Dim sBase           As String
    Dim bRes            As Boolean

    Set oFSO = m_cUtil.GetoFSO

    '// Utilise l'extension de fichier indiquer.
    If (sExtFile <> vbNullString) Then
        sFolder = oFSO.GetParentFolderName(sFullPathFile) & "\"
        sFile = oFSO.GetFileName(sFullPathFile)
        sBase = oFSO.GetBaseName(sFile)
    
        '// Ajoute le '.' si besoin
        If (Left(sExtFile, 1) <> ".") Then sExtFile = "." & sExtFile
        
        sFullPathFile = sFolder & sBase & sExtFile
    End If

    bRes = oFSO.FileExists(sFullPathFile)

    CheckFileExist = bRes
    Set oFSO = Nothing

End Function

' ----------------------------------------------------------------
' Procedure Nom:    OuvreBoite
' Sujet:            Ouvre la boite de dialogue fichiers.
' Procedure Kind:   Function
' Procedure Access: Public
' Références:       Microsoft Office 16.0 Object Library
'
'=== Paramètres ===
' sFltDes (String):                 Désignation du filtre (ex: "Fichiers MS Access").
' sFltExt (String):                 Extension a filtrer (ex : "*.accdb;*.txt").
' sTitre (String):                  Titre de la boite.
' sInitialPath (String):            Dossier de départ (defaut oldforlder use or currentapp path).
' lDialogType (MsoFileDialogType):  Type de boite (defaut Files select).
' bReturnFullPath (Boolean):        Retourne ou non le chemin complet (defaut return fullpath/file).
'==================
'
' Return Type:  String
' Author:       Laurent
' Date:         28/04/2022 - 10:51
' ----------------------------------------------------------------
Public Function OuvreBoite(Optional sFltDes As String = "Tous fichiers", _
                           Optional sFltExt As String = "*.*", _
                           Optional sTitre As String, _
                           Optional sInitialPath As String, _
                           Optional eDialogType As T_FileDialogType = FD_TypeFilePicker) As String
On Error GoTo ERR_OuvreBoite

    Dim oFd             As Object
    Dim vSelectedItem   As Variant
    Dim sTmp            As String
    Dim sValRet         As String
    Dim lTmp            As Long

    Set oFd = Application.FileDialog(eDialogType)

    '// Defini le sous-dossier de départ, se place sur le dossier de l'app, ou sur la valeur indiquer.
    If (sInitialPath = vbNullString) Then
        lTmp = Len(CurrentProject.Path)
        sTmp = Left$(oFd.InitialFileName, lTmp)
        If (sTmp <> CurrentProject.Path) Then oFd.InitialFileName = sTmp
    Else
        oFd.InitialFileName = sInitialPath
    End If
    
    If (sTitre = vbNullString) Then sTitre = "Sélectionnez un dossier /  fichier"

    With oFd

        .Title = sTitre
        .AllowMultiSelect = False
        .InitialView = FD_ViewDetails

        '// Applique le filtre si pas en mode boite dossier.
        If (eDialogType <> FD_TypeFolderPicker) Then
            .Filters.Clear
            .Filters.Add sFltDes, sFltExt, 1
        End If

        '// Ouvre la boite, récupère la sélection.
        If .Show = True Then
            For Each vSelectedItem In .SelectedItems
                sValRet = vSelectedItem
            Next vSelectedItem

            OuvreBoite = sValRet

        End If
    End With

SORTIE_OuvreBoite:
    Set oFd = Nothing
Exit Function

ERR_OuvreBoite:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Gestionnaire d'erreur"
    Resume SORTIE_OuvreBoite
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetBackupFileName
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Return an unused filename for the database backup befor build
'// NOTE Retourne les infos de sauvegarde : folder;BaseBackup;Base use Split for extract.
'---------------------------------------------------------------------------------------
Public Function GetBackupFileName(sFullPath As String) As String

    Const SUFFIX    As String = "_BackUp("
    Dim oFSO        As Object

    Dim sFolder     As String
    Dim sFile       As String
    Dim sBase       As String
    Dim sBaseBackUp As String
    Dim sExt        As String
    Dim iFor        As Integer
    Dim sTest       As String
    Dim sIncrement  As String

    Set oFSO = m_cUtil.GetoFSO

    sFolder = oFSO.GetParentFolderName(sFullPath) & "\"
    sFile = oFSO.GetFileName(sFullPath)
    sBase = oFSO.GetBaseName(sFile)
    sExt = "." & oFSO.GetExtensionName(sFile)
    sIncrement = "00"

    ' Attempt up to 100 versions of the file name. (i.e. Database_VSBackup45.accdb)
    For iFor = 1 To 50
        sBaseBackUp = sBase & SUFFIX & sIncrement & ")" & sExt
        sTest = sFolder & sBaseBackUp
        If oFSO.FileExists(sTest) Then
            ' Try next number.
            sIncrement = CStr(iFor)
            If (Len(sIncrement) < 2) Then sIncrement = "0" & sIncrement
        Else
            Exit For
        End If
    Next iFor

    ' Return file name
    GetBackupFileName = sFolder & ";" & sBaseBackUp & ";" & sFile
    Set oFSO = Nothing

End Function

'---------------------------------------------------------------------------------------
' Procedure : CopyFile
' Author : Daniel Pineault, CARDA Consultants Inc.
' Website : http://www.cardaconsultants.com
' Purpose : Copy a file
' Overwrites existing copy without prompting
' Cannot copy locked files (currently in use)
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
' (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: None required ' ' Input Variables:
' ~~~~~~~~~~~~~~~~
' sSource - Path/Name of the file to be copied
' sDest - Path/Name for copying the file to
'
' Revision History:
' Rev Date(yyyy/mm/dd) Description
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 1 2007-Apr-01 Initial Release
'---------------------------------------------------------------------------------------
Public Function CopyFile(sSource As String, sDest As String) As Boolean
    On Error GoTo CopyFile_Error

    FileCopy sSource, sDest
    CopyFile = True
    Exit Function

CopyFile_Error:
    If Err.Number = 0 Then
    ElseIf Err.Number = 70 Then
        MsgBox "The file is currently in use and therfore is locked and cannot be copied at this" & _
                " time. Please ensure that no one is using the file and try again.", vbOKOnly, _
                "File Currently in Use"
    ElseIf Err.Number = 53 Then
        MsgBox "The Source File '" & sSource & "' could not be found. Please validate the" & _
                " location and name of the specifed Source File and try again", vbOKOnly, _
                "File Currently in Use"
        Else
            MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
                    Err.Number & vbCrLf & "Error Source: CopyFile" & vbCrLf & _
                    "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    End If
    Exit Function
End Function

' ----------------------------------------------------------------
' Procedure Nom:    ExtrairePiecesJointes
' Sujet:            Extraction de pieces jointes d'une table.
' Procedure Kind:   Function
' Procedure Access: Public
'
'=== Paramètres ===
' sNomTable (String):       Table à utiliser.
' sNomChampPJ (String):     Nom du champ contenant la pj.
' sDossier (String):        Dossier ou extraire les pj, si non indiquer utilise le dossier Temp de l'user.
' bCreateFolder (Boolean):  Flag déterminant si on peux créer le dossier si il n'existe pas.
'==================
'
' Return Type: Boolean TRUE si pas de problème.
'
' Author:  Laurent
' Date:    05/05/2022-05:52
' DateMod: 06/05/2022-05:59
'
' ----------------------------------------------------------------
Public Function ExtrairePiecesJointes(sNomTable As String, _
                                      sNomChampPJ As String, _
                                      Optional sDossier As String, _
                                      Optional bCreateFolder As Boolean = False) As Boolean
On Error GoTo ERR_ExtrairePiecesJointes
    
    Dim oDb         As DAO.Database
    Dim oRst        As DAO.Recordset
    Dim oRstPJ      As DAO.Recordset    '// variable recordset pour faire référence au jeu d'enregistrements du champ de type pièce jointe
    Dim sFichier    As String           '// chemin complet du fichier sur le disque
    Dim bRep        As Boolean

    '// Utiliser le dossier Temp ?
    If (sDossier = vbNullString) Then sDossier = Environ("Temp")
    If (Right(sDossier, 1) <> "\") Then sDossier = sDossier & "\"

    '// Vérification dossier, création si besoin.
    bRep = CheckFolderExists(sDossier)
    If (bRep = False And bCreateFolder = True) Then bRep = CreateNewFolder(sDossier)      '// Si le dossier n'existe pas on le crée.

    If (bRep = False) Then
        MsgBox "Le dossier " & sDossier & vbNewLine & "n'existe pas ou n'as pas pu être créer."
        Exit Function
    End If

    Set oDb = CurrentDb()

    ' ouverture du recordset basé sur la table contenant les pièces jointes
    Set oRst = oDb.OpenRecordset(sNomTable)

    Do Until oRst.EOF                                   ' on parcourt les enregistrements de la table
 
        ' on récupère le recordset lié au champ pièce jointe de l'enregistrement courant
        Set oRstPJ = oRst(sNomChampPJ).Value
 
        ' on parcourt les pièces jointes du champ pièce jointe de l'enregistrement
        Do Until oRstPJ.EOF
 
            sFichier = sDossier & oRstPJ("FileName")    ' on compose le chemin complet du fichier sur le disque
 
            ' on s'assure que le fichier n'existe pas déjà avant de le sauvegarder
            If Dir(sFichier) <> "" Then Kill (sFichier) ' s'il existe déjà, on le supprime
 
            oRstPJ("FileData").SaveToFile sFichier      ' on enregistre le fichier à l'emplacement spécifié
            oRstPJ.MoveNext
 
        Loop
 
        oRst.MoveNext
 
    Loop
 
    If (Not oRst Is Nothing) Then oRst.Close
    oDb.Close
    ExtrairePiecesJointes = True
 
SORTIE_ExtrairePiecesJointes:
    On Error Resume Next
    Set oRstPJ = Nothing
    Set oRst = Nothing
    Set oDb = Nothing

    Exit Function
 
ERR_ExtrairePiecesJointes:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.MD_Utilitaires.ExtrairePiecesJointes, ligne " & Erl & "."
    Resume SORTIE_ExtrairePiecesJointes
End Function

Public Function HasAutoexec(ByRef MsBase As DAO.Database) As Boolean
    Dim oRst As DAO.Recordset
    Dim sSql As String
    
    sSql = "SELECT MSysObjects.Name FROM MSysObjects WHERE MSysObjects.Name = 'AutoExec' AND MSysObjects.Type = -32766"
    
    Set oRst = MsBase.OpenRecordset(sSql)
    If Not (oRst.EOF And oRst.BOF) Then HasAutoexec = True

    oRst.Close
    Set oRst = Nothing

End Function

Public Function GetStartUpForm(ByRef MsBase As DAO.Database) As String
    Dim oProp As DAO.Property

    For Each oProp In MsBase.Properties
        If oProp.Name = "StartUpForm" Then
            GetStartUpForm = oProp.Value
            Exit For
        End If
    Next
    Set oProp = Nothing

End Function

Public Function NavigationPane(bShow As Boolean) As Boolean
On Error GoTo ERR_NavigationPane

    Select Case bShow
        Case True
            DoCmd.SelectObject acForm, , True
        Case False
            DoCmd.NavigateTo "acNavigationCategoryObjectType"
            DoCmd.RunCommand acCmdWindowHide
    End Select

SORTIE_ErrHandler:
    Exit Function

ERR_NavigationPane:
    MsgBox "Erreur " & Err.Number & " dans NavigationPane routine : " & vbCrLf & Err.Description, vbOKOnly + vbCritical
    Resume SORTIE_ErrHandler
End Function

'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ END PUB. SUB/FUNC \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'// ################################ PRIVATE SUB/FUNC ####################################
'// Retourne en clair le type de l'objet.
Private Function ObjectTypeName(eType As T_ObjectType) As String
    Dim sType As String

    Select Case eType
        Case Tables_Local
            sType = "Table locale"
        Case Tables_Linked_ODBC
            sType = "Table liée (ODBC)l"
        Case Tables_Linked
            sType = "Table liée"
        Case QueriesType
            sType = "Requête"
        Case FormsType
            sType = "Formulaire"
        Case ReportsType
            sType = "Etat"
        Case MacrosType
            sType = "Macro"
        Case ModulesType
            sType = "Module"
        Case Else
            sType = "???"   'TODO: stype ="???"
    End Select

    ObjectTypeName = sType

End Function
'// ################################# END PRIV. SUB/FUNC #################################

