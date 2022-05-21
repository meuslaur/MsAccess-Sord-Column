Attribute VB_Name = "MD_VbComp"
'@Folder("Outils")
' ------------------------------------------------------
' Name:    MD_VbComp
' Kind:    Module
' Purpose: Outils pour VBE
' Author:  Laurent
' Date:    30/04/2022 - 14:07
' DateMod:
' ------------------------------------------------------
Option Compare Database
Option Explicit

'//&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&     EVENTS        &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'//&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& END EVENTS &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&


'//==================================       PROP        ==================================
'//====================================== END PROP =======================================

'// ################################ PRIVATE SUB/FUNC ####################################
'// ################################# END PRIV. SUB/FUNC #################################



'//::::::::::::::::::::::::::::::::::    VARIABLES      ::::::::::::::::::::::::::::::::::
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal ClassName As String, ByVal WindowName As String) As Long
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
'//:::::::::::::::::::::::::::::::::: END VARIABLES ::::::::::::::::::::::::::::::::::::::

'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PUBLIC SUB/FUNC   \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Lorsque vous avez utilisé le code d’extensibilité, la fenêtre de l’éditeur VBA clignote. Cela peut être réduit avec le code:
'Application.VBE.MainWindow.Visible = False
'Cela masquera la fenêtre VBE, mais vous pouvez toujours la voir scintiller. Pour éviter cela, vous devez utiliser la fonction API Windows LockWindowUpdate.
Public Sub EliminateScreenFlicker(AccAp As Access.Application)
    On Error GoTo ErrH:
    Dim VBEHwnd As Long
    
    
    AccAp.VBE.MainWindow.Visible = False
    
    VBEHwnd = FindWindow("wndclass_desked_gsk", AccAp.VBE.MainWindow.Caption)
    
    If VBEHwnd Then LockWindowUpdate VBEHwnd
    
    '''''''''''''''''''''''''
    ' your code here
    '''''''''''''''''''''''''
    
    'AccAp.VBE.MainWindow.Visible = False
    Exit Sub
ErrH:
    LockWindowUpdate 0&
End Sub

   
'   Il n’existe aucun moyen direct de copier un module d’un projet à un autre. Pour accomplir cette tâche, vous devez exporter le module à partir du VBProject source,
'   puis importer ce fichier dans le VBProject de destination. Le code ci-dessous le fera. La déclaration de fonction est la suivante :
'
'   Function CopyModule(ModuleName As String, _
'       FromVBProject As VBIDE.VBProject, _
'       ToVBProject As VBIDE.VBProject, _
'       OverwriteExisting As Boolean) As Boolean
'
'   ModuleName est le nom du module que vous souhaitez copier d’un projet à un autre.
'
'   FromVBProject est le VBProject qui contient le module à copier. Il s’agit de la source VBProject.
'
'   ToVBProject est le VBProject dans lequel le module doit être copié. Il s’agit de la destination VBProject.
'
'   OverwriteExisting indique ce qu’il faut faire si ModuleName existe déjà dans ToVBProject. Si cette valeur est True, le VBComponent existant sera supprimé du ToVBProject. Si la valeur est False et que VBComponent existe déjà, la fonction ne fait rien et renvoie False.
'
'La fonction renvoie True si une erreur réussit ou False est une erreur. La fonction renvoie False si l’une des valeurs suivantes est vraie :
'   FromVBProject n’est rien.
'   ToVBProject n’est rien.
'   ModuleName est vide.
'   FromVBProject est verrouillé.
'   ToVBProject est verrouillé.
'   ModuleName n’existe pas dans FromVBProject.
'   ModuleName existe dans ToVBProject et OverwriteExisting a la valeur False.
'
Public Function CopyModule(ModuleName As String, _
                    FromVBProject As VBIDE.VBProject, _
                    ToVBProject As VBIDE.VBProject, _
                    Optional OverwriteExisting As Boolean = False) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' CopyModule
    ' This function copies a module from one VBProject to
    ' another. It returns True if successful or  False
    ' if an error occurs.
    '
    ' Parameters:
    ' --------------------------------
    ' FromVBProject         The VBProject that contains the module
    '                       to be copied.
    '
    ' ToVBProject           The VBProject into which the module is
    '                       to be copied.
    '
    ' ModuleName            The name of the module to copy.
    '
    ' OverwriteExisting     If True, the VBComponent named ModuleName
    '                       in ToVBProject will be removed before
    '                       importing the module. If False and
    '                       a VBComponent named ModuleName exists
    '                       in ToVBProject, the code will return
    '                       False.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim VBComp      As VBIDE.VBComponent
    Dim TempVBComp  As VBIDE.VBComponent
    Dim FName       As String
    Dim CompName    As String
    Dim sLine       As String
    Dim SlashPos    As Long
    Dim ExtPos      As Long

    '''''''''''''''''''''''''''''''''''''''''''''
    ' Do some housekeeping validation.
    '''''''''''''''''''''''''''''''''''''''''''''
    If FromVBProject Is Nothing Then
        CopyModule = "VBIDE.VBProject source non initialisé."
        Exit Function
    End If
    
    If Trim$(ModuleName) = vbNullString Then
        CopyModule = "Valeur de ModuleName est Null."
        Exit Function
    End If
    
    If ToVBProject Is Nothing Then
        CopyModule = "VBIDE.VBProject destination non initialisé."
        Exit Function
    End If
    
    If FromVBProject.Protection = vbext_pp_locked Then
        CopyModule = "Le projet source est vérouillé pour l'affichage."
        Exit Function
    End If
    
    If ToVBProject.Protection = vbext_pp_locked Then
        CopyModule = "Le projet destination est vérouillé pour l'affichage."
        Exit Function
    End If
    
    On Error Resume Next
    Set TempVBComp = FromVBProject.VBComponents(ModuleName)
    If Err.Number <> 0 Then
        CopyModule = "Le module : " & ModuleName & " n'existe pas dans le projet source."
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FName is the name of the temporary file to be
    ' used in the Export/Import code.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    FName = Environ$("Temp") & "\" & ModuleName & ".bas"
    If OverwriteExisting = True Then
        ''''''''''''''''''''''''''''''''''''''
        ' If OverwriteExisting is True, Kill
        ' the existing temp file and remove
        ' the existing VBComponent from the
        ' ToVBProject.
        ''''''''''''''''''''''''''''''''''''''
        If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
            Err.Clear
            Kill FName
            If Err.Number <> 0 Then
                CopyModule = "Erreur suppression du fichier : " & FName
                Exit Function
            End If
        End If
        '// Supprime le module.
        With ToVBProject.VBComponents
'            .Remove .Item(ModuleName)
        End With
    Else
        '''''''''''''''''''''''''''''''''''''''''
        ' OverwriteExisting is False. If there is
        ' already a VBComponent named ModuleName,
        ' exit with a return code of False.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        Set VBComp = ToVBProject.VBComponents(ModuleName)
        If (Err.Number <> 0) Then
            If (Err.Number <> 9) Then
                ' other error. get out with return value of False
                CopyModule = "Erreur :" & Err.Description & vbCrLf & "N°:" & Err.Number
                Exit Function
            End If
        Else
            '// Le module exite, et OverwriteExisting False.
            CopyModule = "le module " & ModuleName & " existe déjà dans le projet source."
            Exit Function
        End If
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Do the Export and Import operation using FName
    ' and then Kill FName.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    FromVBProject.VBComponents(ModuleName).Export FileName:=FName
    
    '''''''''''''''''''''''''''''''''''''
    ' Extract the module name from the
    ' export file name.
    '''''''''''''''''''''''''''''''''''''
    SlashPos = InStrRev(FName, "\")
    ExtPos = InStrRev(FName, ".")
    CompName = Mid$(FName, SlashPos + 1, ExtPos - SlashPos - 1)
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Document modules (SheetX and ThisWorkbook)
    ' cannot be removed. So, if we are working with
    ' a document object, delete all code in that
    ' component and add the lines of FName
    ' back in to the module.
    ''''''''''''''''''''''''''''''''''''''''''''''
    Set VBComp = Nothing
    Set VBComp = ToVBProject.VBComponents(CompName)
    
    If VBComp Is Nothing Then
    '// Le module n'existe pas on import le fichier.
        ToVBProject.VBComponents.Import FileName:=FName
    Else
        '// Le module existe et OverwriteExisting a True,
        '// Supprime toute les ligne du module et colle le nouveau code.
        If VBComp.Type = vbext_ct_ClassModule Or VBComp.Type = vbext_ct_StdModule Then
            ' VBComp is destination module
            Set TempVBComp = ToVBProject.VBComponents.Import(FName)
            ' TempVBComp is source module
            With VBComp.CodeModule
                .DeleteLines 1, .CountOfLines
                sLine = TempVBComp.CodeModule.Lines(1, TempVBComp.CodeModule.CountOfLines)
                .InsertLines 1, sLine
            End With
            'On Error GoTo 0
            ToVBProject.VBComponents.Remove TempVBComp
        End If
    End If

    Kill FName
    CopyModule = vbNullString

End Function

' ----------------------------------------------------------------
' Procedure Nom:    ModuleExiste
' Sujet:            Vérifier si un module exisqte dajà dans la base
' Procedure Kind:   Function
' Procedure Access: Public
' Références:       Vérifier si un module exisqte dajà dans la base
'
'=== Paramètres ===
' sModuleName ():       Nom du Module
' VbProjet (VBProject): Projet à utiliser pour la recherche
'==================
'
' Return Type: Boolean Ttue si le module existe dans la base.
'
' Author:  Laurent
' Date:    11/05/2022 - 19:02
' DateMod:
' ----------------------------------------------------------------
Public Function ModuleExiste(sModuleName As String, oVBProjet As VBIDE.VBProject) As Boolean
    On Error Resume Next
    Dim oVBComp As VBIDE.VBComponent
    
    Set oVBComp = oVBProjet.VBComponents(sModuleName)
    ModuleExiste = (Not oVBComp Is Nothing)
    Set oVBComp = Nothing
End Function

'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ END PUB. SUB/FUNC \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

