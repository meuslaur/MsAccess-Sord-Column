VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUtilitaires"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name:    CUtilitaires
' Kind:    Module de classe
' Purpose: Classe pour le module MD_Utilitaires
' Author:  Laurent
' Date:    30/04/2022 - 14:04
' DateMod:
' ------------------------------------------------------
Option Compare Database
Option Explicit

'//::::::::::::::::::::::::::::::::::    VARIABLES      ::::::::::::::::::::::::::::::::::
'//:::::::::::::::::::::::::::::::::: END VARIABLES ::::::::::::::::::::::::::::::::::::::

'//----------------------------------       INIT/TERM       ------------------------------
'//---------------------------------------------------------------------------------------

'//==================================       PROP        ==================================
' Req'd Refs: Late Binding  -> None required
'             Early Binding -> Microsoft Scripting Runtime
#Const FSO_EarlyBind = False

#If FSO_EarlyBind = True Then
Public Property Get GetoFSO() As Scripting.FileSystemObject
#Else
Public Property Get GetoFSO() As Object
#End If


#If FSO_EarlyBind = True Then
Static pFSO               As Scripting.FileSystemObject
#Else
Static pFSO               As Object
#End If
    
    If pFSO Is Nothing Then
        #If FSO_EarlyBind = True Then
            Set pFSO = New FileSystemObject
        #Else
            Set pFSO = CreateObject("Scripting.FileSystemObject")
        #End If
    End If

    Set GetoFSO = pFSO

End Property
'//====================================== END PROP =======================================

'//&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&     EVENTS        &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'//&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& END EVENTS &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&



'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PUBLIC SUB/FUNC   \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ END PUB. SUB/FUNC \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'// ################################ PRIVATE SUB/FUNC ####################################
'// ################################# END PRIV. SUB/FUNC #################################

