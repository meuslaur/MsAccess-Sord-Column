Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Private cUtil As New CUtilitaires


Public Sub test()
Dim bRep As Boolean
bRep = ExtrairePiecesJointes("~T_Info", "pjCode", "Dossier")

End Sub

Private Sub Notepad()
  Shell "notepad.exe c:\test.txt"
End Sub

Public Sub test2()
On Error GoTo ERR_test2

    Dim oMsApp            As Access.Application
    Dim oMsBase           As DAO.Database
    Dim sBaseFullName   As String

    Set oMsApp = New Access.Application
    DoEvents
    oMsApp.Visible = False


    sBaseFullName = "E:\Access365\_Encours\BaseTemp.accdb"
    oMsApp.OpenCurrentDatabase sBaseFullName
    DoEvents

    Set oMsBase = oMsApp.CurrentDb
    If (oMsBase Is Nothing) Then
        MsgBox "Impossible d'ouvrir la base " & sBaseFullName & vbCrLf & _
                "Elle est peu-être déjà ouverte", vbExclamation, "ouvrebase"
    End If

    Dim sRep As String
    sRep = CopyModule("CUtilitaires", Application.VBE.ActiveVBProject, oMsApp.VBE.ActiveVBProject)

    If (Len(sRep) = 0) Then
        oMsApp.DoCmd.Save acModule, "CUtilitaires"
    Else
        MsgBox sRep, vbExclamation, "test2"
    End If

    If (Not oMsBase Is Nothing) Then oMsApp.CloseCurrentDatabase
    DoEvents
    oMsApp.Quit
    Set oMsBase = Nothing
    Set oMsApp = Nothing

    
SORTIE_test2:
    Exit Sub

ERR_test2:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.Module1.test2, ligne " & Erl & "."
    Resume SORTIE_test2
End Sub



Public Sub test5()
Dim sTmp As String
    sTmp = "Private Function FUNCNAME(Optional eActiveImage As T_OnOff = OptionOff, " & vbCrLf & _
                            "Optional eActiveTexte As T_OnOff = OptionOn, " & vbCrLf & _
                            "Optional sPicAsc As String = vbNullString, " & vbCrLf & _
                            "Optional sPicDesc As String = vbNullString, " & vbCrLf & _
                            "Optional sFieldName As String = vbNullString) As Boolean"
Debug.Print Replace(sTmp, "FUNCNAME", "foncname")

End Sub
