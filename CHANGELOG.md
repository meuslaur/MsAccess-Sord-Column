# Changelog
## 0.30

- Ajoute options affichage des images des CommandButton
- Modification controle saisie si optImages est a On.
- Suppréssion de la fonction `VerifDossierImage(sRet)`, utilise le chemin complet.
- Modification indication des erreurs de saisie (label en rouge).
- Ajoute message d'information si aucune base sélectionnée.
- Ajoute 'OptionOn' sur la fonction du form si des images sélectionnées.
- Ajoute code insertion picfolder picasc et picdesc dans l'initialisation de la classe de la function du formulaire.
- Modification du code d'importation de la classe `CsordFormColumn` utilise CopyModule et plus la table T_Info.

## 0.20.5

### Changes

- Modification du code pour utilisation d'un champs String Long dans la table T_Info, a la place du champ dbAttachment.

- Ajout de la fonction `ExtraireCode`.
  
```VB
Private Function ExtraireCode(sID As String, Optional sCtrName As String) As String
```

## 0.20.0

### Changes

- Correction variable du dossier des images.

- Correstion problème enregistrement de la classe.

- Correction code `VerifDossierImage` retourne le nom correct du dossier des images.

```VB
...
    sDosBase = oFSO.GetParentFolderName(m_cCreate.GetBaseFullName)

    If (InStr(sPath, sDosBase) = 0) Then
        MsgBox "Le dossier des images doit être un sous-dossier de l'application", vbExclamation, "Vérification dossier images"
        Exit Function
    End If

    '// Retourne que le dossier des images.
    VerifDossierImage = "\" & oFSO.GetBaseName(sPath) & "\"
...

```
