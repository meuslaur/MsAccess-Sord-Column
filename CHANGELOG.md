# Changelog

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
