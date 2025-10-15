Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(objFSO.GetParentFolderName(WScript.ScriptFullName))

WScript.Echo "Dossier de travail : " & objFolder.Path

' Lire tous les numéros de série dans NumSerieKO.txt
Dim serials, serial, serialFile
If objFSO.FileExists(objFolder.Path & "\NumSerieKO.txt") Then
    WScript.Echo "Lecture du fichier NumSerieKO.txt"
    Set serialFile = objFSO.OpenTextFile(objFolder.Path & "\NumSerieKO.txt", 1)
    serials = Split(serialFile.ReadAll, vbCrLf)
    serialFile.Close
Else
    WScript.Echo "Fichier NumSerieKO.txt introuvable dans : " & objFolder.Path
    WScript.Quit 1
End If

WScript.Echo "Nombre de numéros de série à traiter : " & UBound(serials) + 1

' Parcourir chaque numéro de série
For Each serial In serials
    serial = Trim(serial)
    If serial <> "" Then
        WScript.Echo "\n---\nTraitement du numéro : " & serial
        result = ""
        foundMatches = 0
        ' Rechercher dans chaque fichier ELISA_Prod_Log*
        For Each logfile In objFolder.Files
            If LCase(Left(logfile.Name, 15)) = "elisa_prod_log" Then
                WScript.Echo "Analyse du fichier : " & logfile.Name
                Set objFile = objFSO.OpenTextFile(logfile.Path, 1)
                collecting = False
                found = False
                lineNumber = 0
                Do Until objFile.AtEndOfStream
                    line = objFile.ReadLine
                    lineNumber = lineNumber + 1
                    If InStr(line, "Datamatrix") > 0 Then
                        tmp = Split(line, "#")
                        If UBound(tmp) >= 2 Then
                            sn = Trim(tmp(2))
                            sn = Replace(sn, "=", "")
                            sn = Replace(sn, " ", "")
                            If InStr(sn, "\t") > 0 Then sn = Replace(sn, "\t", "")
                            WScript.Echo "Ligne " & lineNumber & " : Datamatrix avec numéro " & sn
                            If sn = serial Then
                                If collecting = False Then
                                    foundMatches = foundMatches + 1
                                    If result <> "" Then result = result & vbCrLf
                                End If
                                collecting = True
                                WScript.Echo "--> Numéro correspondant trouvé"
                                If result <> "" And Right(result, 2) <> vbCrLf Then
                                    result = result & vbCrLf
                                End If
                                result = result & line
                            Else
                                If collecting Then WScript.Echo "Ligne " & lineNumber & " : Fin du bloc pour " & serial
                                collecting = False
                            End If
                        Else
                            WScript.Echo "Ligne " & lineNumber & " : Datamatrix mal formé (moins de 3 éléments)"
                            collecting = False
                        End If
                    ElseIf collecting Then
                        result = result & vbCrLf & line
                            WScript.Echo "Ligne " & lineNumber & " : Datamatrix avec numéro " & sn
                            If sn = serial Then
                                WScript.Echo "--> Numéro correspondant trouvé"
                                If result <> "" Then result = result & vbCrLf
                                result = result & line
                                found = True
                                foundMatches = foundMatches + 1
                                ' Ajouter les lignes suivantes jusqu'à la prochaine ligne "Datamatrix"
                                Do Until objFile.AtEndOfStream
                                    l2 = objFile.ReadLine
                                    lineNumber = lineNumber + 1
                                    If InStr(l2, "Datamatrix") > 0 Then
                                        WScript.Echo "Ligne " & lineNumber & " : Nouveau bloc Datamatrix, arrêt de la capture"
                                        objFile.SkipLine  ' revenir sur la ligne "Datamatrix" à la prochaine itération
                                        Exit Do
                                    End If
                                    result = result & vbCrLf & l2
                                Loop
                            End If
                        Else
                            WScript.Echo "Ligne " & lineNumber & " : Datamatrix mal formé (moins de 3 éléments)"
                        End If
                    End If
                Loop
                objFile.Close
            End If
        Next
        ' Ecrire le résultat dans un fichier si on a trouvé des lignes
        If result <> "" Then
            Set f = objFSO.CreateTextFile(objFolder.Path & "\" & serial & ".txt", True)
            f.Write result
            f.Close
            WScript.Echo "Résultats écrits dans " & serial & ".txt (" & foundMatches & " bloc(s))"
        Else
            WScript.Echo "Aucun résultat trouvé pour ce numéro"
        End If
    Else
        WScript.Echo "Numéro vide rencontré, ignoré"
    End If
Next
