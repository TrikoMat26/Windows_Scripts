Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(objFSO.GetParentFolderName(WScript.ScriptFullName))

' Lire tous les numéros de série dans NumSerieKO.txt
Dim serials, serial
serials = Split(objFSO.OpenTextFile("NumSerieKO.txt", 1).ReadAll, vbCrLf)

' Parcourir chaque numéro de série
For Each serial In serials
    If Trim(serial) <> "" Then
        result = ""
        ' Rechercher dans chaque fichier ELISA_Prod_Log*
        For Each logfile In objFolder.Files
            If LCase(Left(logfile.Name, 15)) = "ELISA_Prod_Log" Then
                Set objFile = objFSO.OpenTextFile(logfile.Path, 1)
                found = False
                Do Until objFile.AtEndOfStream
                    line = objFile.ReadLine
                    If InStr(line, "Datamatrix") > 0 Then
                        tmp = Split(line, "#")
                        If UBound(tmp) >= 2 Then
                            sn = Trim(tmp(2))
                            If sn = serial Then
                                If result <> "" Then result = result & vbCrLf
                                result = result & line
                                found = True
                                ' Ajouter les lignes suivantes jusqu'à la prochaine ligne "Datamatrix"
                                Do Until objFile.AtEndOfStream
                                    pos = objFile.Line
                                    l2 = objFile.ReadLine
                                    If InStr(l2, "Datamatrix") > 0 Then
                                        objFile.SkipLine  ' revenir sur la ligne "Datamatrix" à la prochaine itération
                                        Exit Do
                                    End If
                                    result = result & vbCrLf & l2
                                Loop
                            End If
                        End If
                    ElseIf found Then
                        ' Si on traitait une séquence, mais c'est géré dans la boucle ci-dessus
                        found = False
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
        End If
    End If
Next