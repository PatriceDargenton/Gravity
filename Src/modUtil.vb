
Module modUtil

    ' Attribut pour éviter que l'IDE s'interrompt en cas d'exception
    '<System.Diagnostics.DebuggerStepThrough()> _
    Public Function iConv%(sVal$, Optional iValDef% = -1)

        If String.IsNullOrEmpty(sVal) Then iConv = iValDef : Exit Function

        Try
            iConv = CInt(sVal)
        Catch
            iConv = iValDef
        End Try

    End Function

    Public Function bDossierExiste(sCheminDossier$, Optional bPrompt As Boolean = False) As Boolean

        ' Retourne True si un dossier correspondant au filtre sFiltre est trouvé

        'Dim di As New IO.DirectoryInfo(sCheminDossier)
        'bDossierExiste = di.Exists()

        bDossierExiste = IO.Directory.Exists(sCheminDossier)

        If Not bDossierExiste And bPrompt Then _
            MsgBox("Impossible de trouver le dossier :" & vbLf & sCheminDossier,
                MsgBoxStyle.Critical, sTitreMsg & " - Dossier introuvable")

    End Function

    Public Sub AfficherMsgErreur2(ByRef Ex As Exception,
        Optional sTitreFct$ = "", Optional sInfo$ = "",
        Optional sDetailMsgErr$ = "",
        Optional bCopierMsgPressePapier As Boolean = True,
        Optional ByRef sMsgErrFinal$ = "")

        If Not Cursor.Current.Equals(Cursors.Default) Then _
            Cursor.Current = Cursors.Default
        Dim sMsg$ = ""
        If sTitreFct <> "" Then sMsg = "Fonction : " & sTitreFct
        If sInfo <> "" Then sMsg &= vbCrLf & sInfo
        If sDetailMsgErr <> "" Then sMsg &= vbCrLf & sDetailMsgErr
        If Ex.Message <> "" Then
            sMsg &= vbCrLf & Ex.Message.Trim
            If Not IsNothing(Ex.InnerException) Then _
                sMsg &= vbCrLf & Ex.InnerException.Message
        End If
        If bCopierMsgPressePapier Then CopierPressePapier(sMsg)
        sMsgErrFinal = sMsg
        MsgBox(sMsg, MsgBoxStyle.Critical)

    End Sub

    Public Sub CopierPressePapier(sInfo$)

        ' Copier des informations dans le presse-papier de Windows
        ' (elles resteront jusqu'à ce que l'application soit fermée)

        Try
            Dim dataObj As New DataObject
            dataObj.SetData(DataFormats.Text, sInfo)
            Clipboard.SetDataObject(dataObj)
        Catch ex As Exception
            ' Le presse-papier peut être indisponible
            AfficherMsgErreur2(ex, "CopierPressePapier",
                bCopierMsgPressePapier:=False)
        End Try

    End Sub

    Public Function bAppliDejaOuverte(Optional bMemeExe As Boolean = True) As Boolean

        ' Détecter si l'application est déja lancée :
        ' - depuis n'importe quelle copie de l'exécutable (bMemeExe=False), ou bien seulement :
        ' - depuis le même emplacement du fichier exécutable sur le disque dur (bMemeExe=True : par défaut)

        Dim sExeProcessAct$ = Diagnostics.Process.GetCurrentProcess.MainModule.ModuleName
        Dim sNomProcessAct$ = IO.Path.GetFileNameWithoutExtension(sExeProcessAct)

        If Not bMemeExe Then
            ' Détecter si l'application est déja lancée depuis n'importe quel exe
            If Process.GetProcessesByName(sNomProcessAct).Length > 1 Then Return True
            Return False
        End If

        ' Détecter si l'application est déja lancée depuis le même exe
        Dim sCheminProcessAct$ = Diagnostics.Process.GetCurrentProcess.MainModule.FileName
        Dim aProcessAct As Diagnostics.Process() = Process.GetProcessesByName(sNomProcessAct)
        Dim processAct As Diagnostics.Process
        Dim iNbApplis% = 0
        For Each processAct In aProcessAct
            Dim sCheminExe$ = processAct.MainModule.FileName
            If sCheminExe = sCheminProcessAct Then iNbApplis += 1
        Next
        If iNbApplis > 1 Then Return True
        Return False

    End Function

End Module