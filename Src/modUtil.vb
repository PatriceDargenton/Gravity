
Imports System.Text

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

    Public Function sInfoRamDll$(Optional sMsg$ = "RAM : ")

        Dim x As Process = System.Diagnostics.Process.GetCurrentProcess
        Dim lRamAllocatedInApp& = x.WorkingSet64

        Dim sRamAllocatedInApp$ = sFormaterTailleOctets(lRamAllocatedInApp)

        If Not Is64BitProcess() Then
            Dim lRamAvailable32 As ULong = CULng(1.6 * 1024 * 1024 * 1024) ' 1.6 Gb
            If lRamAvailable32 < My.Computer.Info.AvailablePhysicalMemory Then
                lRamAvailable32 = My.Computer.Info.AvailablePhysicalMemory
            End If
            Dim sRamAvailable32$ = sFormaterTailleOctets(CLng(lRamAvailable32))
            Dim rPCRAMUsed32! = CSng(lRamAllocatedInApp / lRamAvailable32)
            Dim sRam32$ = sMsg & sRamAllocatedInApp & " / " & sRamAvailable32 & " (" & rPCRAMUsed32.ToString("0.0 %") & ")"
            Return sRam32
        End If

        Dim lRamAvailable As ULong = My.Computer.Info.AvailablePhysicalMemory
        Dim sRamAvailable$ = sFormaterTailleOctets(CLng(lRamAvailable))
        Dim lRamTot As ULong = My.Computer.Info.TotalPhysicalMemory
        Dim sRamTot$ = sFormaterTailleOctets(CLng(lRamTot))
        Dim lAllocatedTot As ULong = lRamTot - lRamAvailable
        Dim sRamAllocatedTot$ = sFormaterTailleOctets(CLng(lAllocatedTot))
        Dim lAllocatedOther As ULong = CULng(lAllocatedTot - lRamAllocatedInApp)
        Dim sRamAllocatedOther$ = sFormaterTailleOctets(CLng(lAllocatedOther))

        Dim rPCRAMUsed! = CSng(lAllocatedTot / lRamTot)
        Dim sRam$ = sMsg & sRamAllocatedInApp & " + " & sRamAllocatedOther & " = " & sRamAllocatedTot &
            " / " & sRamTot & " (" & rPCRAMUsed.ToString("0.0 %") & ")"
        Return sRam

    End Function

    Public Function Is64BitProcess() As Boolean
        Return (IntPtr.Size = 8)
    End Function

    Public Function sFormaterTailleOctets$(lSizeInBytes&,
                Optional bDetail As Boolean = False,
                Optional bRemoveDotZero As Boolean = False)

        ' https://fr.wikipedia.org/wiki/Octet

        Dim rNbKo! = CSng(Math.Round(lSizeInBytes / 1024, 1))
        Dim rNbMo! = CSng(Math.Round(lSizeInBytes / (1024 * 1024), 1))
        Dim rNbGo! = CSng(Math.Round(lSizeInBytes / (1024 * 1024 * 1024), 1))
        Dim sAff$ = ""

        If bDetail Then
            sAff = sFormaterNumerique(lSizeInBytes) & " octets"
            If rNbKo >= 1 Then sAff &= " (" & sFormaterNumerique(rNbKo) & " Ko"
            If rNbMo >= 1 Then sAff &= " = " & sFormaterNumerique(rNbMo) & " Mo"
            If rNbGo >= 1 Then sAff &= " = " & sFormaterNumerique(rNbGo) & " Go"
            If rNbKo >= 1 Or rNbMo >= 1 Or rNbGo >= 1 Then sAff &= ")"
        Else
            If rNbGo >= 1 Then
                sAff = sFormaterNumerique(rNbGo, bRemoveDotZero) & " Go"
            ElseIf rNbMo >= 1 Then
                sAff = sFormaterNumerique(rNbMo, bRemoveDotZero) & " Mo"
            ElseIf rNbKo >= 1 Then
                sAff = sFormaterNumerique(rNbKo, bRemoveDotZero) & " Ko"
            Else
                sAff = sFormaterNumerique(lSizeInBytes,
                    bSupprimerPointZero:=True) & " octets"
            End If
        End If

        sFormaterTailleOctets = sAff

    End Function

    Public Function sFormaterTailleKOctets$(lSizeInBytes&,
            Optional bRemoveDotZero As Boolean = False)

        Dim rNbKb! = CSng(Math.Ceiling(lSizeInBytes / 1024))
        sFormaterTailleKOctets = sFormaterNumerique(rNbKb, bRemoveDotZero) & " Ko"

    End Function

    Public Function sFormaterNumerique$(rVal!,
            Optional bSupprimerPointZero As Boolean = True,
            Optional iNbDecimales% = 1)

        Dim nfi As New Globalization.NumberFormatInfo With {
            .NumberGroupSeparator = " ",
            .NumberDecimalSeparator = ".",
            .NumberGroupSizes = New Integer() {3, 3, 3},
            .NumberDecimalDigits = iNbDecimales
        }

        Dim sAff$ = rVal.ToString("n", nfi)
        If bSupprimerPointZero Then
            If iNbDecimales = 1 Then
                sAff = sAff.Replace(".0", "")
            ElseIf iNbDecimales > 1 Then
                Dim i%
                Dim sb As New StringBuilder(".")
                For i = 1 To iNbDecimales : sb.Append("0") : Next
                sAff = sAff.Replace(sb.ToString, "")
            End If
        End If
        Return sAff

    End Function

End Module