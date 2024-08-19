
Imports System.IO ' Pour Path, FileInfo

Module modDepart

#If DEBUG Then
    Public Const bDebug As Boolean = True
    'Public Const bRelease As Boolean = False
#Else
    Public Const bDebug As Boolean = False
    'Public Const bRelease As Boolean = True
#End If

    Public ReadOnly sNomAppli$ = My.Application.Info.Title
    Public ReadOnly sTitreMsg$ = sNomAppli
    Private Const sDateVersionGravity$ = "19/08/2024"
    Public Const sDateVersionAppli$ = sDateVersionGravity

    Public ReadOnly sVersionAppli$ =
        My.Application.Info.Version.Major & "." &
        My.Application.Info.Version.Minor &
        My.Application.Info.Version.Build
    Public ReadOnly m_sTitreApplication$ = sNomAppli '"Gravity.Net Screen Saver"

    ' Variables globales : à utiliser avec modération
    Public glb_bModeConfiguration As Boolean ' Pas d'autre choix ici !
    Public glb_rDateMessageTitre As Double ' Pour debugger seulement

    Public Sub Main()

        If bDebug Then Depart() : Exit Sub

        Try
            Depart()
        Catch ex As Exception
            AfficherMsgErreur2(ex, "Main " & sTitreMsg)
        End Try

    End Sub

    Private Sub Depart()

        If bAppliDejaOuverte() Then Exit Sub

        'Dim asArgs$() = asArgLigneCmd(sArg0)

        ' Vérification de l'existence du fichier de configuration :
        '  son nom est toujours basé sur le nom de l'assemblage,
        '  mais comme celui-ci change lorsque l'écran de veille est 
        '  installé sous Windows ancien (il est converti en nom DOS 8.3), 
        '  ça complique un peu !
        Dim sRepertoire$ = Application.StartupPath ' Solution + simple
        Dim sCheminExe$ = Application.ExecutablePath ' = Asm.Location
        Dim sExtension$ = Path.GetExtension(sCheminExe)
        sExtension = sExtension.ToLower()
        Dim sNomAppli$ = Path.GetFileNameWithoutExtension(sCheminExe)
        Dim sFichierConfig$ = sRepertoire & "\" & sNomAppli & sExtension & ".config"

        If Not File.Exists(sFichierConfig) And sExtension = ".scr" Then
            If m_bDebugModeEcranVeille Then _
                MsgBox("Le fichier " & sFichierConfig & vbCrLf &
                    "n'existe pas",
                    MsgBoxStyle.OkOnly Or MsgBoxStyle.Information,
                    m_sTitreApplication)

            ' Si l'extension est .scr et que le fichier .exe.config
            '  existe, on le renomme en .scr.config
            Dim sFichierConfigTrouve$ = ""
            Dim sFichierConfigCherche$

            sFichierConfigCherche = sRepertoire & "\" & sNomAppli & ".exe.config"
            If File.Exists(sFichierConfigCherche) Then _
                sFichierConfigTrouve = sFichierConfigCherche : GoTo Suite

            sFichierConfigCherche = sRepertoire & "\Gravity2.scr.config"
            If File.Exists(sFichierConfigCherche) Then _
                sFichierConfigTrouve = sFichierConfigCherche : GoTo Suite

            sFichierConfigCherche = sRepertoire & "\Gravity2.exe.config"
            If File.Exists(sFichierConfigCherche) Then _
                sFichierConfigTrouve = sFichierConfigCherche ': GoTo Suite

            ' Lorsque l'écran de veille est installé, le nom
            '  de l'assemblage est tronqué en nom DOS 8.3 !!!
            '  au lieu de GravityNet.scr mieux vaut donc Gravity2.scr
            'sFichierConfigCherche = sRepertoire & "\GRAVIT~1.scr.config"
            'If File.Exists(sFichierConfigCherche) Then _
            '    sFichierConfigTrouve = sFichierConfigCherche : GoTo Suite

Suite:
            If sFichierConfigTrouve <> "" Then
                Dim oFileInfo As New FileInfo(sFichierConfigTrouve)
                oFileInfo.MoveTo(sFichierConfig)
                If m_bDebugModeEcranVeille Then _
                    MsgBox("Le fichier " & sFichierConfigTrouve & vbCrLf &
                        "a été renommé en " & sFichierConfig,
                        MsgBoxStyle.OkOnly Or MsgBoxStyle.Information,
                        m_sTitreApplication)
            End If
        End If


        Dim sSeparators$ = " "
        Dim sCommands$ = Microsoft.VisualBasic.Command()
        Dim asArgs$() = sCommands.Split(sSeparators.ToCharArray)
        Dim sArgument$ = ""

        ' Extraire l'option passée en argument de la ligne de commande
        If asArgs.Length > 0 Then sArgument = asArgs(0)

        ' Un handle est parfois passé après l'option, 
        '  par ex.: /C:5833086 d'où le Left de 2
        If sArgument <> "" Then _
            sArgument = sArgument.Substring(0, 2)
        ' Autre solution :
        'sArgument = Microsoft.VisualBasic.Left(sArgument, 2)

        ' Les arguments sont parfois en minuscules,
        '  parfois en majuscules
        sArgument = sArgument.ToUpper()
        'sArgument = UCase(sArgument) ' Autre solution

        ' Lancement depuis VisualSudio.Net : pas d'argument
        ' Le menu Configurer avec le bouton droit de la souris
        '  sur le fichier .scr ne renvoie pas d'argument non plus
        ' Dans ce cas, on choisit donc /C sauf pour debugger
        '  le mode écran de veille (pas le mode configuration)
        If Not m_bDebugModeEcranVeille Then _
            If sArgument = "" Then sArgument = "/C"
        ' Autre solution : passer "/C" dans les arguments
        '  de VisualSudio.Net en mode Release seulement
        '  mais cela ne gère pas le menu Configurer avec le 
        '  bouton droit de la souris sur le fichier .scr

        'MsgBox("Argument : [" & sArgument & "]")

        glb_bModeConfiguration = False
        ' Options des Propriétés de l'Affichage, onglet "Ecran de veille"
        If sArgument = "/C" Then
            ' Bouton "Paramètre..."
            glb_bModeConfiguration = True
            ' Il n'y a plus de Form spéciale pour la configuration
        ElseIf sArgument = "/S" Then
            ' Bouton "Aperçu" et aussi avec le bouton droit 
            '  de la souris : Menu Aperçu
            '//Start the screen saver normally.
        ElseIf sArgument = "/A" Then
            ' Case à cocher "Protégé par mot de passe"
            ' La gestion des mots de passe pour les écrans de veille 
            '  marchent aussi pour Windows 2000 mais pas de la même façon :
            '  on ne passe pas par l'écran de veille et le mot de passe
            '  ne doit pas être nul
            '//Display the password dialog			 
            '"Passwords are not available for this screen saver"
            MsgBox("Cet écran de veille ne gère pas les mots de passe",
                MsgBoxStyle.OkOnly Or MsgBoxStyle.Information,
                m_sTitreApplication)
            Exit Sub
        ElseIf sArgument = "/P" Then
            ' Mini-aperçu dans l'onglet "Ecran de veille"
            ' Non géré pour le moment, car il faut sous-classer la 
            '  fenêtre des propriétés de l'affichage, voir en VB6 :
            '  GRAVITY SCREEN SAVER : UN ÉCRAN DE VEILLE CHAOTIQUE
            '  https://codes-sources.commentcamarche.net/source/1743
            Exit Sub
        End If

        ' Lancer l'écran de veille pour tous les autres arguments
        '//For any other args --> start
        'Dim frm As New FrmGravityNet()
        'Application.Run(frm)
        'clsUtil.JolieTransitionTaDaaa(frm)
        Application.Run(New frmGravityNet())

    End Sub

End Module