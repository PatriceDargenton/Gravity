
' Gravity : l'écran de veille chaotique

' Note : Le fichier Gravity2.exe doit être renommé en .scr pour être 
'  installé normalement ainsi que Gravity2.exe.config en Gravity2.src.config
' (voir les évènements de post-build dans les options de compilation)

' Conventions de nommage des variables :
' ------------------------------------
' b pour Boolean (booléen vrai ou faux)
' i pour Integer : % (en VB .Net, l'entier a la capacité du VB6.Long)
' l pour Long : &
' r pour nombre Réel (Single!, Double# ou Decimal : D)
' s pour String : $
' c pour Char ou Byte
' d pour Date
' u pour Unsigned (non signé : entier positif)
' a pour Array (tableau) : ()
' o pour Object : objet instancié localement
' refX pour reference à un objet X préexistant qui n'est pas sensé être fermé
' m_ pour variable Membre de la classe ou de la feuille (Form)
'  (mais pas pour les constantes)
' frm pour Form
' cls pour Classe
' mod pour Module
' ...
' ------------------------------------

Public Class frmGravityNet ' : Inherits Form

#Region "Constantes"

    Private Const sTitreFrmConfig$ = "Configuration de Gravity Screen Saver"

    ' Indexe des paramètres dans la zone de liste listViewPrm
    Private Const iDegreRacine% = 0
    Private Const iDegreRacine_bRnd% = 1
    Private Const iDegreRacineRndMax% = 2
    Private Const iDegreRacine2% = 3
    Private Const iDegreRacine2_bRnd% = 4
    Private Const iDegreRacine2RndMax% = 5
    Private Const ibMasseSym% = 6
    Private Const ibMasseSym_bRnd% = 7
    Private Const ib3D% = 8
    Private Const ib3D_bRnd% = 9

    Private Const ib3D_bPlanetesAxeV% = 10
    Private Const ib3D_bPlanetesAxeV_bRnd% = 11
    Private Const ib3D_iNbPlanetesMaxAxeV% = 12

    Private Const ibChocs% = 13
    Private Const ibChocs_bRnd% = 14
    Private Const iForceGravitation% = 15
    Private Const iTxtBanniere% = 16
    Private Const iFiltreFichiersImgSprite% = 17
    Private Const iFiltreFichiersImgFond% = 18
    Private Const ibMAJToutLEcran% = 19
    Private Const ibMAJGroupeSprites% = 20
    Private Const ibNePasInitFond% = 21
    Private Const ibNePasBufferiserGr% = 22
    Private Const ibFondUni% = 23
    Private Const ibFondDegrade% = 24
    Private Const iTempsMaxScenarioSec% = 25
    Private Const iDelaiMiliSec% = 26
    Private Const ibPauseAnimation% = 27

#End Region

#Region "Déclarations"

    Private m_sTitreAppli$

    Private m_bInit As Boolean
    Private m_iVitesseBanniere% = 5 ' Vitesse de la bannière
    Private m_iPosBanniereX%        ' Position horizontale de la bannière
    ' Pour déterminer si la souris a bougée
    Private m_ptPosSouris As New Point(0, 0)

    Private m_rectEcran As Rectangle
    Private m_bFondInitialise As Boolean
    Private m_grFrm As Graphics ' Pour gérer le graphisme dans la Form
    Private m_bMajListPrm As Boolean
    ' Booléen pour indiquer si tout est configuré
    Private m_bDejaConfigure As Boolean = False

    Private m_bBoucleAnimationEnCours As Boolean = False
    Private m_bQuitterBoucleAnimation As Boolean = False
    Private m_rMemDateDepartAnimation As Double
    Private m_gravity As New SimulteurGravite(Me)

    ' Structures pour les paramètres de l'écran de veille
    Structure TParametresEcran

        Dim sTxtBanniere$
        Dim bAffichageBanniere As Boolean

        Dim bNePasBufferiserGr As Boolean
        Dim bMAJGroupeSprites As Boolean
        Dim bMAJToutLEcran As Boolean
        Dim iTempsMaxScenarioSec%
        Dim iDelaiMiliSec%

        ' Cela permet de dépasser les 100 fps si le tracé est simple,
        '  car le timer est limité à 1 ms au min., ce qui correspond à 100 fps !
        Dim bBoucleAnimation As Boolean

    End Structure

    Public m_prmE As TParametresEcran

#End Region

#Region "Initialisations"

    Private Sub frmGravityNet_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.Text = sTitreFrmConfig

        Dim sVersion$ = " - V" & sVersionAppli & " (" & sDateVersionAppli & ")"
        Dim sDebug$ = " - Debug"
        Dim sTxt$ = Me.Text & sVersion
        If bDebug Then sTxt &= sDebug
        Me.Text = sTxt

        m_sTitreAppli = sTxt

    End Sub

    Private Sub InitialiserEcranDeVeille()

        m_prmE.sTxtBanniere = My.Settings.TxtBanniere
        m_prmE.bAffichageBanniere = My.Settings.bAffichageBanniere
        m_prmE.bMAJToutLEcran = My.Settings.bMAJToutLEcran
        m_prmE.bNePasBufferiserGr = My.Settings.bNePasBufferiserGr
        m_prmE.iTempsMaxScenarioSec = My.Settings.TempsMaxScenarioSec
        m_prmE.iDelaiMiliSec = My.Settings.DelaiMiliSec

        If glb_bModeConfiguration Then

            Me.Text = m_sTitreAppli

            Me.ListViewPrm.Visible = True

            ' Positionnement de la fenêtre par le code : mode manuel
            Me.StartPosition = FormStartPosition.Manual
            Me.Location = My.Settings.frmGravityNetPos
            Me.Size = My.Settings.frmGravityNetTaille

            ' Le ListView n'est pas sizable (sinon ancrer)
            'Me.ListViewPrm.Location = My.Settings.frmConfigPos
            'Me.ListViewPrm.Size = My.Settings.frmConfigTaille

            InitialiserModeConfiguration()

        Else
            If Not bDebug Then
                Cursor.Hide()
                Me.TopMost = True
            End If
            Me.FormBorderStyle = FormBorderStyle.None
            Me.WindowState = FormWindowState.Maximized
            InitialiserTimer()
        End If

    End Sub

    Private Sub InitialiserModeConfiguration()

        InitialiserListePrm()
        Application.DoEvents()
        m_gravity.InitialiserTailleEcran(Me.ClientSize)
        m_bDejaConfigure = True
        MAJAnimation(bTirageAleatoire:=False, bInitialiserFond:=True, bControlerPrm:=True)

    End Sub

    Private Sub InitialiserTimer()

        m_prmE.bBoucleAnimation = False

        If m_bBoucleAnimationEnCours And m_prmE.iDelaiMiliSec > 0 And
            (TimerAnimation.Interval <> m_prmE.iDelaiMiliSec Or
             m_prmE.iDelaiMiliSec = 1) Then

            SauverConfig()
            MsgBox("Ce réglage ne prendra effet qu'au prochain lancement",
                 MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, m_sTitreApplication)

        End If

        If m_prmE.iDelaiMiliSec > 0 Then
            TimerAnimation.Interval = m_prmE.iDelaiMiliSec
        Else
            TimerAnimation.Interval = 1
            m_prmE.bBoucleAnimation = True
        End If

    End Sub

    Private Sub InitialiserListePrm()

        Me.SuspendLayout()

        ' Set the view to show details.
        Me.ListViewPrm.View = View.Details
        ' Allow the user to edit item text.
        Me.ListViewPrm.LabelEdit = True
        ' Display check boxes.
        Me.ListViewPrm.CheckBoxes = True
        ' Display grid lines.
        Me.ListViewPrm.GridLines = True

        ' On peut le faire avec l'éditeur en mode design
        '  mais ça bug pas mal quand même
        Me.ColumnHeader1 = New Windows.Forms.ColumnHeader()
        Me.ColumnHeader2 = New Windows.Forms.ColumnHeader()
        Me.ColumnHeader3 = New Windows.Forms.ColumnHeader()

        Me.ColumnHeader2.Text = "Valeur"
        Me.ColumnHeader2.Width = 130

        Me.ColumnHeader1.Text = "Paramètre"
        Me.ColumnHeader1.Width = 150

        Me.ColumnHeader3.Text = "Explication"
        Me.ColumnHeader3.Width = 0

        Me.ListViewPrm.Columns.AddRange(
            New System.Windows.Forms.ColumnHeader() {
            Me.ColumnHeader2, Me.ColumnHeader1, Me.ColumnHeader3})

        ' Create three items and three sets of subitems for each item.

        Dim aString$(2)
        Const iColNomPrm% = 0
        Const iColExplic% = 2

        Dim sValDegreRacine$ = m_gravity.m_prm.iDegreRacine.ToString ' 1
        Dim item0 As New ListViewItem(CStr(sValDegreRacine))
        aString(iColNomPrm) = sDegreRacine
        aString(iColExplic) = "Degré du système 1 (système principal)"
        item0.SubItems.AddRange(aString)

        Dim bBool As Boolean = m_gravity.m_prm.bDegreRacine_bRnd
        ' Idée : MAJ de la valeur du booléen dans le label
        Dim sChaineVide$ = ""
        Dim item1 As New ListViewItem(sChaineVide)
        item1.Checked = bBool
        aString(iColNomPrm) = sDegreRacine_bRnd
        aString(iColExplic) = "Cochez pour choisir au hasard le degré du système 1"
        item1.SubItems.AddRange(aString)

        Dim iDegreRacineRndMax% = m_gravity.m_prm.iDegreRacineRndMax ' 4
        Dim item2 As New ListViewItem(CStr(iDegreRacineRndMax))
        aString(iColNomPrm) = sDegreRacineRndMax
        aString(iColExplic) = "Degré max. du système 1 si son tirage est aléatoire"
        item2.SubItems.AddRange(aString)

        Dim iDegreRacine2% = m_gravity.m_prm.iDegreRacine2 ' 1
        Dim item3 As New ListViewItem(CStr(iDegreRacine2))
        aString(iColNomPrm) = sDegreRacine2
        aString(iColExplic) =
            "Degré du système 2 (système secondaire imbriqué dans le système 1)"
        item3.SubItems.AddRange(aString)

        bBool = m_gravity.m_prm.bDegreRacine2_bRnd
        Dim item4 As New ListViewItem(sChaineVide)
        item4.Checked = bBool
        aString(iColNomPrm) = sDegreRacine2_bRnd
        aString(iColExplic) = "Cochez pour choisir au hasard le degré du système 2"
        item4.SubItems.AddRange(aString)

        Dim iDegreRacine2RndMax% = m_gravity.m_prm.iDegreRacine2RndMax ' 4
        Dim item5 As New ListViewItem(CStr(iDegreRacine2RndMax))
        aString(iColNomPrm) = sDegreRacine2RndMax
        aString(iColExplic) = "Degré max. du système 2 si son tirage est aléatoire"
        item5.SubItems.AddRange(aString)

        bBool = m_gravity.m_prm.bMasseSym
        Dim item6 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sbMasseSym
        aString(iColExplic) = "Cochez pour que les planètes soient symétriques"
        item6.SubItems.AddRange(aString)
        item6.Checked = bBool

        bBool = m_gravity.m_prm.bMasseSym_bRnd
        Dim item7 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sbMasseSym_bRnd
        aString(iColExplic) =
            "Cochez pour choisir au hasard si les planètes doivent être symétriques"
        item7.SubItems.AddRange(aString)
        item7.Checked = bBool

        bBool = m_gravity.m_prm.b3D
        Dim item8 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sb3D
        aString(iColExplic) = "Cochez pour afficher l'animation en 3 dimensions"
        item8.SubItems.AddRange(aString)
        item8.Checked = bBool

        bBool = m_gravity.m_prm.b3D_bRnd
        Dim item9 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sb3D_bRnd
        aString(iColExplic) = "Cochez pour choisir au hasard si l'animation doit être en 3D"
        item9.SubItems.AddRange(aString)
        item9.Checked = bBool

        bBool = m_gravity.m_prm.b3D_bPlanetesAxeV
        Dim item25 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sb3D_bPlanetesAxeV
        aString(iColExplic) = "Cocher pour ajouter des planètes dans l'axe vertical (3D)"
        item25.SubItems.AddRange(aString)
        item25.Checked = bBool

        bBool = m_gravity.m_prm.b3D_bPlanetesAxeV_bRnd
        Dim item26 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sb3D_bPlanetesAxeV_bRnd
        aString(iColExplic) =
            "Cochez pour choisir au hasard pour ajouter des planètes dans l'axe vertical (3D)"
        item26.SubItems.AddRange(aString)
        item26.Checked = bBool

        Dim iNbPlanetesAxeV% = m_gravity.m_prm.b3D_iNbPlanetesMaxAxeV
        Dim item27 As New ListViewItem(CStr(iNbPlanetesAxeV))
        aString(iColNomPrm) = sb3D_iNbPlanetesMaxAxeV
        aString(iColExplic) = "Nombre max. de planètes à ajouter dans l'axe vertical (3D)"
        item27.SubItems.AddRange(aString)

        bBool = m_gravity.m_prm.bChocs
        Dim item10 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sbChocs
        aString(iColExplic) = "Cochez pour indiquer que les planètes doivent se percuter"
        item10.SubItems.AddRange(aString)
        item10.Checked = bBool

        bBool = m_gravity.m_prm.bChocs_bRnd
        Dim item11 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sbChocs_bRnd
        aString(iColExplic) = "Cochez pour choisir au hasard si les planètes doivent se percuter"
        item11.SubItems.AddRange(aString)
        item11.Checked = bBool

        Dim iForceGravitation% = CInt(m_gravity.m_prm.rForceGravitation) ' 100
        Dim item12 As New ListViewItem(CStr(iForceGravitation))
        aString(iColNomPrm) = sForceGravitation
        aString(iColExplic) = "Paramètre proportionnel à la force de gravitation"
        item12.SubItems.AddRange(aString)

        Dim sValTxtBanniere$ = m_prmE.sTxtBanniere
        Dim item13 As New ListViewItem(sValTxtBanniere)
        aString(iColNomPrm) = sTxtBanniere
        aString(iColExplic) = "Texte de la bannière à afficher (si la case est cochée)"
        item13.SubItems.AddRange(aString)
        ' La case à cocher et le label sont utilisés tous les 2 dans ce cas
        bBool = m_prmE.bAffichageBanniere
        item13.Checked = bBool

        Dim sValFiltreFichiersImgSprite$ = m_gravity.m_prm.sFiltreFichiersImgSprite
        Dim item14 As New ListViewItem(sValFiltreFichiersImgSprite)
        aString(iColNomPrm) = sFiltreFichiersImgSprite
        aString(iColExplic) = "Filtre pour choisir les fichiers images des planètes (si coché)"
        item14.SubItems.AddRange(aString)

        bBool = m_gravity.m_prm.bCercle
        item14.Checked = Not bBool

        Dim sValFiltreFichiersImgFond$ = m_gravity.m_prm.sFiltreFichiersImgFond
        Dim item15 As New ListViewItem(sValFiltreFichiersImgFond)
        aString(iColNomPrm) = sFiltreFichiersImgFond
        aString(iColExplic) = "Filtre pour choisir les fichiers images du fond (si coché)"
        item15.SubItems.AddRange(aString)
        bBool = m_gravity.m_prm.bImageFond
        item15.Checked = bBool

        bBool = m_prmE.bMAJToutLEcran
        Dim item16 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sbMAJToutLEcran
        aString(iColExplic) =
            "Cochez pour mettre à jour tout l'écran à chaque image (frame) de l'animation"
        item16.SubItems.AddRange(aString)
        item16.Checked = bBool

        bBool = m_gravity.m_prm.bMAJGroupeSprites
        Dim item17 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sbMAJGroupeSprites
        aString(iColExplic) =
            "Cochez pour mettre à jour la zone autour des sprites afin de faire une jolie transition"
        item17.SubItems.AddRange(aString)
        item17.Checked = bBool

        bBool = m_gravity.m_prm.bNePasInitFond
        Dim item18 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sbNePasInitFond
        aString(iColExplic) =
            "Cochez pour ne pas initialiser l'image du fond afin de faire une jolie transistion"
        item18.SubItems.AddRange(aString)
        item18.Checked = bBool

        bBool = m_prmE.bNePasBufferiserGr
        Dim item19 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sbNePasBufferiserGr
        aString(iColExplic) =
            "Cochez pour ne pas buffériser le graphisme afin de conserver la trace des sprites"
        item19.SubItems.AddRange(aString)
        item19.Checked = bBool

        bBool = m_gravity.m_prm.bFondUni
        Dim item20 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sbFondUni
        aString(iColExplic) =
            "Cochez pour que le fond soit uni (si les autres options du fond sont désactivées)"
        item20.SubItems.AddRange(aString)
        item20.Checked = bBool

        bBool = m_gravity.m_prm.bFondDegrade
        Dim item21 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = sbFondDegrade
        aString(iColExplic) =
            "Cochez pour que le fond soit un dégradé de couleur (s'il n'y a pas d'image de fond)"
        item21.SubItems.AddRange(aString)
        item21.Checked = bBool

        Dim iTempsMaxScenarioSec% = m_prmE.iTempsMaxScenarioSec
        Dim item22 As New ListViewItem(CStr(iTempsMaxScenarioSec))
        aString(iColNomPrm) = sTempsMaxScenarioSec
        aString(iColExplic) = "Temps max. de l'animation en secondes avant un autre tirage au hasard"
        item22.SubItems.AddRange(aString)

        Dim iDelaiMiliSec% = m_prmE.iDelaiMiliSec ' 0
        Dim item23 As New ListViewItem(CStr(iDelaiMiliSec))
        aString(iColNomPrm) = sDelaiMiliSec
        aString(iColExplic) =
            "Délai en millisecondes entre chaque image de l'animation (0 pour une boucle continue)"
        item23.SubItems.AddRange(aString)

        Dim item24 As New ListViewItem(sChaineVide)
        aString(iColNomPrm) = "Pause"
        aString(iColExplic) = "Cochez pour faire une pause de l'animation"
        item24.SubItems.AddRange(aString)
        item24.Checked = False

        ' Attention : ne pas changer l'ordre des items
        '  car on utilise les index correspondants
        Me.ListViewPrm.Items.AddRange(New ListViewItem() {
            item0, item1, item2, item3, item4, item5, item6,
            item7, item8, item9, item25, item26, item27,
            item10, item11, item12, item13,
            item14, item15, item16, item17, item18, item19, item20,
            item21, item22, item23, item24})

        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Gestion configuration"

    Private Sub ListViewPrm_SelectedIndexChanged(sender As Object,
            e As System.EventArgs) Handles ListViewPrm.SelectedIndexChanged

        Dim iNbItemsSelect% = Me.ListViewPrm.SelectedItems.Count
        If iNbItemsSelect = 0 Then Exit Sub
        Dim itemLignePrm As ListViewItem.ListViewSubItemCollection
        itemLignePrm = Me.ListViewPrm.SelectedItems.Item(0).SubItems
        ' Affichage de la description du paramètre dans la barre de titre
        Me.Text = itemLignePrm(1).Text & " : " & itemLignePrm(2).Text
        glb_rDateMessageTitre = DateAndTime.Timer

    End Sub

    Private Sub ListViewPrm_ItemCheck(sender As Object, e As Windows.Forms.ItemCheckEventArgs) _
            Handles ListViewPrm.ItemCheck

        If Not m_bDejaConfigure Then Exit Sub
        Dim bTirageAleatoire, bInitialiserFond As Boolean
        m_bMajListPrm = True
        If Not bTraiterModifBooleen(e.Index, e.NewValue,
            bTirageAleatoire, bInitialiserFond) Then Exit Sub
        MAJAnimation(bTirageAleatoire, bInitialiserFond,
            bControlerPrm:=True)

    End Sub

    Private Sub ListViewPrm_AfterLabelEdit(sender As Object, e As Windows.Forms.LabelEditEventArgs) _
            Handles ListViewPrm.AfterLabelEdit

        If Not m_bDejaConfigure Then Exit Sub
        Dim bTirageAleatoire, bInitialiserFond As Boolean
        Dim sValeur$
        ' Cela ce produit lorsque le contenu du label n'a pas changé !
        If e.Label Is Nothing Then e.CancelEdit = True : Exit Sub
        sValeur = e.Label
        If Not bTraiterModifTexte(e.Item, sValeur,
            bTirageAleatoire, bInitialiserFond) Then _
                e.CancelEdit = True : Exit Sub

        ' Correction éventuelle des arrondis : 
        '  ne marche pas dans cet evenement
        'Me.ListViewPrm.Items.Item(e.Item).Text = sValeur

        MAJAnimation(bTirageAleatoire, bInitialiserFond,
            bControlerPrm:=True)
        m_bMajListPrm = True

    End Sub

    Private Function bTraiterModifBooleen(iIndexPrm%, CSEtat As CheckState,
            ByRef bTirageAleatoire As Boolean, ByRef bInitialiserFond As Boolean) As Boolean

        Dim bVal As Boolean = CBool(CSEtat)

        bTraiterModifBooleen = True
        Select Case iIndexPrm
            Case iDegreRacine_bRnd
                m_gravity.m_prm.bDegreRacine_bRnd = bVal : bTirageAleatoire = True
            Case iDegreRacine2_bRnd
                m_gravity.m_prm.bDegreRacine2_bRnd = bVal : bTirageAleatoire = True

            Case ibMasseSym
                m_gravity.m_prm.bMasseSym = bVal : bTirageAleatoire = True
                If Not bVal Then m_gravity.m_prm.bMasseSym_bRnd = False
            Case ibMasseSym_bRnd
                m_gravity.m_prm.bMasseSym_bRnd = bVal : bTirageAleatoire = True
                If bVal Then m_gravity.m_prm.bMasseSym = True

            Case ib3D
                m_gravity.m_prm.b3D = bVal : bTirageAleatoire = True
                If Not bVal Then m_gravity.m_prm.b3D_bRnd = False
            Case ib3D_bRnd : m_gravity.m_prm.b3D_bRnd = bVal : bTirageAleatoire = True

            Case ib3D_bPlanetesAxeV
                m_gravity.m_prm.b3D_bPlanetesAxeV = bVal : bTirageAleatoire = True
                If Not bVal Then m_gravity.m_prm.b3D_bPlanetesAxeV_bRnd = False
            Case ib3D_bPlanetesAxeV_bRnd
                m_gravity.m_prm.b3D_bPlanetesAxeV_bRnd = bVal : bTirageAleatoire = True
                If bVal Then m_gravity.m_prm.b3D_bPlanetesAxeV = True

            Case ibChocs
                m_gravity.m_prm.bChocs = bVal
                If Not bVal Then m_gravity.m_prm.bChocs_bRnd = False
            Case ibChocs_bRnd
                m_gravity.m_prm.bChocs_bRnd = bVal : bTirageAleatoire = True
                If bVal Then m_gravity.m_prm.bChocs = True

            Case iFiltreFichiersImgSprite
                m_gravity.m_prm.bCercle = (CSEtat = CheckState.Unchecked)
                ' Pour initialiser les images des sprites
                If Not m_gravity.m_prm.bCercle And m_gravity.m_bMembCercle Then _
                bTirageAleatoire = True
            Case iTxtBanniere : m_prmE.bAffichageBanniere = bVal

            Case iFiltreFichiersImgFond
                m_gravity.m_prm.bImageFond = bVal : bInitialiserFond = True
                If bVal Then m_gravity.m_prm.bFondUni = False : m_gravity.m_prm.bFondDegrade = False

            Case ibMAJToutLEcran : m_prmE.bMAJToutLEcran = bVal
            Case ibMAJGroupeSprites : m_gravity.m_prm.bMAJGroupeSprites = bVal
            Case ibNePasInitFond
                m_gravity.m_prm.bNePasInitFond = bVal : bInitialiserFond = True
            Case ibNePasBufferiserGr
                m_prmE.bNePasBufferiserGr = bVal : bInitialiserFond = True

            Case ibFondUni
                m_gravity.m_prm.bFondUni = bVal
                If bVal Then m_gravity.m_prm.bImageFond = False : m_gravity.m_prm.bFondDegrade = False
            Case ibFondDegrade
                m_gravity.m_prm.bFondDegrade = bVal
                If bVal Then m_gravity.m_prm.bImageFond = False : m_gravity.m_prm.bFondUni = False

            Case ibPauseAnimation : m_gravity.m_prm.bPauseAnimation = bVal

            Case Else
                bTraiterModifBooleen = False
        End Select

        If bTirageAleatoire Then bInitialiserFond = True

    End Function

    Private Function bTraiterModifTexte(iIndexPrm%,
            ByRef sValeur$, ByRef bTirageAleatoire As Boolean,
            ByRef bInitialiserFond As Boolean) As Boolean

        bTraiterModifTexte = True
        Dim bContinuer As Boolean
        Select Case iIndexPrm ' Saisie d'un texte
            Case iTxtBanniere
                If sValeur = "" Then
                    m_prmE.bAffichageBanniere = False
                Else
                    m_prmE.sTxtBanniere = sValeur
                End If

            Case iFiltreFichiersImgFond
                m_gravity.m_prm.sFiltreFichiersImgFond = sValeur : bInitialiserFond = True
            Case iFiltreFichiersImgSprite
                m_gravity.m_prm.sFiltreFichiersImgSprite = sValeur : bTirageAleatoire = True
            Case Else
                bContinuer = True
        End Select
        If bTirageAleatoire Then bInitialiserFond = True
        If Not bContinuer Then Exit Function

        ' Saisie d'un entier
        Dim iVal%
        Try
            'iVal = Integer.Parse(sValeur, Globalization.NumberStyles.Integer)
            iVal = CInt(sValeur) ' Identique dans ce cas
        Catch
            bTraiterModifTexte = False
            Exit Function
        End Try

        sValeur = CStr(iVal) ' Application de l'éventuel arrondi au label

        Select Case iIndexPrm
            Case iDegreRacine
                m_gravity.m_prm.iDegreRacine = CInt(sValeur) : bTirageAleatoire = True
            Case iDegreRacineRndMax
                m_gravity.m_prm.iDegreRacineRndMax = CInt(sValeur) : bTirageAleatoire = True
            Case iDegreRacine2
                m_gravity.m_prm.iDegreRacine2 = CInt(sValeur) : bTirageAleatoire = True
            Case iDegreRacine2RndMax
                m_gravity.m_prm.iDegreRacine2RndMax = CInt(sValeur) : bTirageAleatoire = True
            Case iForceGravitation : m_gravity.m_prm.rForceGravitation = CDec(sValeur)
            Case iDelaiMiliSec : m_prmE.iDelaiMiliSec = CInt(sValeur)
            Case iTempsMaxScenarioSec : m_prmE.iTempsMaxScenarioSec = CInt(sValeur)
            Case ib3D_iNbPlanetesMaxAxeV
                m_gravity.m_prm.b3D_iNbPlanetesMaxAxeV = CInt(sValeur) : bTirageAleatoire = True
            Case Else
                bTraiterModifTexte = False
        End Select

        If bTirageAleatoire Then bInitialiserFond = True

    End Function

    Public Sub ControlerParametres()

        m_gravity.ControlerParametres()

        If m_prmE.iDelaiMiliSec < 0 Then m_prmE.iDelaiMiliSec = 0
        If m_prmE.iDelaiMiliSec > 1000 Then m_prmE.iDelaiMiliSec = 1000
        m_prmE.bBoucleAnimation = False
        If m_prmE.iDelaiMiliSec = 0 Then m_prmE.bBoucleAnimation = True

        If m_prmE.iTempsMaxScenarioSec < 1 Then m_prmE.iTempsMaxScenarioSec = 1
        If m_prmE.iTempsMaxScenarioSec > 1000 Then m_prmE.iTempsMaxScenarioSec = 1000

    End Sub

#End Region

#Region "Traitements"

    Private Sub MAJAnimation(bTirageAleatoire As Boolean, bInitialiserFond As Boolean,
            bControlerPrm As Boolean)

        If bControlerPrm Then
            ControlerParametres()
            InitialiserTimer()
        End If

        If bInitialiserFond Then

            Dim bVal As Boolean = Not m_prmE.bNePasBufferiserGr
            SetStyle(ControlStyles.DoubleBuffer, bVal)
            SetStyle(ControlStyles.AllPaintingInWmPaint, bVal)
            SetStyle(ControlStyles.UserPaint, bVal)

            m_gravity.InitialiserImageFond(Me.ClientSize)
            If Not m_gravity.m_prm.bNePasInitFond Then
                If m_prmE.bNePasBufferiserGr Then
                    m_bFondInitialise = False
                Else
                    Me.Invalidate()
                End If
            End If
        End If

        If bTirageAleatoire Then m_gravity.TirageAleatoire()
        If bTirageAleatoire And SimulteurGravite.bDebugPosEtVitInitiales Then _
            Me.ListViewPrm.Items.Item(ibPauseAnimation).Checked = True

        m_rMemDateDepartAnimation = DateAndTime.Timer

    End Sub

    Private Sub MAJListePrm()

        ' Correction éventuelle des arrondis et autres incohérences
        With Me.ListViewPrm.Items
            .Item(iDegreRacine).Text = CStr(m_gravity.m_prm.iDegreRacine)
            .Item(iDegreRacine).Checked = False
            .Item(iDegreRacine_bRnd).Text = ""
            .Item(iDegreRacine_bRnd).Checked = m_gravity.m_prm.bDegreRacine_bRnd
            .Item(iDegreRacineRndMax).Text = CStr(m_gravity.m_prm.iDegreRacineRndMax)
            .Item(iDegreRacineRndMax).Checked = False
            .Item(iDegreRacine2).Text = CStr(m_gravity.m_prm.iDegreRacine2)
            .Item(iDegreRacine2).Checked = False
            .Item(iDegreRacine2_bRnd).Text = ""
            .Item(iDegreRacine2_bRnd).Checked = m_gravity.m_prm.bDegreRacine2_bRnd
            .Item(iDegreRacine2RndMax).Text = CStr(m_gravity.m_prm.iDegreRacine2RndMax)
            .Item(iDegreRacine2RndMax).Checked = False
            .Item(ibMasseSym).Text = ""
            .Item(ibMasseSym).Checked = m_gravity.m_prm.bMasseSym
            .Item(ibMasseSym_bRnd).Text = ""
            .Item(ibMasseSym_bRnd).Checked = m_gravity.m_prm.bMasseSym_bRnd
            .Item(ib3D).Text = ""
            .Item(ib3D).Checked = m_gravity.m_prm.b3D

            .Item(ib3D_bRnd).Text = ""
            .Item(ib3D_bRnd).Checked = m_gravity.m_prm.b3D_bRnd

            .Item(ib3D_bPlanetesAxeV).Text = ""
            .Item(ib3D_bPlanetesAxeV).Checked = m_gravity.m_prm.b3D_bPlanetesAxeV
            .Item(ib3D_bPlanetesAxeV_bRnd).Text = ""
            .Item(ib3D_bPlanetesAxeV_bRnd).Checked = m_gravity.m_prm.b3D_bPlanetesAxeV_bRnd
            .Item(ib3D_iNbPlanetesMaxAxeV).Text = CStr(m_gravity.m_prm.b3D_iNbPlanetesMaxAxeV)
            .Item(ib3D_iNbPlanetesMaxAxeV).Checked = False

            .Item(ibChocs).Text = ""
            .Item(ibChocs).Checked = m_gravity.m_prm.bChocs
            .Item(ibChocs_bRnd).Text = ""
            .Item(ibChocs_bRnd).Checked = m_gravity.m_prm.bChocs_bRnd
            .Item(iForceGravitation).Text = CStr(m_gravity.m_prm.rForceGravitation)
            .Item(iForceGravitation).Checked = False
            .Item(iDelaiMiliSec).Text = CStr(m_prmE.iDelaiMiliSec)
            .Item(iDelaiMiliSec).Checked = False
            .Item(iTempsMaxScenarioSec).Text = CStr(m_prmE.iTempsMaxScenarioSec)
            .Item(iTempsMaxScenarioSec).Checked = False
            .Item(ibMAJToutLEcran).Text = ""
            .Item(ibMAJToutLEcran).Checked = m_prmE.bMAJToutLEcran
            .Item(ibMAJGroupeSprites).Text = ""
            .Item(ibMAJGroupeSprites).Checked = m_gravity.m_prm.bMAJGroupeSprites
            .Item(ibNePasInitFond).Text = ""
            .Item(ibNePasInitFond).Checked = m_gravity.m_prm.bNePasInitFond
            .Item(ibNePasBufferiserGr).Text = ""
            .Item(ibNePasBufferiserGr).Checked = m_prmE.bNePasBufferiserGr
            .Item(ibFondUni).Text = ""
            .Item(ibFondUni).Checked = m_gravity.m_prm.bFondUni
            .Item(ibFondDegrade).Text = ""
            .Item(ibFondDegrade).Checked = m_gravity.m_prm.bFondDegrade
            .Item(ibPauseAnimation).Text = ""
            .Item(ibPauseAnimation).Checked = m_gravity.m_prm.bPauseAnimation

            If m_gravity.m_prm.bFondUni OrElse
           m_gravity.m_prm.bFondDegrade Then .Item(iFiltreFichiersImgFond).Checked = False

        End With
        m_bMajListPrm = False
        Me.ListViewPrm.Refresh() ' Utile lorsque le délai des frames diminue

    End Sub

    Private Sub SauverConfig()

        My.Settings.frmGravityNetPos = Me.Location
        My.Settings.frmGravityNetTaille = Me.Size

        With Me.ListViewPrm.Items

            My.Settings.TxtBanniere = .Item(iTxtBanniere).Text
            My.Settings.bAffichageBanniere = .Item(iTxtBanniere).Checked

            My.Settings.FiltreFichiersImgSprite = .Item(iFiltreFichiersImgSprite).Text
            My.Settings.FiltreFichiersImgFond = .Item(iFiltreFichiersImgFond).Text

            ' Ne pas sauver cette option, elle n'est pas proposée
            'My.Settings.bCercle = .Item(ibCercle).Checked

            My.Settings.DegreRacine = iConv(
            .Item(iDegreRacine).Text, iDegreRacineDef)
            '.Item(iDegreRacine).Text = CStr(m_gravity.m_prm.iDegreRacine)
            '.Item(iDegreRacine).Checked = False

            My.Settings.DegreRacine_bRnd = .Item(iDegreRacine_bRnd).Checked
            '.Item(iDegreRacine_bRnd).Checked = m_gravity.m_prm.bDegreRacine_bRnd

            My.Settings.DegreRacineRndMax = iConv(
            .Item(iDegreRacineRndMax).Text, iDegreRacineMaxDef)
            '.Item(iDegreRacineRndMax).Text = CStr(m_gravity.m_prm.iDegreRacineRndMax)
            '.Item(iDegreRacineRndMax).Checked = False

            My.Settings.DegreRacine2 = iConv(
            .Item(iDegreRacine2).Text, iDegreRacineDef)
            '.Item(iDegreRacine2).Text = CStr(m_gravity.m_prm.iDegreRacine2)
            '.Item(iDegreRacine2).Checked = False

            My.Settings.DegreRacine2_bRnd = .Item(iDegreRacine2_bRnd).Checked
            '.Item(iDegreRacine2_bRnd).Checked = m_gravity.m_prm.bDegreRacine2_bRnd

            My.Settings.DegreRacine2RndMax = iConv(
            .Item(iDegreRacine2RndMax).Text, iDegreRacineMaxDef)
            '.Item(iDegreRacine2RndMax).Text = CStr(m_gravity.m_prm.iDegreRacine2RndMax)
            '.Item(iDegreRacine2RndMax).Checked = False

            My.Settings.bMasseSym = .Item(ibMasseSym).Checked
            '.Item(ibMasseSym).Checked = m_gravity.m_prm.bMasseSym

            My.Settings.bMasseSym_bRnd = .Item(ibMasseSym_bRnd).Checked
            '.Item(ibMasseSym_bRnd).Checked = m_gravity.m_prm.bMasseSym_bRnd

            My.Settings.b3D = .Item(ib3D).Checked
            '.Item(ib3D).Checked = m_gravity.m_prm.b3D

            My.Settings.b3D_bRnd = .Item(ib3D_bRnd).Checked
            '.Item(ib3D_bRnd).Checked = m_gravity.m_prm.b3D_bRnd

            My.Settings.b3D_bPlanetesAxeV = .Item(ib3D_bPlanetesAxeV).Checked
            My.Settings.b3D_bPlanetesAxeV_bRnd = .Item(ib3D_bPlanetesAxeV_bRnd).Checked
            My.Settings.b3D_iNbPlanetesMaxAxeV = iConv(
            .Item(ib3D_iNbPlanetesMaxAxeV).Text, 1)

            My.Settings.bChocs = .Item(ibChocs).Checked
            '.Item(ibChocs).Checked = m_gravity.m_prm.bChocs

            My.Settings.bChocs_bRnd = .Item(ibChocs_bRnd).Checked
            '.Item(ibChocs_bRnd).Checked = m_gravity.m_prm.bChocs_bRnd

            My.Settings.ForceGravitation = iConv(
            .Item(iForceGravitation).Text, iForceGravitationDef)
            '.Item(iForceGravitation).Checked = False

            My.Settings.DelaiMiliSec = iConv(
            .Item(iDelaiMiliSec).Text, iDelaiMiliSecDef)
            '.Item(iDelaiMiliSec).Text = CStr(m_prm.iDelaiMiliSec)
            '.Item(iDelaiMiliSec).Checked = False

            My.Settings.TempsMaxScenarioSec = iConv(
            .Item(iTempsMaxScenarioSec).Text, iTempsMaxScenarioSecDef)
            '.Item(iTempsMaxScenarioSec).Text = CStr(m_prm.iTempsMaxScenarioSec)
            '.Item(iTempsMaxScenarioSec).Checked = False

            My.Settings.bMAJToutLEcran = .Item(ibMAJToutLEcran).Checked
            '.Item(ibMAJToutLEcran).Checked = m_prm.bMAJToutLEcran

            My.Settings.bMAJGroupeSprites = .Item(ibMAJGroupeSprites).Checked
            '.Item(ibMAJGroupeSprites).Checked = m_gravity.m_prm.bMAJGroupeSprites

            My.Settings.bNePasInitFond = .Item(ibNePasInitFond).Checked
            '.Item(ibNePasInitFond).Checked = m_gravity.m_prm.bNePasInitFond

            My.Settings.bNePasBufferiserGr = .Item(ibNePasBufferiserGr).Checked
            '.Item(ibNePasBufferiserGr).Checked = m_prm.bNePasBufferiserGr

            My.Settings.bFondUni = .Item(ibFondUni).Checked
            '.Item(ibFondUni).Checked = m_gravity.m_prm.bFondUni

            My.Settings.bFondDegrade = .Item(ibFondDegrade).Checked
            '.Item(ibFondDegrade).Checked = m_gravity.m_prm.bFondDegrade

            ' Ne pas sauver cette option
            '.Item(ibPauseAnimation).Checked = m_gravity.m_prm.bPauseAnimation

        End With

        My.Settings.Save()

    End Sub

    Private Sub TimerAnimation_Tick(sender As Object, e As EventArgs) Handles TimerAnimation.Tick

        If m_prmE.bBoucleAnimation Then
            TimerAnimation.Enabled = False
            BoucleAnimation()
            Exit Sub
        End If

        Animer()

    End Sub

    Private Sub BoucleAnimation()

        If m_bBoucleAnimationEnCours Then Exit Sub
        m_bBoucleAnimationEnCours = True
        Do
            Animer()
            Application.DoEvents()
        Loop While Not m_bQuitterBoucleAnimation

    End Sub

    Private Sub Animer()

        If glb_bModeConfiguration Then
            If m_bMajListPrm Then MAJListePrm()

            ' Calcul du nombre de Frames Par Seconde
            Static iNbFrames%
            Static rMemDate As Double
            Dim rFps As Double
            Dim rDate As Double
            Const iNbFramesCalculMoy% = 30
            iNbFrames = iNbFrames + 1
            If iNbFrames = iNbFramesCalculMoy Then
                rDate = DateAndTime.Timer
                If rDate <> rMemDate Then
                    rFps = iNbFramesCalculMoy / (rDate - rMemDate)
                    ' Ne pas effacer tout de suite l'explication du prm
                    If rMemDate <> 0 And rDate - glb_rDateMessageTitre > 5 Then _
                        Me.Text = m_sTitreAppli & " - Frames/s : " &
                            rFps.ToString("####.0") &
                            ", RAM : " & GC.GetTotalMemory(False) &
                            " octets utilisés, " &
                            CInt(m_prmE.iTempsMaxScenarioSec -
                            (DateAndTime.Timer - m_rMemDateDepartAnimation))
                    rMemDate = rDate
                    iNbFrames = 0
                End If
            End If
        End If

        ' Utile seulement pour déplacer plusieurs contrôles de la frm 
        '  d'un seul coup (LblBanniere par exemple), ce n'est pas utile 
        '  pour le tracé dans la frm
        'Me.SuspendLayout()

        If m_gravity.m_bToutesPlanetesHorsEcran Or
            DateAndTime.Timer - m_rMemDateDepartAnimation >
                m_prmE.iTempsMaxScenarioSec Then
            MAJAnimation(bTirageAleatoire:=True, bInitialiserFond:=True,
                bControlerPrm:=False)
        End If
        m_gravity.SimulerGravite()

        If m_prmE.bNePasBufferiserGr Then
            ' Si on ne bufférise pas le graphisme, 
            '  on trace directement dans la form
            If m_grFrm Is Nothing Then m_grFrm = Me.CreateGraphics
            If Not m_bFondInitialise Then
                If Not m_gravity.m_prm.bNePasInitFond Then _
                    m_gravity.DessinerFond(m_grFrm)
                m_bFondInitialise = True
            End If
            m_gravity.Dessiner(m_grFrm, m_prmE.bNePasBufferiserGr)
        Else
            If m_prmE.bMAJToutLEcran Then
                Me.Invalidate()
            Else
                If m_gravity.m_prm.bMAJGroupeSprites Then
                    Me.Invalidate(m_gravity.m_rectMAJGroupeSprites)
                Else
                    Dim i%
                    For i = 0 To m_gravity.m_iNbSprites - 1
                        Me.Invalidate(m_gravity.m_aSprites(i).m_rectMAJ)
                    Next i
                End If
            End If
        End If

        If m_prmE.bAffichageBanniere Then
            If LblBanniere.Text <> m_prmE.sTxtBanniere Then
                LblBanniere.Text = m_prmE.sTxtBanniere
                LblBanniere.Height = LblBanniere.Font.Height
                LblBanniere.Width = CInt(LblBanniere.Text.Length * LblBanniere.Font.Size)
            End If
            LblBanniere.Visible = True
        Else
            LblBanniere.Visible = False
        End If

        If Not m_prmE.bAffichageBanniere Then GoTo Fin

        ' Gestion de la bannière
        ' ----------------------

        Dim pt As Point
        pt.X = m_rectEcran.Width - m_iPosBanniereX
        pt.Y = LblBanniere.Location.Y
        '//Increment the label distance based on the speed set by the user.
        m_iPosBanniereX += m_iVitesseBanniere
        '//If the label is offscreen, then we want to reposition it to the right.
        If pt.X <= -LblBanniere.Width Then
            m_iPosBanniereX = 0 '//Reset the distance to 0.
            If pt.Y = 0 Then
                '//If the label is at the top, move it to the middle.
                pt.Y = m_rectEcran.Height \ 2
            ElseIf pt.Y = CInt(m_rectEcran.Height / 2) Then
                '// If label is in the middle of the screen move it to the bottom.
                pt.Y = m_rectEcran.Height - LblBanniere.Height
            Else
                pt.Y = 0 '//Move the label back to the top.
            End If
        End If
        LblBanniere.Location = pt
        ' Pour éviter d'afficher la bannière avant le départ
        If Not LblBanniere.Visible Then LblBanniere.Visible = True

Fin:
        'Me.ResumeLayout(False)

    End Sub

    Protected Overrides Sub OnPaintBackground(pevent As PaintEventArgs)
        If Not m_gravity.m_prm.bNePasInitFond Then MyBase.OnPaintBackground(pevent)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)

        ' Appel de la fonction de base du tracé
        MyBase.OnPaint(e)

        ' Si on trace directement depuis l'animation, c'est déjà fait
        If m_prmE.bNePasBufferiserGr Then Exit Sub
        m_gravity.Dessiner(e.Graphics, bNePasBufferiserGr:=False)

    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)

        '//Determine if the mouse cursor position has been stored previously.
        If m_ptPosSouris.X = 0 And m_ptPosSouris.Y = 0 Then
            '//Store the mouse cursor coordinates.
            m_ptPosSouris.X = e.X
            m_ptPosSouris.Y = e.Y
            Exit Sub
        ElseIf e.X <> m_ptPosSouris.X Or e.Y <> m_ptPosSouris.Y Then
            '//Has the mouse cursor moved since the screen saver was started? 
            If Not glb_bModeConfiguration Then Quitter()
        End If

    End Sub

    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        If Not glb_bModeConfiguration Then Quitter()
    End Sub

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        If Not glb_bModeConfiguration Then Quitter()
    End Sub

    Private Sub FrmGravityNet_Click(sender As Object, e As EventArgs) Handles MyBase.Click
        If glb_bModeConfiguration Then _
            MAJAnimation(bTirageAleatoire:=True, bInitialiserFond:=True,
                bControlerPrm:=False)
    End Sub

    Protected Overrides Sub OnSizeChanged(e As System.EventArgs)
        InitialiserTailleEcran()
    End Sub

    Private Sub InitialiserTailleEcran()

        If m_prmE.bNePasBufferiserGr Then
            If Not (m_grFrm Is Nothing) Then m_grFrm.Dispose()
            m_grFrm = Me.CreateGraphics
        End If
        m_gravity.InitialiserTailleEcran(Me.ClientSize)
        m_rectEcran = New Rectangle(0, 0, Me.ClientSize.Width, Me.ClientSize.Height)
        m_iPosBanniereX = m_rectEcran.Width \ 2
        LblBanniere.Location = New Point(m_rectEcran.Width \ 2,
            m_rectEcran.Height \ 2)
        If Not m_gravity.m_prm.bNePasInitFond Then Me.Invalidate()

    End Sub

    Private Sub ArreterAnimation()
        Cursor.Show()
        TimerAnimation.Enabled = False
        m_bQuitterBoucleAnimation = True
    End Sub

    Private Sub Quitter()
        ArreterAnimation()
        Me.Close()
    End Sub

    Private Sub FrmVBNetScreenSaver_Closing(sender As Object,
            e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        ArreterAnimation()
        If glb_bModeConfiguration Then SauverConfig()
    End Sub

#End Region

End Class