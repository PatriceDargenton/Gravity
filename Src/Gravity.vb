
Imports System.IO ' Pour Path, FileInfo
Imports System.Drawing.Drawing2D ' Pour LinearGradientBrush

Public Class SimulteurGravite : Implements IDisposable

#Region "Constantes"

    Public Const bDebugPosEtVitInitiales As Boolean = False

    ' Booléen pour annuler la gravité (debug choc)
    Private Const bGravite As Boolean = True

    'Private Const bTestOrb3D As Boolean = False

    Private Const bDebugChoc As Boolean = False
    Private Const bDebugRectMAJ As Boolean = False
    Private Const bChocsMoyennes As Boolean = False

    ' Pour regler la vitesse initiale
    Private Const rRapportVitesse As Decimal = 5
    Private Const iDegreRacineMax% = 6
    Private Const iDegreRacineMin% = 1

#End Region

#Region "Déclarations"

    ' Pour savoir s'il faut initialiser les img de sprites
    Public m_bMembCercle As Boolean
    Public m_bToutesPlanetesHorsEcran As Boolean
    Public m_szTailleFenetre As Size
    Public m_rectMAJGroupeSprites As Rectangle ' Rectangle de mise à jour
    Public m_iNbSprites% = 0
    Public m_aSprites() As Sprite

    Public m_prm As TParametres
    ' Etat effectif après tirage (ne pas changer la case à cocher selon le tirage)
    Private m_b3D, m_bChocs, m_b3D_bPlanetesAxeV As Boolean

    Private m_aff As TAffichage
    Private m_iNbPtsTot%
    Private m_pt() As TPoint
    Private m_aCoordZ!()
    Private m_aIndexCoordZ%()

    Private m_bSystemeInitialise As Boolean

    ' Booléen pour indiquer si l'écran de veille est configuré
    Private m_bImgFondInitialisee As Boolean = False
    Private m_bImageFondTrouve As Boolean = False
    Private m_imgFond As Bitmap ' Bitmap au lieu d'Image pour GetHbitmap
    Private m_rectEcran As Rectangle
    Private m_rectImgFond As Rectangle
    Private m_lgbFondDegrade As LinearGradientBrush

    Private m_frm As Form

#End Region

#Region "Structures"

    ' Structure pour gérer les coordonnées des planètes
    Private Structure TPoint
        Dim rX, rY, rZ As Decimal    ' Positions
        Dim rM As Decimal            ' Masse
        Dim rVx, rVy, rVz As Decimal ' Vitesses
        Dim rAx, rAy, rAz As Decimal ' Accélérations

        Dim iNbChocs%
        Dim rSomX, rSomY As Decimal  ' Test infructueux
        Dim rSomVX, rSomVY As Decimal
        ' Positions corrigées à l'instant précis du choc
        Dim rXC, rYC, rZC As Decimal
        Dim rAngleChoc As Decimal
        Dim rDx, rDy As Decimal
    End Structure

    Private Structure TAffichage
        Dim rMaxx As Decimal ' Amplitude horizontale   de tracé
        Dim rMaxy As Decimal ' Amplitude verticale     de tracé
        Dim rMaxz As Decimal ' Amplitude en profondeur du tracé
        Dim rMaxH As Decimal ' Taille horizontale de l'écran
        Dim rMaxV As Decimal ' Taille verticale   de l'écran
        Dim rZoom As Decimal ' Pour zoomer le dessin au besoin
    End Structure

    ' Structure des propriétés d'une planète pour gérer les symétries
    Private Structure TProprietesPlanete
        Dim rMasse As Decimal ' masse de la planète
        Dim iNumImg% ' n° de l'image de la planète
        Dim rSpin As Decimal ' Spin (rotation sur elle-même) de la planète
    End Structure

    ' Structure pour le calcul initial des orbites (vitesses et positions)
    '  des planètes des systèmes principal et secondaire
    Private Structure TSysteme
        ' Angle pour le calcul initial des positions et vitesses
        Dim rAngle As Decimal
        Dim iNbPts% ' Nombre de planètes du système
        Dim iNbPtsMin% ' Nombre min. de planètes du système

        ' Degré de la racine unitaire complexe du système
        ' = Nombre de planètes (ou sous-sytèmes) du système
        Dim iDegreRacine%

        Dim iDegreRacineMin% ' Degré min.  pour le tirage aléatoire
        Dim iDegreRacineMax% ' Degré max.  pour le tirage aléatoire
        Dim iDegreRacineFin% ' Degré final pour le tirage aléatoire
        Dim rAmplitPos As Decimal    ' Amplitude          de la position 
        ' (=rayon de l'orbite)
        Dim rAmplitPosMin As Decimal ' Amplitude minimale de la position
        Dim rAmplitPosMax As Decimal ' Amplitude maximale de la position
        Dim rAmplitVit As Decimal    ' Amplitude          de la vitesse initiale
        Dim rAmplitVitMin As Decimal ' Amplitude minimale de la vitesse initiale
        Dim rAmplitVitMax As Decimal ' Amplitude maximale de la vitesse initiale

        Dim aPlanete() As TProprietesPlanete
        Dim aPlaneteSym() As TProprietesPlanete

        Dim rAxz As Decimal ' Angles pour le calcul initial des positions
        Dim rAxy As Decimal '  et vitesses dans le cas 3D.
        Dim rAxyP As Decimal
    End Structure

    ' Structures pour les paramètres de gravity
    Structure TParametres

        Dim bSystemeFixe As Boolean
        Dim bMasseSym As Boolean
        Dim bMasseSym_bRnd As Boolean
        Dim iDegreRacine% ' Voir la structure TSysteme
        Dim iDegreRacineRndMax%
        Dim iDegreRacine2%
        Dim iDegreRacine2RndMax%
        Dim bDegreRacine_bRnd As Boolean ' Choix au hasard du degré
        Dim bDegreRacine2_bRnd As Boolean
        Dim bChocs As Boolean
        Dim bChocs_bRnd As Boolean

        Dim b3D As Boolean
        Dim b3D_bRnd As Boolean

        Dim b3D_bPlanetesAxeV As Boolean

        ' A la fois pour Oui/Non et aussi pour le nombre de planètes à ajouter
        Dim b3D_bPlanetesAxeV_bRnd As Boolean

        Dim b3D_iNbPlanetesMaxAxeV%

        Dim bCercle As Boolean
        Dim rForceGravitation As Decimal

        Dim bMAJGroupeSprites As Boolean
        Dim bImageFond As Boolean
        Dim bNePasInitFond As Boolean
        Dim bNePasDecentrerImgFond As Boolean
        ' On gagne 5 à 10 % en vitesse
        Dim bNePasAgrandirImgFond As Boolean
        Dim bClipping As Boolean
        Dim bFondUni As Boolean
        Dim bFondDegrade As Boolean

        Dim bPauseAnimation As Boolean
        Dim sFiltreFichiersImgSprite$
        Dim sFiltreFichiersImgFond$

    End Structure

#End Region

#Region "Initialisations"

    Public Sub New(ByRef FrmGravityNet As Form)

        m_frm = FrmGravityNet ' Pour afficher un msg dans la barre de titre

        m_bSystemeInitialise = False
        m_bToutesPlanetesHorsEcran = True

        m_prm.bImageFond = My.Settings.bImageFond
        m_prm.bMAJGroupeSprites = My.Settings.bMAJGroupeSprites
        m_prm.bNePasInitFond = My.Settings.bNePasInitFond
        m_prm.bFondDegrade = My.Settings.bFondDegrade
        m_prm.bFondUni = My.Settings.bFondUni

        m_prm.bSystemeFixe = True ' Pour centrer l'ensemble des planètes
        m_prm.bMasseSym = My.Settings.bMasseSym
        m_prm.bMasseSym_bRnd = My.Settings.bMasseSym_bRnd
        m_prm.bCercle = My.Settings.bCercle
        m_prm.iDegreRacine = My.Settings.DegreRacine
        m_prm.iDegreRacineRndMax = My.Settings.DegreRacineRndMax
        m_prm.bDegreRacine_bRnd = My.Settings.DegreRacine_bRnd
        m_prm.iDegreRacine2 = My.Settings.DegreRacine2
        m_prm.iDegreRacine2RndMax = My.Settings.DegreRacine2RndMax
        m_prm.bDegreRacine2_bRnd = My.Settings.DegreRacine2_bRnd
        m_prm.rForceGravitation = My.Settings.ForceGravitation
        m_prm.b3D = My.Settings.b3D
        m_prm.b3D_bRnd = My.Settings.b3D_bRnd
        m_prm.bChocs = My.Settings.bChocs
        m_prm.bChocs_bRnd = My.Settings.bChocs_bRnd

        m_prm.b3D_bPlanetesAxeV = My.Settings.b3D_bPlanetesAxeV
        m_prm.b3D_bPlanetesAxeV_bRnd = My.Settings.b3D_bPlanetesAxeV_bRnd
        m_prm.b3D_iNbPlanetesMaxAxeV = My.Settings.b3D_iNbPlanetesMaxAxeV

        If bDebugChoc Then
            m_prm.bMasseSym = True
            m_prm.bMasseSym_bRnd = False
            m_prm.bCercle = True
            m_prm.bDegreRacine_bRnd = False
            m_prm.iDegreRacine2 = 1
            m_prm.bDegreRacine2_bRnd = False
        End If

        m_prm.sFiltreFichiersImgFond = My.Settings.FiltreFichiersImgFond
        m_prm.sFiltreFichiersImgSprite = My.Settings.FiltreFichiersImgSprite

        ControlerParametres()

    End Sub

    Protected Overridable Overloads Sub Dispose(disposing As Boolean)
        If disposing Then ' Dispose managed resources
            m_lgbFondDegrade.Dispose()
        End If
        ' Free native resources
    End Sub
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Public Sub ControlerParametres()

        'If m_prm.bFondUni OrElse m_prm.bFondDegrade Then m_prm.bImageFond = False
        If m_prm.bImageFond Then m_prm.bFondUni = False : m_prm.bFondDegrade = False
        If m_prm.bFondDegrade Then m_prm.bFondUni = False : m_prm.bImageFond = False
        If m_prm.bFondUni Then m_prm.bFondDegrade = False : m_prm.bImageFond = False

        If m_prm.iDegreRacineRndMax <= m_prm.iDegreRacine Then _
            m_prm.iDegreRacineRndMax = m_prm.iDegreRacine + 1
        If m_prm.iDegreRacineRndMax < iDegreRacineMin Then _
            m_prm.iDegreRacineRndMax = iDegreRacineMin
        If m_prm.iDegreRacineRndMax > iDegreRacineMax Then _
            m_prm.iDegreRacineRndMax = iDegreRacineMax

        If m_prm.iDegreRacine2RndMax <= m_prm.iDegreRacine2 Then _
            m_prm.iDegreRacine2RndMax = m_prm.iDegreRacine2 + 1
        If m_prm.iDegreRacine2RndMax < iDegreRacineMin Then _
            m_prm.iDegreRacine2RndMax = iDegreRacineMin
        If m_prm.iDegreRacine2RndMax > iDegreRacineMax Then _
            m_prm.iDegreRacine2RndMax = iDegreRacineMax

        If m_prm.iDegreRacine < iDegreRacineMin Then _
            m_prm.iDegreRacine = iDegreRacineMin
        If m_prm.iDegreRacine > iDegreRacineMax Then _
            m_prm.iDegreRacine = iDegreRacineMax

        If m_prm.iDegreRacine2 < iDegreRacineMin Then _
            m_prm.iDegreRacine2 = iDegreRacineMin
        If m_prm.iDegreRacine2 > iDegreRacineMax Then _
            m_prm.iDegreRacine2 = iDegreRacineMax

        If Not m_prm.bDegreRacine_bRnd And m_prm.iDegreRacine = 1 _
           And m_prm.iDegreRacine2 = 1 Then m_prm.iDegreRacine = 2

        If m_prm.rForceGravitation <= 0 Then m_prm.rForceGravitation = 0

        If m_prm.b3D_bRnd Then m_prm.b3D = True
        If m_prm.b3D_bPlanetesAxeV_bRnd Then m_prm.b3D_bPlanetesAxeV = True
        If m_prm.bChocs_bRnd Then m_prm.bChocs = True

        ' Si on active les chocs, on désactive la 3D, car je n'arrive 
        '  pas à faire coincider précisément l'instant du choc avec
        '  le graphisme (le système de projection est trop simpliste)
        If m_prm.bChocs And Not bDebugChoc Then m_prm.b3D = False : m_prm.b3D_bRnd = False

    End Sub

    Private Function iRandomiser%(iMin%, iMax%, Optional rRnd As Decimal = -1D)
        If iMin = iMax Then iRandomiser = iMin : Exit Function
        If rRnd = -1D Then rRnd = CDec(Rnd())
        iRandomiser = CInt(rRnd * (iMax - iMin)) + iMin
        If iRandomiser > iMax Then
            Stop
            iRandomiser = iMax
        End If
    End Function

    Private Function rRandomiser(rMin As Decimal, rMax As Decimal) As Decimal
        If rMin = rMax Then rRandomiser = rMin : Exit Function
        rRandomiser = CDec(Rnd() * (rMax - rMin)) + rMin
    End Function

    Public Sub TirageAleatoire()

        m_bSystemeInitialise = False

        If bDebugPosEtVitInitiales Then m_prm.bPauseAnimation = True

        ' Imbrication de deux systèmes en rotation
        Dim sys1, sys2 As TSysteme
        sys2.aPlaneteSym = Nothing ' Pour éviter Warning

        Dim k, i, j, l As Integer
        Dim rAmplitMasseMin, rAmplitMasseMax As Decimal
        Dim rSpinMaxDeg As Decimal
        Dim iNbPts2Sur2%
        Dim bMasseSym As Boolean

        ' Numéro de chaque constante du tirage
        '  pour faciliter une future sauvegarde de la session,
        '  il suffit de sauver chaque Rnd(0 à iNbRnd)
        '  avec un format de précision
        Const iRndDeg1% = 0
        'Const iRndPlanete1% = 1
        Const iRndAmplitOrb1% = 2
        Const iRndAmplitVit1% = 3
        Const iRndAngleDepart1% = 4

        Const iRndNbPts2% = 5
        'Const iRndPlanete2% = 6
        Const iRndbMasseSym2% = 7
        'Const iRndMasse2% = 8
        Const iRndAmplitOrb2% = 9
        Const iRndAmplitVit2% = 10
        Const iRndAngleDepart2% = 11

        Const iRndNumImgSym2% = 12
        Const iRndSpinImgSym2% = 13
        Const iRndb3D% = 14
        Const iRndbChocs% = 15
        Const iRndb3D_bPlanetesAxeV_bRnd% = 16

        Const iNbRnd% = 16
        Dim arRnd(iNbRnd) As Decimal

        Dim asFichiersImg$() = Nothing

        Dim iNbFichiersPlanetes% = 0
        Dim sFiltre$ = m_prm.sFiltreFichiersImgSprite
        If sFiltre = "" Then GoTo Nouveau_Tirage

        Dim sRepertoire$ = Application.StartupPath

        Dim sRepertoireImgSprite$ = sRepertoire

        Dim iPos% = sFiltre.IndexOf("\")
        If iPos > 0 Then
            sRepertoireImgSprite = sRepertoire & "\" &
                sFiltre.Substring(0, iPos)
            sFiltre = sFiltre.Substring(iPos + 1)
        End If

        If Not bDossierExiste(sRepertoireImgSprite) Then GoTo Nouveau_Tirage

        Try
            asFichiersImg = Directory.GetFiles(sRepertoireImgSprite, sFiltre) '"star_*.*") 
            iNbFichiersPlanetes = asFichiersImg.Length
        Catch
            iNbFichiersPlanetes = 0
        End Try


Nouveau_Tirage:
        m_bMembCercle = m_prm.bCercle
        Randomize() ' Initialise le générateur de nombres aléatoires.
        For i = 0 To iNbRnd
            arRnd(i) = CDec(Rnd())
        Next i

        m_bChocs = m_prm.bChocs
        If m_prm.bChocs_bRnd Then _
            m_bChocs = (arRnd(iRndbChocs) > 0.5) ' 1 chance sur 2

        m_b3D = m_prm.b3D
        If m_prm.b3D_bRnd Then _
            m_b3D = (arRnd(iRndb3D) > 0.5)

        ' Si on active les chocs, on désactive la 3D
        If m_bChocs And Not bDebugChoc Then m_b3D = False

        m_b3D_bPlanetesAxeV = m_prm.b3D_bPlanetesAxeV
        If m_b3D And m_prm.b3D_bPlanetesAxeV_bRnd Then _
            m_b3D_bPlanetesAxeV = (arRnd(iRndb3D_bPlanetesAxeV_bRnd) > 0.5)

        ' Définition des tailles d'écran
        m_aff.rMaxH = m_szTailleFenetre.Width
        m_aff.rMaxV = m_szTailleFenetre.Height
        m_aff.rMaxx = m_aff.rMaxH
        m_aff.rMaxy = m_aff.rMaxV
        If m_b3D Then 'Or bTestOrb3D Then
            ' Définition de la perspective 3D
            m_aff.rMaxx = 0.66D * m_aff.rMaxH
            m_aff.rMaxy = 0.66D * m_aff.rMaxV
            m_aff.rMaxz = 0.33D * m_aff.rMaxH
            If 0.33 * m_aff.rMaxV < m_aff.rMaxz Then _
                m_aff.rMaxz = 0.33D * m_aff.rMaxV
        End If
        m_aff.rZoom = m_aff.rMaxH / m_aff.rMaxx
        If m_aff.rMaxV / m_aff.rMaxy < m_aff.rZoom Then _
            m_aff.rZoom = m_aff.rMaxV / m_aff.rMaxy
        If m_b3D Then 'Or bTestOrb3D Then
            If m_aff.rMaxH / m_aff.rMaxz < m_aff.rZoom Then _
                m_aff.rZoom = m_aff.rMaxH / m_aff.rMaxz
            If m_aff.rMaxV / m_aff.rMaxz < m_aff.rZoom Then _
                m_aff.rZoom = m_aff.rMaxV / m_aff.rMaxz
        End If

        sys1.iDegreRacineMin = m_prm.iDegreRacine
        If m_prm.bDegreRacine_bRnd Then
            sys1.iDegreRacineMax = m_prm.iDegreRacineRndMax
        Else
            sys1.iDegreRacineMax = m_prm.iDegreRacine
        End If

        sys2.iNbPtsMin = m_prm.iDegreRacine2
        If m_prm.bDegreRacine2_bRnd Then
            sys2.iDegreRacineMax = m_prm.iDegreRacine2RndMax
        Else
            sys2.iDegreRacineMax = m_prm.iDegreRacine2
        End If

        rSpinMaxDeg = 7
        rAmplitMasseMax = 70 '150
        rAmplitMasseMin = 5 '10
        sys1.rAmplitVitMin = 0
        sys1.rAmplitVitMax = 1.5D
        sys2.rAmplitVitMin = 0
        sys2.rAmplitVitMax = 1

        ' Test d'une autre représentation 3D
        Const bTest3D As Boolean = False

        ' Angles des plans xz et xy du système 2
        If bTest3D Then
            sys2.rAxz = 0
            sys2.rAxy = CDec(Math.PI / 2)
            sys2.rAxyP = CDec(Math.PI / 2)
            rAmplitMasseMin = 20 '5
            sys1.rAmplitVitMin = 0.5D
            sys2.rAmplitVitMin = 0.5D
            sys1.rAmplitVitMax = sys1.rAmplitVitMin
            rAmplitMasseMax = rAmplitMasseMin
            sys2.rAmplitVitMax = sys2.rAmplitVitMin
        End If

        ' Détermination du nombre de planètes (=pt) de chaque système
        sys2.iNbPts = sys2.iNbPtsMin +
            CInt((sys2.iDegreRacineMax - sys2.iNbPtsMin) * arRnd(iRndNbPts2))

        ' S'il n'y a qu'un pt dans le système secondaire,
        '  pas de rotation autour d'un axe
        If sys2.iNbPts <= 1 Then sys2.rAmplitPos = 0 : sys2.rAmplitVit = 0
        sys1.iDegreRacineFin = sys1.iDegreRacineMin +
            CInt((sys1.iDegreRacineMax - sys1.iDegreRacineMin) * arRnd(iRndDeg1))
        sys1.iNbPts = sys1.iDegreRacineFin
        m_iNbPtsTot = sys2.iNbPts * sys1.iNbPts
        If m_iNbPtsTot = 1 Then GoTo Nouveau_Tirage

        ' Test Orbites 3D
        Dim planeteAxeZ As TSysteme
        planeteAxeZ.aPlanete = Nothing ' Pour éviter Warning

        Dim iNbPlanetesZ%

        'If bTestOrb3D Then
        If m_b3D And m_b3D_bPlanetesAxeV Then

            ' Ajout de planètes dans l'axe vertical 3D :
            '  Cela ne pertube pas l'équilibre du plan horizontal

            'Const iNbPlanetesZMax% = 5
            Dim iNbPlanetesZMax% = m_prm.b3D_iNbPlanetesMaxAxeV
            If m_prm.b3D_bPlanetesAxeV_bRnd Then
                iNbPlanetesZ = iRandomiser(0, iNbPlanetesZMax)
            Else
                iNbPlanetesZ = iNbPlanetesZMax
            End If

            ' ToDo : faire une fonction : code dupliqué
            For j = 0 To iNbPlanetesZ - 1
                m_iNbPtsTot += 1
                ReDim Preserve planeteAxeZ.aPlanete(j)
                planeteAxeZ.aPlanete(j).rMasse =
                    rRandomiser(rAmplitMasseMin, rAmplitMasseMax)
                If iNbFichiersPlanetes > 0 Then
                    planeteAxeZ.aPlanete(j).iNumImg =
                        iRandomiser(0, iNbFichiersPlanetes - 1)
                End If
                planeteAxeZ.aPlanete(j).rSpin = rRandomiser(-rSpinMaxDeg, rSpinMaxDeg)
            Next j

        End If

        ' Détermination des amplitudes max. et min. des orbites
        ' Relatif à maxx
        sys1.rAmplitPosMin = 0.1D
        ' 50% de l'espace Horiz. pour le rayon orbitale max. du système 1 :
        sys1.rAmplitPosMax = 0.5D
        ' L'amplitude du système 2 est relatif à l'amplitude du système primaire choisit
        '  sauf si celui ci est de degré 1 (c.a.d. système 2 seulement)
        If sys1.iDegreRacineFin = 1 Then
            sys2.rAmplitPosMin = 0.1D
            sys2.rAmplitPosMax = 0.5D
        Else
            sys2.rAmplitPosMin = sys1.rAmplitPosMin / 2
            sys2.rAmplitPosMax = sys1.rAmplitPosMax / 2
        End If

        sys1.rAmplitPos = m_aff.rMaxx * (sys1.rAmplitPosMin +
            (sys1.rAmplitPosMax - sys1.rAmplitPosMin) * arRnd(iRndAmplitOrb1))
        If sys1.iDegreRacineFin = 1 Then
            sys2.rAmplitPos = CDec(0.5 * m_aff.rMaxx * (sys2.rAmplitPosMin +
                (sys2.rAmplitPosMax - sys2.rAmplitPosMin) * arRnd(iRndAmplitOrb2)))
        Else
            sys2.rAmplitPos = m_aff.rMaxx * sys2.rAmplitPosMin +
                (sys2.rAmplitPosMax - sys2.rAmplitPosMin) *
                sys1.rAmplitPos * arRnd(iRndAmplitOrb2)
        End If

        ' Détermination des vitesses initiales des planètes
        sys1.rAmplitVit = sys1.rAmplitVitMin +
            (sys1.rAmplitVitMax - sys1.rAmplitVitMin) * arRnd(iRndAmplitVit1)
        sys2.rAmplitVit = sys2.rAmplitVitMin +
            (sys2.rAmplitVitMax - sys2.rAmplitVitMin) * arRnd(iRndAmplitVit2)
        sys1.rAmplitVit = CDec(sys1.rAmplitVit *
            Math.Sqrt(m_prm.rForceGravitation / rRapportVitesse))
        sys2.rAmplitVit = CDec(sys2.rAmplitVit *
            Math.Sqrt(m_prm.rForceGravitation / rRapportVitesse))

        If bTest3D Then
            sys1.rAmplitPos = m_aff.rMaxx * 0.5D
            sys2.rAmplitPos = sys1.rAmplitPos / 6
            If sys2.iNbPts <= 1 Then sys2.rAmplitPos = 0 : sys2.rAmplitVit = 0
        End If

        m_iNbSprites = 0
        ReDim m_aSprites(0)
        ReDim m_aCoordZ(0)
        ReDim m_aIndexCoordZ(0)
        'GC.Collect() ' Récuperer tout de suite la mémoire allouée des sprites

        ReDim sys2.aPlanete(sys2.iNbPts - 1)

        ReDim m_pt(m_iNbPtsTot - 1)
        If m_b3D Then 'Or bTestOrb3D Then
            ReDim m_aCoordZ(m_iNbPtsTot - 1)
            ReDim m_aIndexCoordZ(m_iNbPtsTot - 1)
        End If


        ' Recherche des multiples de 2 pour le système secondaire
        '  (la symétrie n'est possible que s'il y a un nombre pair de planètes)
        Dim iDivNbPts2%
        iDivNbPts2 = sys2.iNbPts Mod 2

        bMasseSym = m_prm.bMasseSym
        If m_prm.bMasseSym_bRnd Then _
            bMasseSym = (arRnd(iRndbMasseSym2) > 0.5) ' 1 chance sur 2

        For i = 0 To sys2.iNbPts - 1
            sys2.aPlanete(i).rMasse =
                rRandomiser(rAmplitMasseMin, rAmplitMasseMax)
            If iNbFichiersPlanetes > 0 Then
                sys2.aPlanete(i).iNumImg = iRandomiser(0, iNbFichiersPlanetes - 1)
            End If
            sys2.aPlanete(i).rSpin = rRandomiser(0, 1) 'CDec(Rnd())
            'sys2.aPlanete(i).rSpin = rRandomiser(-rSpinMaxDeg, rSpinMaxDeg)
        Next i

        If bMasseSym Then
            iNbPts2Sur2 = sys2.iNbPts \ 2
            If sys2.iNbPts = 1 Then iNbPts2Sur2 = 1
            ReDim sys2.aPlaneteSym(iNbPts2Sur2)
            For i = 0 To iNbPts2Sur2 - 1
                sys2.aPlaneteSym(i).rMasse =
                    rRandomiser(rAmplitMasseMin, rAmplitMasseMax)
                If iNbFichiersPlanetes > 0 Then
                    sys2.aPlaneteSym(i).iNumImg =
                        iRandomiser(0, iNbFichiersPlanetes - 1, arRnd(iRndNumImgSym2))
                End If
                sys2.aPlaneteSym(i).rSpin = arRnd(iRndSpinImgSym2)
            Next i
        End If

        ' Calcul des racines unitaires complexes : Z^n = 1
        '  avec Z un nombre complexe et n = degré ou nombre de planètes par système

        Dim rAngleDepart1 As Decimal '= 0 'Math.PI / 2
        Dim rAngleDepart2 As Decimal '= 0 'Math.PI / 2
        rAngleDepart1 = CDec(arRnd(iRndAngleDepart1) * Math.PI * 2)
        rAngleDepart2 = CDec(arRnd(iRndAngleDepart2) * Math.PI * 2)

        i = 0 : j = 0
        For l = 0 To sys1.iNbPts - 1
            For k = 0 To sys2.iNbPts - 1
                i = l * sys2.iNbPts
                sys1.rAngle = CDec(rAngleDepart1 + 2D * Math.PI * l / sys1.iNbPts)
                sys2.rAngle = CDec(rAngleDepart2 + sys1.rAngle +
                2D * Math.PI * k / sys2.iNbPts)
                m_pt(i + k).rX = CDec(m_aff.rMaxx * 0.5 +
                sys1.rAmplitPos * Math.Cos(sys1.rAngle) +
                sys2.rAmplitPos * Math.Cos(sys2.rAngle))
                m_pt(i + k).rY = CDec(m_aff.rMaxy * 0.5 +
                sys1.rAmplitPos * Math.Sin(sys1.rAngle) +
                sys2.rAmplitPos * Math.Sin(sys2.rAngle))
                m_pt(i + k).rZ = 0
                If m_b3D Then
                    If bTest3D Then
                        sys2.rAngle = CDec(2D * Math.PI * k / sys2.iNbPts)
                        m_pt(i + k).rX = CDec(m_aff.rMaxx * 0.5 +
                        sys2.rAmplitPos * Math.Sin(sys2.rAngle + sys2.rAxz) *
                        Math.Sin(sys1.rAngle + sys2.rAxyP) +
                        sys1.rAmplitPos * Math.Cos(sys1.rAngle))
                        m_pt(i + k).rY = CDec(m_aff.rMaxy * 0.5 +
                        sys2.rAmplitPos * Math.Sin(sys2.rAngle + sys2.rAxz) *
                        Math.Cos(sys1.rAngle + sys2.rAxyP) +
                        sys1.rAmplitPos * Math.Sin(sys1.rAngle))
                        m_pt(i + k).rZ = CDec(m_aff.rMaxz * 0.5 +
                        sys2.rAmplitPos * Math.Cos(sys2.rAngle + sys2.rAxz))
                    Else
                        m_pt(i + k).rX = CDec(m_aff.rMaxx * 0.5 +
                        sys1.rAmplitPos * Math.Cos(sys1.rAngle) +
                        sys2.rAmplitPos * Math.Cos(sys2.rAngle))
                        m_pt(i + k).rY = m_aff.rMaxy * 0.5D
                        m_pt(i + k).rZ = CDec(m_aff.rMaxz * 0.5 +
                        sys1.rAmplitPos * Math.Sin(sys1.rAngle) +
                        sys2.rAmplitPos * Math.Sin(sys2.rAngle))
                    End If
                End If
                ' Déphasage des vecteurs vitesses par rapport aux positions
                sys1.rAngle = CDec(rAngleDepart1 +
                Math.PI * (2D * l / sys1.iNbPts - 0.5D))
                sys2.rAngle = CDec(rAngleDepart2 + sys1.rAngle +
                2D * Math.PI * k / sys2.iNbPts)
                m_pt(i + k).rVx = CDec(sys1.rAmplitVit * Math.Cos(sys1.rAngle) +
                sys2.rAmplitVit * Math.Cos(sys2.rAngle))
                m_pt(i + k).rVy = CDec(sys1.rAmplitVit * Math.Sin(sys1.rAngle) +
                sys2.rAmplitVit * Math.Sin(sys2.rAngle))
                m_pt(i + k).rVz = 0

                If m_b3D Then
                    If bTest3D Then
                        sys1.rAngle = CDec(2D * Math.PI * l / sys1.iNbPts)
                        sys2.rAngle = CDec(2D * Math.PI * k / sys2.iNbPts)
                        m_pt(i + k).rVx = CDec(-sys2.rAmplitVit *
                        Math.Cos(sys2.rAngle + sys2.rAxz) *
                        Math.Sin(sys1.rAngle + sys2.rAxy) -
                        sys1.rAmplitVit * Math.Sin(sys1.rAngle))
                        m_pt(i + k).rVy = CDec(-sys2.rAmplitVit *
                        Math.Cos(sys2.rAngle + sys2.rAxz) *
                        Math.Cos(sys1.rAngle + sys2.rAxy) +
                        sys1.rAmplitVit * Math.Cos(sys1.rAngle))
                        m_pt(i + k).rVz = CDec(sys2.rAmplitVit * Math.Sin(sys2.rAngle + sys2.rAxz))
                    Else
                        m_pt(i + k).rVy = 0
                        m_pt(i + k).rVz = CDec(
                        sys1.rAmplitVit * Math.Sin(sys1.rAngle) +
                        sys2.rAmplitVit * Math.Sin(sys2.rAngle))
                    End If
                End If

                m_pt(i + k).rM = sys2.aPlanete(k).rMasse
                If bMasseSym Then m_pt(i + k).rM = sys2.aPlaneteSym(j).rMasse

                If bDebugChoc Then
                    m_pt(i + k).rVx = 0
                    m_pt(i + k).rVy = 0
                    m_pt(i + k).rVz = 0
                    'If i + k = 0 Then m_pt(i + k).rVy = 1
                    'If i + k = 0 Then m_pt(i + k).rVx = 1
                    m_pt(i + k).rM = 50 '+ i * 50
                End If

                If m_iNbSprites = 0 Then
                    ReDim m_aSprites(0)
                Else
                    ReDim Preserve m_aSprites(m_iNbSprites)
                End If
                m_iNbSprites += 1

                ' Diametre du cercle : il dépend du zoom, car la détection
                '  des chocs est basée sur les positions qui sont zoomées
                Dim iDiametre% = CInt(m_aff.rZoom * m_pt(i + k).rM * 2)
                ' En 3D le diamètre est proportionnel à la moitié du zoom
                'Or bTestOrb3D Then _
                If m_b3D Then _
                iDiametre = CInt(0.5 * m_aff.rZoom * m_pt(i + k).rM * 2)

                Dim rRnd! = sys2.aPlanete(k).rSpin
                If bMasseSym Then rRnd = sys2.aPlaneteSym(j).rSpin
                Dim rDeltaAngleRotImg! = CSng(2 * rSpinMaxDeg * (rRnd - 0.5))
                ' Raffinement : tentative d'équilibrage des moments d'inertie
                If bMasseSym Then rDeltaAngleRotImg *= CSng(Math.Pow(-1, i + k))
                m_aSprites(i + k) = New Sprite(iDiametre, rDeltaAngleRotImg)

                If m_prm.bCercle Or iNbFichiersPlanetes = 0 Then
                    m_aSprites(i + k).m_bCercle = True
                Else
                    m_aSprites(i + k).m_bCercle = False
                    Dim iNumImg% = sys2.aPlanete(k).iNumImg
                    If bMasseSym Then iNumImg = sys2.aPlaneteSym(j).iNumImg
                    Dim sFichierImagePlanete$ = asFichiersImg(iNumImg)
                    m_aSprites(i + k).InitialiserImage(sFichierImagePlanete)
                End If

                If bMasseSym Then
                    j = j + 1
                    If j >= iNbPts2Sur2 Then j = 0
                End If

            Next k
        Next l

        'If Not bTestOrb3D Then Exit Sub
        If Not (m_b3D And m_b3D_bPlanetesAxeV) Then Exit Sub

        ' Test Orbites 3D
        ' ToDo : faire une fonction pour l'ajout d'un sprite
        For i = m_iNbPtsTot - iNbPlanetesZ To m_iNbPtsTot - 1
            If m_b3D Then
                m_pt(i).rX = CDec(m_aff.rMaxx * 0.5)
                m_pt(i).rY = CDec(m_aff.rMaxy * (2 * Rnd() - 1))
                m_pt(i).rZ = CDec(m_aff.rMaxz * 0.5)
            Else ' ???
                m_pt(i).rX = CDec(m_aff.rMaxx * 0.5)
                m_pt(i).rY = CDec(m_aff.rMaxy * 0.5)
                m_pt(i).rZ = CDec(m_aff.rMaxz * (2 * Rnd() - 1))
            End If
            m_pt(i).rVx = 0
            m_pt(i).rVy = 0
            m_pt(i).rVz = 0
            j = i - (m_iNbPtsTot - iNbPlanetesZ)
            m_pt(i).rM = planeteAxeZ.aPlanete(j).rMasse
            If m_pt(i).rM = 0 Then Stop
            Dim iDiametre0% = CInt(m_aff.rZoom * m_pt(i).rM * 2)
            'Or bTestOrb3D Then _
            If m_b3D Then _
                iDiametre0 = CInt(0.5 * m_aff.rZoom * m_pt(i).rM * 2)
            ReDim Preserve m_aSprites(m_iNbSprites)
            m_iNbSprites += 1
            Dim rSpin! = planeteAxeZ.aPlanete(j).rSpin
            m_aSprites(i) = New Sprite(iDiametre0, rSpin)
            If m_prm.bCercle Or iNbFichiersPlanetes = 0 Then
                m_aSprites(i).m_bCercle = True
            Else
                m_aSprites(i).m_bCercle = False
                Dim iNumImg% = planeteAxeZ.aPlanete(j).iNumImg
                'If bMasseSym Then iNumImg = planeteAxeZ.aPlaneteSym(j).iNumImg
                Dim sFichierImagePlanete$ = asFichiersImg(iNumImg)
                m_aSprites(i).InitialiserImage(sFichierImagePlanete)
            End If
        Next i

    End Sub

    Private Function rLireAngle(rAbscisse As Decimal, rOrdonnee As Decimal) As Decimal

        Dim rInterm As Decimal
        If (rAbscisse <> 0) Then
            rInterm = CDec(Math.Atan(Math.Abs(rOrdonnee / rAbscisse)))
        Else
            If (rOrdonnee < 0) Then
                rInterm = CDec(0.5D * Math.PI)
            Else
                rInterm = CDec(1.5D * Math.PI)
            End If
        End If

        If (rAbscisse > 0 And rOrdonnee > 0) Then _
            rInterm = CDec(2D * Math.PI - rInterm)
        If (rAbscisse < 0 And rOrdonnee <= 0) Then _
            rInterm = CDec(Math.PI - rInterm)
        If (rAbscisse < 0 And rOrdonnee >= 0) Then _
            rInterm = CDec(Math.PI + rInterm)
        rLireAngle = rInterm

    End Function

#End Region

#Region "Traitements"

    Public Sub SimulerGravite()

        Dim i%, j%
        Dim rDy, rFacteurGravtitation, rNorme2, rMinNorme, rNorme, rDx, rDz As Decimal
        Dim rMVy, rMFG, rSFGy, rSFGx, rSFGz, rMVx, rMVz As Decimal
        Dim rSM As Decimal ' Somme des masses
        Dim rCMy, rCMx, rCMz As Decimal ' Centres de masse

        ' Détermination de la stabilité de la vitesse totale du système
        If Not m_bSystemeInitialise Then
            m_bSystemeInitialise = True
            If (m_prm.bSystemeFixe) Then
                rMVx = 0 : rMVy = 0 : rMVz = 0
                rCMx = 0 : rCMy = 0 : rCMz = 0
                For i = 0 To m_iNbPtsTot - 1
                    rMVx += m_pt(i).rM * m_pt(i).rVx
                    rMVy += m_pt(i).rM * m_pt(i).rVy
                    rMVz += m_pt(i).rM * m_pt(i).rVz
                    rSM = rSM + m_pt(i).rM
                    rCMx += m_pt(i).rM * m_pt(i).rX
                    rCMy += m_pt(i).rM * m_pt(i).rY
                    rCMz += m_pt(i).rM * m_pt(i).rZ
                Next i

                ' Stabilisation du système au centre de l'écran
                For i = 0 To m_iNbPtsTot - 1
                    m_pt(i).rVx -= rMVx / rSM  ' Vitesses
                    m_pt(i).rVy -= rMVy / rSM
                    m_pt(i).rVz -= rMVz / rSM
                    m_pt(i).rX += m_aff.rMaxx / 2 - rCMx / rSM ' Positions
                    m_pt(i).rY += m_aff.rMaxy / 2 - rCMy / rSM
                    m_pt(i).rZ += m_aff.rMaxz / 2 - rCMz / rSM
                Next i
            End If
        End If

        For i = 0 To m_iNbPtsTot - 1

            Dim rX! = m_pt(i).rX
            Dim rY! = m_pt(i).rY
            Dim rZ! = m_pt(i).rZ
            Dim rVX! = m_pt(i).rVx
            Dim rVY! = m_pt(i).rVy
            Dim rVZ! = m_pt(i).rVz

            m_pt(i).iNbChocs = 0
            m_pt(i).rSomX = 0 : m_pt(i).rSomY = 0
            m_pt(i).rSomVX = 0 : m_pt(i).rSomVY = 0

            m_pt(i).rAx = 0 : m_pt(i).rAy = 0 : m_pt(i).rAz = 0
        Next i

        For i = 0 To m_iNbPtsTot - 1
            For j = i + 1 To m_iNbPtsTot - 1
                ' Analyse de la distance entre chaque planète 2 à 2
                '  afin d'en évaluer l'attraction gravitationnelle
                rDx = m_pt(j).rX - m_pt(i).rX
                rDy = m_pt(j).rY - m_pt(i).rY
                rDz = m_pt(j).rZ - m_pt(i).rZ

                rNorme2 = rDx * rDx + rDy * rDy

                'Or bTestOrb3D 
                If m_b3D Then rNorme2 += rDz * rDz

                rNorme = CDec(Math.Sqrt(rNorme2))
                ' La masse est représentée par le rayon,
                '  l'instant du choc correpond au 2 rayons
                rMinNorme = (m_pt(i).rM + m_pt(j).rM)

                If (rNorme >= rMinNorme) Then _
                rFacteurGravtitation = 1 / rNorme2 : GoTo Suite

                If Not m_bChocs Then
                    ' Au lieu de simuler un choc lorsque les planètes se touchent,
                    '  on annule progressivement la gravité comme si les
                    '  planètes se superposaient en 3D. On imagine que les planètes
                    '  s'attirent sans jamais se toucher : fantômes
                    '  La norme passe du coté supérieur de la fraction pour
                    '  inverser la gravité !
                    '  On divise par rMinNorme3 afin que la force gravitationnelle
                    '  soit continue au point de contact !
                    rFacteurGravtitation = rNorme / (rMinNorme * rMinNorme * rMinNorme)
                    ' Minoration de la norme pour minorer l'attraction
                    rNorme = rMinNorme
                    GoTo Suite
                End If

                GererChoc(i, j, rDx, rDy, rDz,
                rNorme, rMinNorme, rNorme2, rFacteurGravtitation)

Suite:
                ' Loi de la gravitation
                ' rMFG = module de la force gravitationnelle
                rMFG = m_prm.rForceGravitation * m_pt(i).rM * m_pt(j).rM *
                rFacteurGravtitation
                ' [rDx, rDy] / rNorme = vecteur unitaire
                ' rSFG = somme des forces gravitationnelles
                rSFGx = rMFG * rDx / rNorme
                rSFGy = rMFG * rDy / rNorme
                rSFGz = rMFG * rDz / rNorme
                ' Transformation de la force de gravité en accélération
                m_pt(i).rAx += rSFGx / m_pt(i).rM
                m_pt(i).rAy += rSFGy / m_pt(i).rM
                m_pt(i).rAz += rSFGz / m_pt(i).rM
                m_pt(j).rAx -= rSFGx / m_pt(j).rM
                m_pt(j).rAy -= rSFGy / m_pt(j).rM
                m_pt(j).rAz -= rSFGz / m_pt(j).rM
            Next j : Next i

        ' Vérification si tous les points sont sortis de l'écran
        m_bToutesPlanetesHorsEcran = True

        ' Détermination de la nouvelle position de chaque planète : (iH, iV)
        m_rectMAJGroupeSprites = New Rectangle(0, 0, 0, 0)

        Dim iV%, iH%, iRayon%, iDiametre%, rZoomDA!
        Const iRayonMinAffichage% = 3

        For i = 0 To m_iNbPtsTot - 1

            If m_pt(i).iNbChocs > 0 And bChocsMoyennes Then
                m_pt(i).rVx = m_pt(i).rSomVX / m_pt(i).iNbChocs
                m_pt(i).rVy = m_pt(i).rSomVY / m_pt(i).iNbChocs
                m_pt(i).rX = m_pt(i).rSomX / m_pt(i).iNbChocs
                m_pt(i).rY = m_pt(i).rSomY / m_pt(i).iNbChocs
            End If
            If bGravite And Not m_prm.bPauseAnimation Then
                ' m_bGravite : Booléen pour annuler la gravité (debug choc)
                m_pt(i).rVx += m_pt(i).rAx
                m_pt(i).rVy += m_pt(i).rAy
                m_pt(i).rVz += m_pt(i).rAz
            End If

            ' Transformation des vitesses en mouvement effectif
            If Not m_prm.bPauseAnimation Then
                m_pt(i).rX += m_pt(i).rVx
                m_pt(i).rY += m_pt(i).rVy
                m_pt(i).rZ += m_pt(i).rVz
            End If

            ProjeterCoord(iH, iV, rZoomDA,
                m_pt(i).rX, m_pt(i).rY, m_pt(i).rZ)

            m_aSprites(i).FixerPosition(New Point(iH, iV))
            m_aSprites(i).DiametreApparent(rZoomDA, iRayonMinAffichage)
            m_aSprites(i).m_bCercle = m_prm.bCercle
            m_aSprites(i).m_bPositionInitialisee = True

            iRayon = CInt(m_pt(i).rM * m_aff.rZoom)

            If m_b3D Then 'Or bTestOrb3D Then
                ' En 3D le diametre est 2 fois + faible,
                '  car il est projeté 50-50 sur x et z / y et z
                iDiametre = CInt(iRayon * rZoomDA)
            Else
                iDiametre = CInt(iRayon * 2 * rZoomDA)
            End If
            ' Il faut une marge à cause de la rotation
            iDiametre = CInt(iDiametre * 1.2)

            iRayon = iDiametre \ 2
            If iRayon < iRayonMinAffichage Then _
                iRayon = iRayonMinAffichage
            If iDiametre < 2 * iRayonMinAffichage Then _
                iDiametre = 2 * iRayonMinAffichage

            ' Méthode de MAJ précise : on invalide chaque rectangle : 
            '  c'est + rapide car la zone de MAJ est + petite
            m_aSprites(i).m_rectPos = New Rectangle(
                iH - iRayon - 1, iV - iRayon - 1,
                iDiametre + 2, iDiametre + 2)
            If m_aSprites(i).m_rectMemPos.Width = 0 Then
                m_aSprites(i).m_rectMAJ = m_aSprites(i).m_rectPos
            Else
                m_aSprites(i).m_rectMAJ = Rectangle.Union(
                    m_aSprites(i).m_rectPos,
                    m_aSprites(i).m_rectMemPos)
            End If

            ' Méthode de MAJ en groupe : jolie pour les transitions
            ' bMAJGroupeSprites = True
            If m_rectMAJGroupeSprites.Width = 0 Then
                m_rectMAJGroupeSprites = m_aSprites(i).m_rectPos
            Else
                m_rectMAJGroupeSprites = Rectangle.Union(
                    m_rectMAJGroupeSprites,
                    m_aSprites(i).m_rectPos)
            End If
            If m_aSprites(i).m_rectMemPos.Width <> 0 Then _
                m_rectMAJGroupeSprites = Rectangle.Union(
                    m_rectMAJGroupeSprites,
                    m_aSprites(i).m_rectMemPos)

            m_aSprites(i).m_rectMemPos = m_aSprites(i).m_rectPos

            If iH + iRayon >= 0 And iH - iRayon <= m_aff.rMaxH And
                iV + iRayon >= 0 And iV - iRayon <= m_aff.rMaxV Then _
                m_bToutesPlanetesHorsEcran = False

        Next i

    End Sub

    Private Sub GererChoc(i%, j%,
        ByRef rDx As Decimal, ByRef rDy As Decimal, ByRef rDz As Decimal,
        ByRef rNorme As Decimal, ByRef rMinNorme As Decimal,
        ByRef rNorme2 As Decimal, ByRef rFacteurGravtitation As Decimal)

        ' Gestion des chocs, merci à Alcys :
        ' CHOCS ENTRE BILLES DE MASSES DIFFERENTES : 
        ' http://www.flashkod.com/article.aspx?Val=118
        ' J'ai ajouté une correction de position au point précis du choc
        '  ainsi qu'une tentative de moyennage des chocs (cf. + loin)

        Const rAmortissement As Decimal = 1D ' Choc élastique
        Dim xi, yi, vxi, vyi, ri, mi As Decimal
        Dim xj, yj, vxj, vyj, rj, mj As Decimal
        Dim a, b, cc, c As Decimal

        xi = m_pt(i).rX
        vxi = m_pt(i).rVx
        ri = m_pt(i).rM ' Rayon
        mi = m_pt(i).rM

        xj = m_pt(j).rX
        vxj = m_pt(j).rVx
        rj = m_pt(j).rM
        mj = m_pt(j).rM

        ' Distance entre les centres des 2 boules
        Dim Dxu, Dyu As Decimal ' u pour unitaire : normé
        Dim rAngleChoc As Decimal ' rAngleChoc : angle de l'axe du choc

        Dxu = rDx / rNorme

        If m_b3D Then
            yi = m_pt(i).rZ
            vyi = m_pt(i).rVz
            yj = m_pt(j).rZ
            vyj = m_pt(j).rVz
            Dyu = rDz / rNorme
        Else
            yi = m_pt(i).rY
            vyi = m_pt(i).rVy
            yj = m_pt(j).rY
            vyj = m_pt(j).rVy
            Dyu = rDy / rNorme
        End If

        rAngleChoc = rLireAngle(Dxu, Dyu)
        m_pt(i).rAngleChoc = rAngleChoc
        m_pt(j).rAngleChoc = rAngleChoc
        Dim rCosAngleChoc, rSinAngleChoc As Decimal
        rCosAngleChoc = CDec(Math.Cos(rAngleChoc))
        rSinAngleChoc = CDec(Math.Sin(rAngleChoc))

        ' Distances de correction : 
        '  véritable position du cercle au moment du choc
        Dim rDistCorrection_i, rDistCorrection_j As Decimal
        Dim rDistCorrection As Decimal
        rDistCorrection = rMinNorme - rNorme
        rDistCorrection_i = rDistCorrection * mi / (mi + mj)
        rDistCorrection_j = rDistCorrection * mj / (mi + mj)

        xi -= rDistCorrection_i * rCosAngleChoc
        yi += rDistCorrection_i * rSinAngleChoc
        xj += rDistCorrection_j * rCosAngleChoc
        yj -= rDistCorrection_j * rSinAngleChoc

        If bDebugChoc Then
            ' Droite liant les centres
            'm_pt(j).rDx = -Dxu * 100
            'm_pt(j).rDy = -Dyu * 100
            m_pt(j).rDx = -100 * rCosAngleChoc
            m_pt(j).rDy = -100 * -rSinAngleChoc

            m_pt(i).rXC = xi
            m_pt(j).rXC = xj
            If m_b3D Then
                m_pt(i).rYC = m_pt(i).rY ' Inchangé
                m_pt(j).rYC = m_pt(j).rY ' Inchangé
                m_pt(i).rZC = yi
                m_pt(j).rZC = yj
            Else
                m_pt(i).rYC = yi
                m_pt(j).rYC = yj
                m_pt(i).rZC = m_pt(i).rZ ' Inchangé
                m_pt(j).rZC = m_pt(j).rZ ' Inchangé
            End If
        End If

        a = xi - xj
        b = yi - yj
        cc = CDec(Math.Sqrt(a * a + b * b))
        c = ri + rj
        'dis = a * a + b * b - c * c

        ' Quantité de mouvement avant et après le choc
        Dim q1, q2 As Decimal
        If bDebugChoc Then
            Dim qx1, qy1 As Decimal
            qx1 = vxi * mi + vxj * mj
            qy1 = vyi * mi + vyj * mj
            q1 = CDec(Math.Sqrt(qx1 * qx1 + qy1 * qy1))
        End If

        Dim nx, ny, tx, ty, rm, e As Decimal
        Dim xx, yy, xx1, yy1 As Decimal
        nx = a / cc
        ny = b / cc
        tx = -ny
        ty = nx
        rm = mj / mi ' Rapport des masses
        e = rAmortissement
        xx = (1 - rm * e) / (1 + rm) * (vxi * nx + vyi * ny) +
                rm * (1 + e) / (1 + rm) * (vxj * nx + vyj * ny)
        yy = vxi * tx + vyi * ty
        xx1 = (1 + e) / (1 + rm) * (vxi * nx + vyi * ny) +
                (rm - e) / (1 + rm) * (vxj * nx + vyj * ny)
        yy1 = vxj * tx + vyj * ty
        vxi = xx * nx + yy * tx
        vyi = xx * ny + yy * ty
        vxj = xx1 * nx + yy1 * tx
        vyj = xx1 * ny + yy1 * ty
        xi = xj + (c + 1) * nx
        yi = yj + (c + 1) * ny

        ' Ce n'est pas juste : par contre l'idée est à tester
        '  avec les centres de masse oppposés

        If bChocsMoyennes Then
            m_pt(i).iNbChocs += 1
            m_pt(i).rSomX += xi
            m_pt(i).rSomY += yi
            m_pt(i).rSomVX += vxi
            m_pt(i).rSomVY += vyi

            m_pt(j).iNbChocs += 1
            m_pt(j).rSomX += xj
            m_pt(j).rSomY += yj
            m_pt(j).rSomVX += vxj
            m_pt(j).rSomVY += vyj

            ' On utilise les positions corrigées pour refaire
            '  le calcul de l'accélération
            m_pt(i).rX = xi
            m_pt(i).rY = yi
            m_pt(j).rX = xj
            m_pt(j).rY = yj

        Else

            m_pt(i).iNbChocs += 1 ' Pour débug choc
            m_pt(j).iNbChocs += 1

            ' Ne marche pas très bien avec plus de 2 chocs simultanés
            '  car l'ordre des chocs est subjectif
            m_pt(i).rX = xi
            m_pt(i).rVx = vxi
            m_pt(j).rX = xj
            m_pt(j).rVx = vxj
            If m_b3D Then
                m_pt(i).rZ = yi
                m_pt(i).rVz = vyi
                m_pt(j).rZ = yj
                m_pt(j).rVz = vyj
            Else
                m_pt(i).rY = yi
                m_pt(i).rVy = vyi
                m_pt(j).rY = yj
                m_pt(j).rVy = vyj
            End If
        End If

        ' Correction des positions, donc de l'accélarion max.
        rDx = xj - xi
        If m_b3D Then
            'rDy est inchangé
            rDz = yj - yi
        Else
            rDy = yj - yi
            'rDz = 0 inchangé
        End If
        rNorme2 = rDx * rDx + rDy * rDy
        If m_b3D Then rNorme2 = rNorme2 + rDz * rDz
        rNorme = CDec(Math.Sqrt(rNorme2))
        rFacteurGravtitation = 1 / rNorme2

        If bDebugChoc Then
            ' Vérification de la conservation de la quantité de mouvement
            Dim qx2, qy2 As Decimal
            qx2 = vxi * mi + vxj * mj
            qy2 = vyi * mi + vyj * mj
            q2 = CDec(Math.Sqrt(qx2 * qx2 + qy2 * qy2))
            m_frm.Text = "Qté de mvt q2-q1 = " & q2 & " - " & q1 & " = " & q2 - q1 &
                ", Angle Choc = " & rAngleChoc
            glb_rDateMessageTitre = DateAndTime.Timer
        End If

    End Sub

    Private Sub ProjeterCoord(ByRef iH%, ByRef iV%, ByRef rZoomDA!, rX!, rY!, rZ!)

        rZoomDA = 1
        If m_b3D Then 'Or bTestOrb3D Then

            iH = CInt(0.5 * m_aff.rMaxH + m_aff.rZoom * 0.5 *
                (rX - 0.5 * m_aff.rMaxx + rZ - 0.5 * m_aff.rMaxz))
            iV = CInt(0.5 * m_aff.rMaxV - m_aff.rZoom * 0.5 *
                (rY - 0.5 * m_aff.rMaxy + rZ - 0.5 * m_aff.rMaxz))

            ' Fixer le diamètre apparent de la planète en fonction
            '  de sa coordonnée de profondeur Z
            '  sauf si Choc en 3D : ce n'est pas consistant
            '  car on ne peut pas déterminer la position du choc
            If m_bChocs Then Exit Sub
            Dim rMinZ! = -m_aff.rMaxz
            Dim rMaxZ! = m_aff.rMaxz
            rZoomDA = CSng(1 - 0.5 * (rZ - rMinZ) / (rMaxZ - rMinZ))

        Else

            iH = CInt(0.5 * m_aff.rMaxH +
                m_aff.rZoom * (rX - 0.5 * m_aff.rMaxx))
            iV = CInt(0.5 * m_aff.rMaxV -
                m_aff.rZoom * (rY - 0.5 * m_aff.rMaxy))

        End If

    End Sub

    Public Sub Dessiner(ByRef dc As Graphics, bNePasBufferiserGr As Boolean)

        ' Dessin général

        If dc.SmoothingMode <> SmoothingMode.HighSpeed Then _
            dc.SmoothingMode = SmoothingMode.HighSpeed

        ' Fonctionnalité du GDI+ pas encore disp. en .Net :
        ' InterpolationModeNearestNeighbor is the lowest-quality mode and
        ' InterpolationModeHighQualityBicubic is the highest-quality mode.
        'dc.SetInterpolationMode(InterpolationModeNearestNeighbor)

        If Not bNePasBufferiserGr Then DessinerFond(dc)

        Dim i%

        If m_b3D Then 'Or bTestOrb3D Then
            If m_iNbPtsTot = 0 Then Exit Sub
            ' Tri des planètes dans l'ordre des Z décroissants
            For i = 0 To m_iNbPtsTot - 1
                m_aCoordZ(i) = -m_pt(i).rZ
                m_aIndexCoordZ(i) = i
            Next i

            Array.Sort(m_aCoordZ, m_aIndexCoordZ)

            Dim iIndexZDec%
            For i = 0 To m_iNbSprites - 1
                iIndexZDec = m_aIndexCoordZ(i)
                ' On ne peut voir les traces que si l'on ne bufférise pas
                m_aSprites(iIndexZDec).m_bLaisserTraceCercleGris = bNePasBufferiserGr
                m_aSprites(iIndexZDec).AnimerSpin()
                m_aSprites(iIndexZDec).Dessiner(dc)

                If bDebugRectMAJ Then _
                    dc.DrawRectangle(New Pen(Color.Blue, 2),
                        m_aSprites(i).m_rectMAJ)

                If m_pt(i).iNbChocs > 0 And bDebugChoc Then
                    Dim iH1%, iH2%, iV1%, iV2%, rZoomDA!
                    ProjeterCoord(iH1, iV1, rZoomDA,
                        m_pt(i).rX, m_pt(i).rY, m_pt(i).rZ)
                    ProjeterCoord(iH2, iV2, rZoomDA,
                        m_pt(i).rX + m_pt(i).rDx, m_pt(i).rY,
                        m_pt(i).rZ + m_pt(i).rDy)
                    dc.DrawLine(New Pen(Color.Blue, 2), iH1, iV1, iH2, iV2)
                    Dim iH%, iV%
                    ProjeterCoord(iH, iV, rZoomDA, m_pt(i).rXC, m_pt(i).rYC, m_pt(i).rZC)
                    Dim iRayon% = CInt(m_pt(i).rM * m_aff.rZoom * 0.5 * rZoomDA)
                    Dim iDiam% = 2 * iRayon
                    dc.DrawEllipse(New Pen(Color.Blue, 2),
                        iH - iRayon, iV - iRayon, iDiam, iDiam)
                End If

                If bDebugPosEtVitInitiales Then
                    Dim iH1%, iH2%, iV1%, iV2%, rZoomDA!
                    ProjeterCoord(iH1, iV1, rZoomDA,
                        m_pt(i).rX, m_pt(i).rY, m_pt(i).rZ)
                    ProjeterCoord(iH2, iV2, rZoomDA,
                        m_pt(i).rX + m_pt(i).rVx * 20,
                        m_pt(i).rY + m_pt(i).rVy * 20,
                        m_pt(i).rZ + m_pt(i).rVz * 20)
                    dc.DrawLine(New Pen(Color.Yellow, 2), iH1, iV1, iH2, iV2)
                End If

            Next i

        Else ' 2D
            For i = 0 To m_iNbSprites - 1
                'm_aSprites(i).m_bCercle = m_prm.bCercle
                m_aSprites(i).m_bLaisserTraceCercleGris = bNePasBufferiserGr
                m_aSprites(i).AnimerSpin()
                m_aSprites(i).Dessiner(dc)
                If bDebugRectMAJ Then _
                    dc.DrawRectangle(New Pen(Color.Blue, 2),
                        m_aSprites(i).m_rectMAJ)

                If m_pt(i).iNbChocs > 0 And bDebugChoc Then
                    Dim iH1%, iH2%, iV1%, iV2%, rZoomDA!
                    ProjeterCoord(iH1, iV1, rZoomDA,
                        m_pt(i).rX, m_pt(i).rY, m_pt(i).rZ)
                    ProjeterCoord(iH2, iV2, rZoomDA,
                        m_pt(i).rX + m_pt(i).rDx,
                        m_pt(i).rY + m_pt(i).rDy, m_pt(i).rZ)
                    dc.DrawLine(New Pen(Color.Blue, 2), iH1, iV1, iH2, iV2)

                    Dim iH%, iV%
                    ProjeterCoord(iH, iV, rZoomDA,
                        m_pt(i).rXC, m_pt(i).rYC, m_pt(i).rZC)
                    Dim iRayon% = CInt(m_pt(i).rM * m_aff.rZoom)
                    Dim iDiam% = 2 * iRayon
                    dc.DrawEllipse(New Pen(Color.Blue, 2),
                        iH - iRayon, iV - iRayon, iDiam, iDiam)
                End If

                If bDebugPosEtVitInitiales Then
                    Dim iH1%, iH2%, iV1%, iV2%, rZoomDA!
                    ProjeterCoord(iH1, iV1, rZoomDA,
                        m_pt(i).rX, m_pt(i).rY, m_pt(i).rZ)
                    ProjeterCoord(iH2, iV2, rZoomDA,
                        m_pt(i).rX + m_pt(i).rVx * 20,
                        m_pt(i).rY + m_pt(i).rVy * 20, m_pt(i).rZ)
                    dc.DrawLine(New Pen(Color.Yellow, 2), iH1, iV1, iH2, iV2)
                End If

            Next i
        End If

        If bDebugRectMAJ Then _
            dc.DrawRectangle(New Pen(Color.Blue, 2), m_rectMAJGroupeSprites)

    End Sub

    Public Sub DessinerFond(ByRef dc As Graphics)

        If Not m_prm.bFondUni And Not m_prm.bFondDegrade And Not m_bImageFondTrouve Then Exit Sub

        ' Ne pas utiliser les variables images avant leur initialisation
        If Not m_bImgFondInitialisee Then Exit Sub

        Dim rectFondLocal As Rectangle
        rectFondLocal = m_rectEcran

        If m_prm.bFondUni Then dc.Clear(Color.Navy)

        ' Créer un fond avec un dégradé de couleur
        If m_prm.bFondDegrade Then _
            dc.FillRectangle(m_lgbFondDegrade, rectFondLocal)

        If Not m_bImageFondTrouve Then Exit Sub

        If m_prm.bNePasAgrandirImgFond Then
            ' Ne pas agrandir pour optimiser la vitesse
            dc.DrawImage(m_imgFond, 0, 0, m_imgFond.Width, m_imgFond.Height)
            Exit Sub
        End If

        If m_prm.bNePasDecentrerImgFond Then
            ' Ok pour le clipping, mais ne marche pas avec le décentrage
            Dim rectDest As Rectangle = rectFondLocal
            Dim rectSrc As Rectangle = rectFondLocal
            dc.DrawImage(m_imgFond, rectDest, rectSrc, GraphicsUnit.Pixel)
            Exit Sub
        End If

        ' Ok avec le décentrage mais ne marche pas avec le clipping
        '  (les rectangles sont plus compliqués à calculer dans ce cas)
        'dc.DrawImage(m_imgFond, m_rectEcran, m_rectImgFond, GraphicsUnit.Pixel)

        ' Sinon, tracé de toute l'image de fond sans optimisation
        '  Pb : Ratio H/V non préservé
        'dc.DrawImage(m_imgFond, 0, 0, m_rectEcran.Width, m_rectEcran.Height)

        ' Ok : pas d'optimisation, mais le décentrage marche et
        '  le ration H/V est préservé : cf. calcul de m_rectImgFond
        dc.DrawImage(m_imgFond, m_rectImgFond)

    End Sub

    Public Sub InitialiserImageFond(szClientSize As Size)

        m_bImageFondTrouve = False
        If Not m_prm.bImageFond Then GoTo Fin

        Dim sRepertoire$ = Application.StartupPath
        Dim sFichierImageFond$ 'sFichierImagePlanete$, 

        ' Recherche des fichiers images
        Dim asFichiersImg() As String = Nothing
        Dim iNbFichiersImg% = 0
        Dim sFiltre$ = m_prm.sFiltreFichiersImgFond
        Dim sRepertoireImgFond$ = sRepertoire
        If sFiltre = "" Then _
            m_bImageFondTrouve = False : GoTo Fin
        Dim iPos% = sFiltre.IndexOf("\")
        If iPos > 0 Then
            sRepertoireImgFond = sRepertoire & "\" &
                sFiltre.Substring(0, iPos)
            sFiltre = sFiltre.Substring(iPos + 1)
        End If

        If Not bDossierExiste(sRepertoireImgFond) Then GoTo Fin

        Try
            asFichiersImg = Directory.GetFiles(sRepertoireImgFond, sFiltre) ' "space_*.jpg")
            iNbFichiersImg = asFichiersImg.Length
        Catch
            iNbFichiersImg = 0
        End Try

        If iNbFichiersImg = 0 Then _
            m_bImageFondTrouve = False : GoTo Fin

        m_bImageFondTrouve = True
        Dim iNumImg% = CInt(Rnd() * iNbFichiersImg)
        If iNumImg >= iNbFichiersImg Then iNumImg = iNumImg - 1
        sFichierImageFond = asFichiersImg(iNumImg)
        'MsgBox("Fichier choisi : " & sFichierImageFond)

        If Not m_imgFond Is Nothing Then m_imgFond.Dispose()

        'm_imgFond = Image.FromFile(sFichierImageFond)
        m_imgFond = CType(Image.FromFile(sFichierImageFond), Bitmap)

        If m_prm.bNePasDecentrerImgFond Then
            m_rectImgFond = New Rectangle(0, 0, m_imgFond.Width, m_imgFond.Height)
        Else
            ' Décentrage de l'image à gauche et en Haut,
            '  ainsi qu'à droite et en bas,
            '  ceci afin de ne jamais afficher une image au même endroit
            '  (un écran de veille sert justement à éviter cela)

            m_rectEcran = New Rectangle(0, 0,
                szClientSize.Width, szClientSize.Height)

            Dim rZoomV! = CSng(m_rectEcran.Height / m_imgFond.Height)
            Dim rZoomH! = CSng(m_rectEcran.Width / m_imgFond.Width)
            Dim rMaxZoom! = rZoomV
            If rZoomH > rMaxZoom Then rMaxZoom = rZoomH

            Dim rAg! = 1 + 1 * Rnd()
            Dim rDec! = Rnd()
            'rAg = 1 : rDec = 0 : Centrée
            'rAg = 2 : rDec = 1 : Quart GH
            'rAg = 1 : rDec = 1 : Quart DB
            Dim rZoomDeb! = rMaxZoom * rDec
            Dim rZoomFin! = rMaxZoom * (rAg + rDec)

            m_rectImgFond = New Rectangle(
                CInt(-m_imgFond.Width * rZoomDeb),
                CInt(-m_imgFond.Height * rZoomDeb),
                CInt(m_imgFond.Width * rZoomFin),
                CInt(m_imgFond.Height * rZoomFin))

        End If

Fin:
        m_bImgFondInitialisee = True

    End Sub

    Public Sub InitialiserTailleEcran(szClientSize As Size)

        m_rectEcran = New Rectangle(0, 0,
            szClientSize.Width, szClientSize.Height)
        m_szTailleFenetre = szClientSize

        m_lgbFondDegrade = New LinearGradientBrush(m_rectEcran,
            Color.Red, Color.Yellow, LinearGradientMode.BackwardDiagonal)

    End Sub

#End Region

End Class