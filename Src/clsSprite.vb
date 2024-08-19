
Public Class Sprite : Implements IDisposable ' Petite image en mouvement

    Public m_bCercle As Boolean = False
    Public m_bLaisserTraceCercleGris As Boolean = False
    Private Const m_bTransparenceImg As Boolean = True
    Private Const m_bRotationImg As Boolean = True

    Public m_rectMAJ, m_rectPos, m_rectMemPos As Rectangle
    Private m_szTailleImg As Size
    Private m_iGdCoteImg%
    Private m_rZoomImg! = 1
    Public m_bPositionInitialisee As Boolean = False
    Public m_ptPos As New Point(0, 0)
    Private m_ptMemPos As New Point(0, 0)
    Private m_iRayon%, m_iMemRayon%, m_iMemDiam%

    ' Tracé d'un cercle
    Private m_iDiametreCercle% = 100
    Private m_iDiametreApparentCercle%
    Private Const m_iLargPinceauCercle% = 2
    ' Couleur du cercle
    Private m_penCercle As New Pen(Color.Cyan, m_iLargPinceauCercle)
    ' Laisser une trace du cercle 
    Private m_penCercleGris As New Pen(Color.DarkGray, m_iLargPinceauCercle)

    ' Pour avoir la méthode MakeTransparent (non disp. dans la classe Image)
    Private m_imgSprite As Bitmap
    Private m_rAngleRotImg!
    ' Variation de l'angle de rotation de l'image
    Private m_rDeltaAngleRotImg! = 2
    Private m_graphicsContainer As Drawing2D.GraphicsContainer

    ' Constructeur de la classe
    Public Sub New(iDiametre%, rDeltaAngleRotImg0!)
        m_bPositionInitialisee = False
        m_ptPos.X = 0 : m_ptPos.Y = 0
        m_iDiametreCercle = iDiametre
        m_iDiametreApparentCercle = m_iDiametreCercle
        m_iRayon = iDiametre \ 2
        m_rZoomImg = iDiametre
        m_rDeltaAngleRotImg = rDeltaAngleRotImg0
        m_iMemRayon = 0
    End Sub

    Protected Overridable Overloads Sub Dispose(disposing As Boolean)
        If disposing Then ' Dispose managed resources
            m_penCercle.Dispose()
            m_penCercleGris.Dispose()
        End If
        ' Free native resources
    End Sub
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Public Sub InitialiserImage(sCheminImage$)

        m_bPositionInitialisee = False

        If Not m_bCercle Then
            m_imgSprite = CType(Image.FromFile(sCheminImage), Bitmap)
            m_iGdCoteImg = m_imgSprite.Width
            If m_imgSprite.Height > m_iGdCoteImg Then _
                m_iGdCoteImg = m_imgSprite.Height
            CalculerTailleImg(m_rZoomImg)
            ' Définir la couleur de transparence du Bitmap
            If m_bTransparenceImg Then _
                m_imgSprite.MakeTransparent(m_imgSprite.GetPixel(1, 1))
        End If

    End Sub

    Private Sub CalculerTailleImg(rZoom!)
        m_szTailleImg.Width = CInt(m_imgSprite.Width * rZoom / m_iGdCoteImg - 1)
        m_szTailleImg.Height = CInt(m_imgSprite.Height * rZoom / m_iGdCoteImg - 1)
        m_iRayon = m_szTailleImg.Width \ 2
        If 0.5 * m_szTailleImg.Height > m_iRayon Then m_iRayon = m_szTailleImg.Height \ 2
    End Sub

    Public Sub AnimerSpin()
        m_rAngleRotImg = m_rAngleRotImg + m_rDeltaAngleRotImg
    End Sub

    Public Sub DiametreApparent(rZoomPosition!, iRayonMinAffichage%)

        If m_bCercle Or m_imgSprite Is Nothing Then
            If m_bPositionInitialisee Then
                m_iMemRayon = m_iRayon
                m_iMemDiam = m_iDiametreApparentCercle
            End If
            m_iDiametreApparentCercle = CInt(m_iDiametreCercle * rZoomPosition)
            m_iRayon = m_iDiametreApparentCercle \ 2
        Else
            Dim rZoomImg0! = m_rZoomImg * rZoomPosition
            CalculerTailleImg(rZoomImg0)
        End If
        If m_iRayon < iRayonMinAffichage Then _
            m_iRayon = iRayonMinAffichage
        If m_iDiametreApparentCercle < 2 * iRayonMinAffichage Then _
            m_iDiametreApparentCercle = 2 * iRayonMinAffichage

    End Sub

    Public Sub FixerPosition(ptPosition As Point)
        m_ptMemPos = m_ptPos
        m_ptPos = ptPosition
        If Not m_bPositionInitialisee Then m_ptMemPos = m_ptPos
    End Sub

    Public Sub Dessiner(ByRef dc As Graphics)

        If Not m_bPositionInitialisee Then Exit Sub

        If m_bCercle Or m_imgSprite Is Nothing Then
            Dim iRayon% = m_iMemRayon - m_iLargPinceauCercle \ 2
            Dim iDiam% = m_iMemDiam - m_iLargPinceauCercle - 1
            ' Effacer la position précédente du cercle
            If m_bLaisserTraceCercleGris And m_iMemRayon <> 0 Then _
                dc.DrawEllipse(m_penCercleGris,
                    m_ptMemPos.X - iRayon, m_ptMemPos.Y - iRayon, iDiam, iDiam)
            iRayon = m_iRayon - m_iLargPinceauCercle \ 2
            iDiam = m_iDiametreApparentCercle - m_iLargPinceauCercle - 1
            dc.DrawEllipse(m_penCercle,
                m_ptPos.X - iRayon, m_ptPos.Y - iRayon, iDiam, iDiam)
            Exit Sub
        End If

        If m_bRotationImg Then
            ' Déplacement des coordonnées au centre de l'image en rotation 
            dc.TranslateTransform(m_ptPos.X, m_ptPos.Y)
            ' Définition d'un "container" de transformation
            m_graphicsContainer = dc.BeginContainer()
            dc.RotateTransform(m_rAngleRotImg)
            dc.DrawImage(m_imgSprite,
                -CInt(m_szTailleImg.Width / 2),
                -CInt(m_szTailleImg.Height / 2),
                m_szTailleImg.Width, m_szTailleImg.Height)
            dc.EndContainer(m_graphicsContainer)
            ' Restauration des coordonnées normales
            dc.TranslateTransform(-m_ptPos.X, -m_ptPos.Y)
            Exit Sub
        End If

        dc.DrawImage(m_imgSprite,
            m_ptPos.X - m_iRayon, m_ptPos.Y - m_iRayon,
            m_szTailleImg.Width, m_szTailleImg.Height)

    End Sub

End Class