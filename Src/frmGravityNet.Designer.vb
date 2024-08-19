<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGravityNet : Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "

Public Sub New()

    MyBase.New()

    'This call is required by the Windows Form Designer.
    InitializeComponent()

    'Add any initialization after the InitializeComponent() call
    InitialiserEcranDeVeille()

End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(disposing As Boolean)
        m_bQuitterBoucleAnimation = True
        If disposing Then
            If Not (components Is Nothing) Then components.Dispose()
            If Not (m_gravity Is Nothing) Then m_gravity.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

'NOTE: The following procedure is required by the Windows Form Designer
'It can be modified using the Windows Form Designer.  
'Do not modify it using the code editor.
Friend WithEvents LblBanniere As System.Windows.Forms.Label
    Friend WithEvents TimerAnimation As System.Windows.Forms.Timer
    Friend WithEvents ListViewPrm As System.Windows.Forms.ListView
        Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
        Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
        Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader

<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.components = New System.ComponentModel.Container
Me.TimerAnimation = New System.Windows.Forms.Timer(Me.components)
Me.LblBanniere = New System.Windows.Forms.Label
Me.ListViewPrm = New System.Windows.Forms.ListView
Me.SuspendLayout()
'
'TimerAnimation
'
Me.TimerAnimation.Enabled = True
Me.TimerAnimation.Interval = 2000
'
'LblBanniere
'
Me.LblBanniere.BackColor = System.Drawing.Color.Transparent
Me.LblBanniere.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.LblBanniere.ForeColor = System.Drawing.Color.Yellow
Me.LblBanniere.Location = New System.Drawing.Point(16, 112)
Me.LblBanniere.Name = "LblBanniere"
Me.LblBanniere.Size = New System.Drawing.Size(232, 24)
Me.LblBanniere.TabIndex = 1
Me.LblBanniere.Text = "Gravity.Net Screen Saver"
Me.LblBanniere.Visible = False
'
'ListViewPrm
'
Me.ListViewPrm.Location = New System.Drawing.Point(8, 8)
Me.ListViewPrm.Name = "ListViewPrm"
Me.ListViewPrm.Size = New System.Drawing.Size(300, 92)
Me.ListViewPrm.TabIndex = 13
Me.ListViewPrm.UseCompatibleStateImageBehavior = False
Me.ListViewPrm.Visible = False
'
'frmGravityNet
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.BackColor = System.Drawing.Color.Gray
Me.CausesValidation = False
Me.ClientSize = New System.Drawing.Size(752, 333)
Me.Controls.Add(Me.ListViewPrm)
Me.Controls.Add(Me.LblBanniere)
Me.ForeColor = System.Drawing.Color.Navy
Me.Name = "frmGravityNet"
Me.ShowInTaskbar = False
Me.Text = "frmGravityNet"
Me.ResumeLayout(False)

End Sub
#End Region
End Class