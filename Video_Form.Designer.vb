<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Video_Form
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Video_Form))
        Me.MediaPlayer3 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.oImg_Logo1 = New System.Windows.Forms.PictureBox()
        CType(Me.MediaPlayer3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.oImg_Logo1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MediaPlayer3
        '
        Me.MediaPlayer3.Enabled = True
        Me.MediaPlayer3.Location = New System.Drawing.Point(-26, 92)
        Me.MediaPlayer3.Name = "MediaPlayer3"
        Me.MediaPlayer3.OcxState = CType(resources.GetObject("MediaPlayer3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MediaPlayer3.Size = New System.Drawing.Size(343, 277)
        Me.MediaPlayer3.TabIndex = 43
        '
        'oImg_Logo1
        '
        Me.oImg_Logo1.BackColor = System.Drawing.Color.Transparent
        Me.oImg_Logo1.Location = New System.Drawing.Point(467, 191)
        Me.oImg_Logo1.Name = "oImg_Logo1"
        Me.oImg_Logo1.Size = New System.Drawing.Size(343, 277)
        Me.oImg_Logo1.TabIndex = 44
        Me.oImg_Logo1.TabStop = False
        Me.oImg_Logo1.Visible = False
        '
        'Video_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(784, 561)
        Me.Controls.Add(Me.oImg_Logo1)
        Me.Controls.Add(Me.MediaPlayer3)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Video_Form"
        Me.Text = "Form1"
        CType(Me.MediaPlayer3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.oImg_Logo1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents oImg_Logo1 As PictureBox
    Friend WithEvents MediaPlayer3 As AxWMPLib.AxWindowsMediaPlayer
End Class
