<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class fVisor
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(fVisor))
        Me.oMediaPlayer3 = New AxWMPLib.AxWindowsMediaPlayer()
        CType(Me.oMediaPlayer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'oMediaPlayer3
        '
        Me.oMediaPlayer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.oMediaPlayer3.Enabled = True
        Me.oMediaPlayer3.Location = New System.Drawing.Point(0, 0)
        Me.oMediaPlayer3.Name = "oMediaPlayer3"
        Me.oMediaPlayer3.OcxState = CType(resources.GetObject("oMediaPlayer3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.oMediaPlayer3.Size = New System.Drawing.Size(784, 561)
        Me.oMediaPlayer3.TabIndex = 0
        '
        'fVisor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 561)
        Me.Controls.Add(Me.oMediaPlayer3)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "fVisor"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.oMediaPlayer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents oMediaPlayer3 As AxWMPLib.AxWindowsMediaPlayer
End Class
