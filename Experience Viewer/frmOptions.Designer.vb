<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOptions
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lstNoms = New System.Windows.Forms.ListBox()
        Me.lstValeurs = New System.Windows.Forms.ListBox()
        Me.cmdAppliquer = New System.Windows.Forms.Button()
        Me.lstLog = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'lstNoms
        '
        Me.lstNoms.FormattingEnabled = True
        Me.lstNoms.Location = New System.Drawing.Point(12, 12)
        Me.lstNoms.Name = "lstNoms"
        Me.lstNoms.Size = New System.Drawing.Size(284, 602)
        Me.lstNoms.TabIndex = 0
        '
        'lstValeurs
        '
        Me.lstValeurs.FormattingEnabled = True
        Me.lstValeurs.Location = New System.Drawing.Point(302, 12)
        Me.lstValeurs.Name = "lstValeurs"
        Me.lstValeurs.Size = New System.Drawing.Size(534, 602)
        Me.lstValeurs.TabIndex = 1
        '
        'cmdAppliquer
        '
        Me.cmdAppliquer.Location = New System.Drawing.Point(354, 734)
        Me.cmdAppliquer.Name = "cmdAppliquer"
        Me.cmdAppliquer.Size = New System.Drawing.Size(75, 23)
        Me.cmdAppliquer.TabIndex = 2
        Me.cmdAppliquer.Text = "Apply"
        Me.cmdAppliquer.UseVisualStyleBackColor = True
        '
        'lstLog
        '
        Me.lstLog.FormattingEnabled = True
        Me.lstLog.Location = New System.Drawing.Point(15, 620)
        Me.lstLog.Name = "lstLog"
        Me.lstLog.Size = New System.Drawing.Size(821, 108)
        Me.lstLog.TabIndex = 4
        '
        'frmOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(845, 762)
        Me.Controls.Add(Me.lstLog)
        Me.Controls.Add(Me.cmdAppliquer)
        Me.Controls.Add(Me.lstValeurs)
        Me.Controls.Add(Me.lstNoms)
        Me.MinimizeBox = False
        Me.Name = "frmOptions"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "OPTIONS UCI"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lstNoms As System.Windows.Forms.ListBox
    Friend WithEvents lstValeurs As System.Windows.Forms.ListBox
    Friend WithEvents cmdAppliquer As System.Windows.Forms.Button
    Friend WithEvents lstLog As System.Windows.Forms.ListBox
End Class
