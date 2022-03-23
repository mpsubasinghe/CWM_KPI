<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Dat = New System.Windows.Forms.DateTimePicker()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(23, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 13)
        Me.Label3.TabIndex = 40
        Me.Label3.Text = "Month"
        '
        'Dat
        '
        Me.Dat.Location = New System.Drawing.Point(79, 73)
        Me.Dat.Name = "Dat"
        Me.Dat.Size = New System.Drawing.Size(269, 20)
        Me.Dat.TabIndex = 39
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(273, 99)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 41
        Me.Button1.Text = "View"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 147)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(351, 23)
        Me.ProgressBar1.TabIndex = 42
        '
        'ProgressBar2
        '
        Me.ProgressBar2.Location = New System.Drawing.Point(12, 128)
        Me.ProgressBar2.Name = "ProgressBar2"
        Me.ProgressBar2.Size = New System.Drawing.Size(351, 13)
        Me.ProgressBar2.TabIndex = 43
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MV Boli", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(94, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(197, 34)
        Me.Label1.TabIndex = 44
        Me.Label1.Text = "KPI Document"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(389, 194)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ProgressBar2)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Dat)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "KPI"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Dat As System.Windows.Forms.DateTimePicker
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
