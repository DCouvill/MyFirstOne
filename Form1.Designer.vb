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
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.lblAppStatus = New System.Windows.Forms.Label()
        Me.lblEnvironment = New System.Windows.Forms.Label()
        Me.checkRefreshPIP = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(31, 174)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(470, 290)
        Me.ListBox1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label1.Location = New System.Drawing.Point(27, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(121, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "PIPexport - v2"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label2.Location = New System.Drawing.Point(28, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(267, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Export PIP MAIN to SQL Server database"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label3.Location = New System.Drawing.Point(28, 54)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(118, 17)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Both FFF && ELSH"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label4.Location = New System.Drawing.Point(30, 71)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(283, 17)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Files are copied from PIP production folders"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label5.Location = New System.Drawing.Point(30, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(199, 17)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Export is done from the copies"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label6.Location = New System.Drawing.Point(30, 105)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(205, 17)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Only certian fields are exported"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label7.Location = New System.Drawing.Point(64, 122)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(135, 17)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "HealthInfo database"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label8.Location = New System.Drawing.Point(58, 139)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(0, 17)
        Me.Label8.TabIndex = 8
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label9.Location = New System.Drawing.Point(64, 139)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(120, 17)
        Me.Label9.TabIndex = 9
        Me.Label9.Text = "tbl_PIPindex table"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(434, 11)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(67, 26)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'lblAppStatus
        '
        Me.lblAppStatus.AutoSize = True
        Me.lblAppStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAppStatus.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lblAppStatus.Location = New System.Drawing.Point(334, 139)
        Me.lblAppStatus.Name = "lblAppStatus"
        Me.lblAppStatus.Size = New System.Drawing.Size(87, 17)
        Me.lblAppStatus.TabIndex = 11
        Me.lblAppStatus.Text = "lblAppStatus"
        '
        'lblEnvironment
        '
        Me.lblEnvironment.AutoSize = True
        Me.lblEnvironment.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEnvironment.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lblEnvironment.Location = New System.Drawing.Point(334, 122)
        Me.lblEnvironment.Name = "lblEnvironment"
        Me.lblEnvironment.Size = New System.Drawing.Size(101, 17)
        Me.lblEnvironment.TabIndex = 12
        Me.lblEnvironment.Text = "lblEnvironment"
        '
        'checkRefreshPIP
        '
        Me.checkRefreshPIP.AutoSize = True
        Me.checkRefreshPIP.Location = New System.Drawing.Point(337, 54)
        Me.checkRefreshPIP.Name = "checkRefreshPIP"
        Me.checkRefreshPIP.Size = New System.Drawing.Size(121, 17)
        Me.checkRefreshPIP.TabIndex = 13
        Me.checkRefreshPIP.Text = "Copy PIP to backup"
        Me.checkRefreshPIP.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(525, 477)
        Me.Controls.Add(Me.checkRefreshPIP)
        Me.Controls.Add(Me.lblEnvironment)
        Me.Controls.Add(Me.lblAppStatus)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListBox1)
        Me.ForeColor = System.Drawing.Color.MediumBlue
        Me.Name = "Form1"
        Me.Text = "PIPexport_v2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lblAppStatus As System.Windows.Forms.Label
    Friend WithEvents lblEnvironment As System.Windows.Forms.Label
    Friend WithEvents checkRefreshPIP As System.Windows.Forms.CheckBox

End Class
