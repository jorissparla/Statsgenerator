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
        Me.components = New System.ComponentModel.Container()
        Me.BtnRun = New System.Windows.Forms.Button()
        Me.LbRanges = New System.Windows.Forms.ListBox()
        Me.lbReps = New System.Windows.Forms.ListBox()
        Me.LblStatus = New System.Windows.Forms.Label()
        Me.lbOwnerGroups = New System.Windows.Forms.ListBox()
        Me.lbRegions = New System.Windows.Forms.ListBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CboManagers = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.CboPages = New System.Windows.Forms.ComboBox()
        Me.LbQuerySets = New System.Windows.Forms.ListBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'BtnRun
        '
        Me.BtnRun.Location = New System.Drawing.Point(36, 403)
        Me.BtnRun.Name = "BtnRun"
        Me.BtnRun.Size = New System.Drawing.Size(99, 27)
        Me.BtnRun.TabIndex = 1
        Me.BtnRun.Text = "Run"
        Me.BtnRun.UseVisualStyleBackColor = True
        '
        'LbRanges
        '
        Me.LbRanges.FormattingEnabled = True
        Me.LbRanges.Location = New System.Drawing.Point(278, 53)
        Me.LbRanges.Name = "LbRanges"
        Me.LbRanges.Size = New System.Drawing.Size(188, 329)
        Me.LbRanges.TabIndex = 2
        '
        'lbReps
        '
        Me.lbReps.FormattingEnabled = True
        Me.lbReps.Location = New System.Drawing.Point(36, 53)
        Me.lbReps.Name = "lbReps"
        Me.lbReps.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lbReps.Size = New System.Drawing.Size(220, 329)
        Me.lbReps.TabIndex = 3
        '
        'LblStatus
        '
        Me.LblStatus.AutoSize = True
        Me.LblStatus.Location = New System.Drawing.Point(168, 417)
        Me.LblStatus.Name = "LblStatus"
        Me.LblStatus.Size = New System.Drawing.Size(37, 13)
        Me.LblStatus.TabIndex = 4
        Me.LblStatus.Text = "Status"
        '
        'lbOwnerGroups
        '
        Me.lbOwnerGroups.FormattingEnabled = True
        Me.lbOwnerGroups.Location = New System.Drawing.Point(484, 53)
        Me.lbOwnerGroups.Name = "lbOwnerGroups"
        Me.lbOwnerGroups.Size = New System.Drawing.Size(188, 329)
        Me.lbOwnerGroups.TabIndex = 5
        '
        'lbRegions
        '
        Me.lbRegions.FormattingEnabled = True
        Me.lbRegions.Location = New System.Drawing.Point(687, 53)
        Me.lbRegions.Name = "lbRegions"
        Me.lbRegions.Size = New System.Drawing.Size(188, 329)
        Me.lbRegions.TabIndex = 6
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(36, 469)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(110, 27)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Clear Selections"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CboManagers
        '
        Me.CboManagers.FormattingEnabled = True
        Me.CboManagers.Items.AddRange(New Object() {"(All)", "Maribel Aguilella", "Marina Ter Haar", "Owen Beer", "Joris Sparla"})
        Me.CboManagers.Location = New System.Drawing.Point(36, 26)
        Me.CboManagers.Name = "CboManagers"
        Me.CboManagers.Size = New System.Drawing.Size(140, 21)
        Me.CboManagers.TabIndex = 8
        Me.CboManagers.Text = "(All)"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(33, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Manager"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(36, 436)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(110, 27)
        Me.Button2.TabIndex = 10
        Me.Button2.Text = "Select all"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'CboPages
        '
        Me.CboPages.FormattingEnabled = True
        Me.CboPages.Items.AddRange(New Object() {"(All)", "Maribel Aguilella", "Marina Ter Haar", "Owen Beer", "Joris Sparla"})
        Me.CboPages.Location = New System.Drawing.Point(171, 442)
        Me.CboPages.Name = "CboPages"
        Me.CboPages.Size = New System.Drawing.Size(140, 21)
        Me.CboPages.TabIndex = 11
        Me.CboPages.Text = "(All)"
        '
        'LbQuerySets
        '
        Me.LbQuerySets.FormattingEnabled = True
        Me.LbQuerySets.Location = New System.Drawing.Point(355, 401)
        Me.LbQuerySets.Name = "LbQuerySets"
        Me.LbQuerySets.Size = New System.Drawing.Size(317, 108)
        Me.LbQuerySets.TabIndex = 12
        '
        'ToolTip1
        '
        Me.ToolTip1.IsBalloon = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(897, 524)
        Me.Controls.Add(Me.LbQuerySets)
        Me.Controls.Add(Me.CboPages)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CboManagers)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.lbRegions)
        Me.Controls.Add(Me.lbOwnerGroups)
        Me.Controls.Add(Me.LblStatus)
        Me.Controls.Add(Me.lbReps)
        Me.Controls.Add(Me.LbRanges)
        Me.Controls.Add(Me.BtnRun)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnRun As System.Windows.Forms.Button
    Friend WithEvents LbRanges As System.Windows.Forms.ListBox
    Friend WithEvents lbReps As System.Windows.Forms.ListBox
    Friend WithEvents LblStatus As System.Windows.Forms.Label
    Friend WithEvents lbOwnerGroups As System.Windows.Forms.ListBox
    Friend WithEvents lbRegions As System.Windows.Forms.ListBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents CboManagers As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents CboPages As System.Windows.Forms.ComboBox
    Friend WithEvents LbQuerySets As System.Windows.Forms.ListBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip

End Class
