﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class StartForm
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
        Me.components = New System.ComponentModel.Container
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.RunBtn = New System.Windows.Forms.Button
        Me.ExportBtn = New System.Windows.Forms.Button
        Me.ExitBtb = New System.Windows.Forms.Button
        Me.OpenFileDialog2 = New System.Windows.Forms.OpenFileDialog
        Me.MdbSaveFileDialog = New System.Windows.Forms.SaveFileDialog
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.AcceptsReturn = True
        Me.TextBox1.AllowDrop = True
        Me.TextBox1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Font = New System.Drawing.Font("Stencil", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(26, 35)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(528, 49)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = "Welcome to ChCal - The Main Calibration Program"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(81, 118)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(434, 178)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Please Select a Run Option"
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(44, 122)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(190, 24)
        Me.RadioButton3.TabIndex = 2
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "Process only All Stocks"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(44, 82)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(336, 24)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Process Out of Base Stocks then All Stocks"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(44, 38)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(258, 24)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Process only Out of Base Stocks"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RunBtn
        '
        Me.RunBtn.BackColor = System.Drawing.Color.Olive
        Me.RunBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunBtn.Location = New System.Drawing.Point(81, 336)
        Me.RunBtn.Name = "RunBtn"
        Me.RunBtn.Size = New System.Drawing.Size(84, 40)
        Me.RunBtn.TabIndex = 2
        Me.RunBtn.Text = "RUN"
        Me.RunBtn.UseVisualStyleBackColor = False
        '
        'ExportBtn
        '
        Me.ExportBtn.BackColor = System.Drawing.Color.Goldenrod
        Me.ExportBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExportBtn.Location = New System.Drawing.Point(431, 336)
        Me.ExportBtn.Name = "ExportBtn"
        Me.ExportBtn.Size = New System.Drawing.Size(84, 40)
        Me.ExportBtn.TabIndex = 3
        Me.ExportBtn.Text = "Export"
        Me.ExportBtn.UseVisualStyleBackColor = False
        '
        'ExitBtb
        '
        Me.ExitBtb.BackColor = System.Drawing.Color.Chocolate
        Me.ExitBtb.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExitBtb.Location = New System.Drawing.Point(256, 336)
        Me.ExitBtb.Name = "ExitBtb"
        Me.ExitBtb.Size = New System.Drawing.Size(84, 40)
        Me.ExitBtb.TabIndex = 4
        Me.ExitBtb.Text = "Exit"
        Me.ExitBtb.UseVisualStyleBackColor = False
        '
        'OpenFileDialog2
        '
        Me.OpenFileDialog2.FileName = "OpenFileDialog2"
        '
        'StartForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ClientSize = New System.Drawing.Size(608, 411)
        Me.Controls.Add(Me.ExitBtb)
        Me.Controls.Add(Me.ExportBtn)
        Me.Controls.Add(Me.RunBtn)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "StartForm"
        Me.Text = "StartForm"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RunBtn As System.Windows.Forms.Button
    Friend WithEvents ExportBtn As System.Windows.Forms.Button
    Friend WithEvents ExitBtb As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog2 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents MdbSaveFileDialog As System.Windows.Forms.SaveFileDialog
End Class
