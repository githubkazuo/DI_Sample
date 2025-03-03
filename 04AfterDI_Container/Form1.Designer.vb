<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.btnUserInfo = New System.Windows.Forms.Button()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.lblID = New System.Windows.Forms.Label()
        Me.lblCurrentDb = New System.Windows.Forms.Label()
        Me.btnSwitchDb = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtName
        '
        Me.txtName.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtName.Location = New System.Drawing.Point(114, 177)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(413, 28)
        Me.txtName.TabIndex = 0
        '
        'btnUserInfo
        '
        Me.btnUserInfo.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnUserInfo.Location = New System.Drawing.Point(114, 108)
        Me.btnUserInfo.Name = "btnUserInfo"
        Me.btnUserInfo.Size = New System.Drawing.Size(255, 43)
        Me.btnUserInfo.TabIndex = 1
        Me.btnUserInfo.Text = "ユーザ情報取得"
        Me.btnUserInfo.UseVisualStyleBackColor = True
        '
        'txtID
        '
        Me.txtID.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtID.Location = New System.Drawing.Point(151, 42)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(189, 28)
        Me.txtID.TabIndex = 2
        Me.txtID.Text = "1"
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblID.Location = New System.Drawing.Point(114, 45)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(31, 21)
        Me.lblID.TabIndex = 3
        Me.lblID.Text = "ID"
        '
        'lblCurrentDb
        '
        Me.lblCurrentDb.AutoSize = True
        Me.lblCurrentDb.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblCurrentDb.Location = New System.Drawing.Point(275, 301)
        Me.lblCurrentDb.Name = "lblCurrentDb"
        Me.lblCurrentDb.Size = New System.Drawing.Size(31, 21)
        Me.lblCurrentDb.TabIndex = 7
        Me.lblCurrentDb.Text = "ID"
        '
        'btnSwitchDb
        '
        Me.btnSwitchDb.Font = New System.Drawing.Font("MS UI Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnSwitchDb.Location = New System.Drawing.Point(114, 290)
        Me.btnSwitchDb.Name = "btnSwitchDb"
        Me.btnSwitchDb.Size = New System.Drawing.Size(134, 43)
        Me.btnSwitchDb.TabIndex = 6
        Me.btnSwitchDb.Text = "DB切替"
        Me.btnSwitchDb.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lblCurrentDb)
        Me.Controls.Add(Me.btnSwitchDb)
        Me.Controls.Add(Me.lblID)
        Me.Controls.Add(Me.txtID)
        Me.Controls.Add(Me.btnUserInfo)
        Me.Controls.Add(Me.txtName)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtName As TextBox
    Friend WithEvents btnUserInfo As Button
    Friend WithEvents txtID As TextBox
    Friend WithEvents lblID As Label
    Friend WithEvents lblCurrentDb As Label
    Friend WithEvents btnSwitchDb As Button
End Class
