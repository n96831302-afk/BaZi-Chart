<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class f0Main
  Inherits System.Windows.Forms.Form

  'Form 覆寫 Dispose 以清除元件清單。
  <System.Diagnostics.DebuggerNonUserCode()>
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    Try
      If disposing AndAlso components IsNot Nothing Then
        components.Dispose()
      End If
    Finally
      MyBase.Dispose(disposing)
    End Try
  End Sub

  '為 Windows Form 設計工具的必要項
  Private components As System.ComponentModel.IContainer

  '注意: 以下為 Windows Form 設計工具所需的程序
  '可以使用 Windows Form 設計工具進行修改。
  '請不要使用程式碼編輯器進行修改。
  <System.Diagnostics.DebuggerStepThrough()>
  Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(f0Main))
        Me.btn八字排盤1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btn八字排盤1
        '
        Me.btn八字排盤1.Location = New System.Drawing.Point(0, 12)
        Me.btn八字排盤1.Name = "btn八字排盤1"
        Me.btn八字排盤1.Size = New System.Drawing.Size(71, 30)
        Me.btn八字排盤1.TabIndex = 11
        Me.btn八字排盤1.Text = "八字盤1"
        Me.btn八字排盤1.UseVisualStyleBackColor = True
        '
        'f0Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(221, 49)
        Me.Controls.Add(Me.btn八字排盤1)
        Me.Font = New System.Drawing.Font("微軟正黑體", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "f0Main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SKY"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btn八字排盤1 As Button
End Class
