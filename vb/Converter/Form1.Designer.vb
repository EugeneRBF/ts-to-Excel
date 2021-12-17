<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Converter
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.parserBtn = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LangsList = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'parserBtn
        '
        Me.parserBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.parserBtn.Location = New System.Drawing.Point(47, 45)
        Me.parserBtn.Name = "parserBtn"
        Me.parserBtn.Size = New System.Drawing.Size(197, 47)
        Me.parserBtn.TabIndex = 0
        Me.parserBtn.Text = "Парсинг .ts"
        Me.parserBtn.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(289, 58)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(361, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Выберите один или несколько файлов .ts"
        '
        'LangsList
        '
        Me.LangsList.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.LangsList.FormattingEnabled = True
        Me.LangsList.ItemHeight = 20
        Me.LangsList.Location = New System.Drawing.Point(47, 134)
        Me.LangsList.Name = "LangsList"
        Me.LangsList.Size = New System.Drawing.Size(197, 204)
        Me.LangsList.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label2.Location = New System.Drawing.Point(290, 134)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(268, 20)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Список обнаруженных локалей"
        '
        'Converter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(731, 440)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LangsList)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.parserBtn)
        Me.Name = "Converter"
        Me.Text = "Converter TypeScript to Excel"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents parserBtn As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents LangsList As ListBox
    Friend WithEvents Label2 As Label
End Class
