Imports System.Drawing
Imports System.Windows.Forms


    Public Class GoToLineForm
        Inherits Form

        Private lblPrompt As Label
        Private txtLineNumber As TextBox
        Private WithEvents btnOK As Button
        Private WithEvents btnCancel As Button

        Private _lineNumber As Integer = 1

        Public ReadOnly Property LineNumber As Integer
            Get
                Return _lineNumber
            End Get
        End Property

        Public Sub New(currentLine As Integer)
            _lineNumber = currentLine
            InitializeComponent()
        End Sub

        Private Sub InitializeComponent()
            Me.Text = "Go To Line"
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.StartPosition = FormStartPosition.CenterParent
            Me.Size = New Size(320, 160)
            Me.Font = New Font("Segoe UI", 9)
            Me.AcceptButton = Nothing
            Me.KeyPreview = True

            lblPrompt = New Label() With {
                .Text = "Line number (1 - ...):",
                .Location = New Point(12, 16),
                .Size = New Size(280, 20),
                .AutoSize = False
            }

            txtLineNumber = New TextBox() With {
                .Location = New Point(12, 40),
                .Size = New Size(280, 24),
                .Text = _lineNumber.ToString(),
                .Font = New Font("Consolas", 11)
            }

            btnOK = New Button() With {
                .Text = "Go",
                .Location = New Point(116, 80),
                .Size = New Size(85, 30),
                .DialogResult = DialogResult.OK
            }

            btnCancel = New Button() With {
                .Text = "Cancel",
                .Location = New Point(207, 80),
                .Size = New Size(85, 30),
                .DialogResult = DialogResult.Cancel
            }

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel

            Me.Controls.AddRange({lblPrompt, txtLineNumber, btnOK, btnCancel})

            AddHandler Me.Shown, Sub()
                                     txtLineNumber.SelectAll()
                                     txtLineNumber.Focus()
                                 End Sub
        End Sub

        Private Sub BtnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
            Dim num As Integer
            If Integer.TryParse(txtLineNumber.Text.Trim(), num) AndAlso num >= 1 Then
                _lineNumber = num
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show("Please enter a valid line number.", "Go To Line",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtLineNumber.SelectAll()
                txtLineNumber.Focus()
            End If
        End Sub

    End Class

