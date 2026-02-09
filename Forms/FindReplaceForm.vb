Imports System.Drawing
Imports System.Windows.Forms
Imports ScintillaNET

    ''' <summary>
    ''' Find and Replace Dialog (modeless)
    ''' Ported from VB6 frmFind.frm
    ''' </summary>
    Public Class FindReplaceForm
        Inherits Form

        Private txtFind As TextBox
        Private txtReplace As TextBox
        Private chkMatchCase As CheckBox
        Private chkWholeWord As CheckBox
        Private chkWrapAround As CheckBox
        Private chkRegExp As CheckBox
        Private btnFindNext As Button
        Private btnFindPrev As Button
        Private btnReplace As Button
        Private btnReplaceAll As Button
        Private btnClose As Button
        Private lblStatus As Label
        Private lblFind As Label
        Private lblReplace As Label

        Private _editor As Scintilla
        Private _mainForm As Form
        Private _showReplace As Boolean

        Public Sub New(mainForm As Form, editor As Scintilla, showReplace As Boolean)
            _editor = editor
            _mainForm = mainForm
            _showReplace = showReplace
            InitializeComponent()
        End Sub

        Private Sub InitializeComponent()
            Me.Text = If(_showReplace, "Find and Replace", "Find")
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.StartPosition = FormStartPosition.CenterParent
            Me.ClientSize = New Size(490, 240)
            Me.Font = New Font("Segoe UI", 9)
            Me.TopMost = True

            ' Find label + textbox
            lblFind = New Label() With {.Text = "&Find:", .Location = New Point(12, 14), .AutoSize = True}
            Me.Controls.Add(lblFind)

            txtFind = New TextBox() With {
                .Location = New Point(88, 11), .Size = New Size(272, 23),
                .Font = New Font("Segoe UI", 9.75F)
            }
            AddHandler txtFind.KeyDown, AddressOf TxtFind_KeyDown
            Me.Controls.Add(txtFind)

            ' Replace label + textbox
            lblReplace = New Label() With {.Text = "&Replace:", .Location = New Point(12, 46), .AutoSize = True}
            Me.Controls.Add(lblReplace)

            txtReplace = New TextBox() With {
                .Location = New Point(88, 43), .Size = New Size(272, 23),
                .Font = New Font("Segoe UI", 9.75F)
            }
            Me.Controls.Add(txtReplace)

            If Not _showReplace Then
                lblReplace.Visible = False
                txtReplace.Visible = False
            End If

            ' Checkboxes
            chkMatchCase = New CheckBox() With {.Text = "Match &case", .Location = New Point(16, 80), .AutoSize = True}
            chkWholeWord = New CheckBox() With {.Text = "&Whole word", .Location = New Point(16, 104), .AutoSize = True}
            chkWrapAround = New CheckBox() With {.Text = "Wrap aro&und", .Location = New Point(16, 128), .AutoSize = True, .Checked = True}
            chkRegExp = New CheckBox() With {.Text = "Regular e&xpression", .Location = New Point(16, 152), .AutoSize = True}
            Me.Controls.AddRange({chkMatchCase, chkWholeWord, chkWrapAround, chkRegExp})

            ' Buttons
            Dim bx As Integer = 376
            btnFindNext = New Button() With {.Text = "Find &Next", .Location = New Point(bx, 10), .Size = New Size(100, 28)}
            AddHandler btnFindNext.Click, AddressOf BtnFindNext_Click
            Me.Controls.Add(btnFindNext)

            btnFindPrev = New Button() With {.Text = "Find &Prev", .Location = New Point(bx, 42), .Size = New Size(100, 28)}
            AddHandler btnFindPrev.Click, AddressOf BtnFindPrev_Click
            Me.Controls.Add(btnFindPrev)

            btnReplace = New Button() With {.Text = "&Replace", .Location = New Point(bx, 80), .Size = New Size(100, 28)}
            AddHandler btnReplace.Click, AddressOf BtnReplace_Click
            btnReplace.Enabled = _showReplace
            Me.Controls.Add(btnReplace)

            btnReplaceAll = New Button() With {.Text = "Replace &All", .Location = New Point(bx, 112), .Size = New Size(100, 28)}
            AddHandler btnReplaceAll.Click, AddressOf BtnReplaceAll_Click
            btnReplaceAll.Enabled = _showReplace
            Me.Controls.Add(btnReplaceAll)

            btnClose = New Button() With {.Text = "Close", .Location = New Point(bx, 152), .Size = New Size(100, 28)}
            AddHandler btnClose.Click, Sub() Me.Close()
            Me.Controls.Add(btnClose)

            ' Status label
            lblStatus = New Label() With {
                .Location = New Point(12, 185),
                .Size = New Size(350, 24),
                .ForeColor = Color.FromArgb(0, 0, 128),
                .Font = New Font("Segoe UI", 9)
            }
            Me.Controls.Add(lblStatus)

            Me.AcceptButton = btnFindNext
            Me.CancelButton = btnClose
        End Sub

        Protected Overrides Sub OnActivated(e As EventArgs)
            MyBase.OnActivated(e)
            txtFind.Focus()
            txtFind.SelectAll()
            lblStatus.Text = ""
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            ' Hide instead of destroy, but only if user clicked close/X
            If e.CloseReason = CloseReason.UserClosing Then
                e.Cancel = True
                Me.Hide()
            End If
            MyBase.OnFormClosing(e)
        End Sub

        Private Function GetSearchFlags() As SearchFlags
            Dim flags As SearchFlags = SearchFlags.None
            If chkMatchCase.Checked Then flags = flags Or SearchFlags.MatchCase
            If chkWholeWord.Checked Then flags = flags Or SearchFlags.WholeWord
            If chkRegExp.Checked Then flags = flags Or SearchFlags.Regex
            Return flags
        End Function

        Private Function DoFind(forward As Boolean) As Boolean
            Dim text = txtFind.Text
            If String.IsNullOrEmpty(text) Then
                lblStatus.Text = "Nothing to find."
                Return False
            End If

            ' Update MainForm's last find text
            If TypeOf _mainForm Is MainForm Then
                DirectCast(_mainForm, MainForm).LastFindText = text
            End If

            _editor.SearchFlags = GetSearchFlags()

            If forward Then
                _editor.TargetStart = _editor.SelectionEnd
                _editor.TargetEnd = _editor.TextLength
            Else
                _editor.TargetStart = _editor.SelectionStart
                _editor.TargetEnd = 0
            End If

            Dim pos = _editor.SearchInTarget(text)

            If pos < 0 AndAlso chkWrapAround.Checked Then
                ' Wrap
                If forward Then
                    _editor.TargetStart = 0
                    _editor.TargetEnd = _editor.TextLength
                Else
                    _editor.TargetStart = _editor.TextLength
                    _editor.TargetEnd = 0
                End If
                pos = _editor.SearchInTarget(text)
                If pos >= 0 Then
                    _editor.SetSel(_editor.TargetStart, _editor.TargetEnd)
                    _editor.ScrollCaret()
                    lblStatus.Text = "Found (wrapped around)."
                    Return True
                End If
            End If

            If pos >= 0 Then
                _editor.SetSel(_editor.TargetStart, _editor.TargetEnd)
                _editor.ScrollCaret()
                lblStatus.Text = "Found."
                Return True
            Else
                lblStatus.Text = "Not found."
                Return False
            End If
        End Function

        Private Sub BtnFindNext_Click(sender As Object, e As EventArgs)
            DoFind(True)
        End Sub

        Private Sub BtnFindPrev_Click(sender As Object, e As EventArgs)
            DoFind(False)
        End Sub

        Private Sub BtnReplace_Click(sender As Object, e As EventArgs)
            If String.IsNullOrEmpty(txtFind.Text) Then Return

            ' If current selection matches the find text, replace it
            If Not String.IsNullOrEmpty(_editor.SelectedText) Then
                Dim match As Boolean
                If chkMatchCase.Checked Then
                    match = String.Equals(_editor.SelectedText, txtFind.Text, StringComparison.Ordinal)
                Else
                    match = String.Equals(_editor.SelectedText, txtFind.Text, StringComparison.OrdinalIgnoreCase)
                End If
                If match Then
                    _editor.ReplaceSelection(txtReplace.Text)
                End If
            End If

            ' Find next
            DoFind(True)
        End Sub

        Private Sub BtnReplaceAll_Click(sender As Object, e As EventArgs)
            If String.IsNullOrEmpty(txtFind.Text) Then Return

            Dim count As Integer = 0
            _editor.SearchFlags = GetSearchFlags()

            ' Start from beginning
            _editor.TargetStart = 0
            _editor.TargetEnd = _editor.TextLength

            While _editor.SearchInTarget(txtFind.Text) >= 0
                _editor.ReplaceTarget(txtReplace.Text)
                count += 1
                ' Continue searching after the replacement
                _editor.TargetStart = _editor.TargetEnd
                _editor.TargetEnd = _editor.TextLength
            End While

            lblStatus.Text = $"{count} replacement(s) made."
        End Sub

        Private Sub TxtFind_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter Then
                BtnFindNext_Click(Nothing, Nothing)
                e.Handled = True
                e.SuppressKeyPress = True
            End If
        End Sub

        ''' <summary>Pre-populate the find text and show the form</summary>
        Public Sub ShowWithText(text As String, showReplace As Boolean)
            If Not String.IsNullOrEmpty(text) Then
                txtFind.Text = text
            End If
            If showReplace Then
                lblReplace.Visible = True
                txtReplace.Visible = True
                btnReplace.Enabled = True
                btnReplaceAll.Enabled = True
                Me.Text = "Find and Replace"
            Else
                lblReplace.Visible = False
                txtReplace.Visible = False
                btnReplace.Enabled = False
                btnReplaceAll.Enabled = False
                Me.Text = "Find"
            End If
            _showReplace = showReplace
            Me.Show()
            Me.BringToFront()
            txtFind.Focus()
            txtFind.SelectAll()
        End Sub
    End Class
