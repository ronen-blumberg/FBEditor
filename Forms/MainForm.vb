Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports ScintillaNET


    Public Class MainForm
        Inherits Form

        ' ---- Controls ----
        Private WithEvents menuMain As MenuStrip
        Private WithEvents toolMain As ToolStrip
        Private toolDebug As ToolStrip
        Private WithEvents statusMain As StatusStrip
        Private lblStatus As ToolStripStatusLabel
        Private lblPosition As ToolStripStatusLabel
        Private lblInsOvr As ToolStripStatusLabel
        Private lblLineCount As ToolStripStatusLabel
        Private lblEncoding As ToolStripStatusLabel
        Private lblLineEnding As ToolStripStatusLabel
        Private lblDebugState As ToolStripStatusLabel

        Private splitMain As SplitContainer      ' Left panel | Editor+Output
        Private splitEditor As SplitContainer    ' Editor | Output
        Private splitLeft As SplitContainer      ' Project tree | Outline tree

        Private WithEvents tvProject As TreeView
        Private WithEvents tvOutline As TreeView
        Private WithEvents cboOpenFiles As ComboBox
        Private WithEvents scintilla As Scintilla

        ' Output panel
        Private tabOutput As TabControl
        Private tabPageOutput As TabPage
        Private tabPageAIChat As TabPage
        Private tabPageDebugOutput As TabPage
        Private tabPageLocals As TabPage
        Private tabPageCallStack As TabPage
        Private WithEvents txtOutput As TextBox
        Private WithEvents txtAIChat As RichTextBox
        Private WithEvents txtAIInput As TextBox
        Private WithEvents btnAISend As Button
        Private WithEvents btnAISendCode As Button
        Private WithEvents btnAIInsertCode As Button
        Private WithEvents btnAICopyReply As Button
        Private WithEvents btnAIReplaceAll As Button
        Private WithEvents btnAINewFile As Button
        Private WithEvents btnAIClearChat As Button
        Private pnlAIActions As FlowLayoutPanel
        Private pnlAIInput As Panel
        Private WithEvents txtDebugOutput As TextBox
        Private WithEvents txtGDBCommand As TextBox
        Private WithEvents lvLocals As ListView
        Private WithEvents lvWatch As ListView
        Private WithEvents lvCallStack As ListView

        ' Menu items (for checked state)
        Private mnuViewLineNumbers As ToolStripMenuItem
        Private mnuViewWordWrap As ToolStripMenuItem
        Private mnuViewWhitespace As ToolStripMenuItem
        Private mnuViewIndentGuides As ToolStripMenuItem
        Private mnuViewProjectExplorer As ToolStripMenuItem
        Private mnuViewOutputPanel As ToolStripMenuItem
        Private mnuViewCodeFolding As ToolStripMenuItem
        Private mnuEditAutoCompleteEnable As ToolStripMenuItem
        Private mnuEncANSI As ToolStripMenuItem
        Private mnuEncUTF8 As ToolStripMenuItem
        Private mnuEncUTF8BOM As ToolStripMenuItem
        Private mnuEncDefaultANSI As ToolStripMenuItem
        Private mnuEncDefaultUTF8 As ToolStripMenuItem

        ' Debug toolbar buttons (for enable/disable state management)
        Private btnDebugStart As ToolStripButton
        Private btnDebugStop As ToolStripButton
        Private btnDebugPause As ToolStripButton
        Private btnDebugStepOver As ToolStripButton
        Private btnDebugStepInto As ToolStripButton
        Private btnDebugStepOut As ToolStripButton

        ' ---- State ----
        Private _files As New List(Of OpenFileInfo)()
        Private _activeFile As Integer = -1
        Private _switchingFile As Boolean = False
        Private _lastFindText As String = ""
        Private _aiChat As New AIChatManager()
        Private _autoCompleteList As String = ""
        Private _baseAutoCompleteList As String = ""
        Private _acTimer As Timer
        Private _foldTimer As Timer
        Private _initialized As Boolean = False
        Private _findReplaceForm As FindReplaceForm = Nothing  ' BUG FIX: reuse form, prevent leak

        ' ---- Debugger ----
        Private WithEvents _debugger As New GDBDebugger()
        Private _compilerErrors As New List(Of CompilerError)()

        ' ---- Marker constants ----
        Private Const MARKER_BOOKMARK As Integer = 1
        Private Const MARKER_BREAKPOINT As Integer = 2
        Private Const MARKER_DEBUGLINE As Integer = 3

        ' ---- Recent files ----
        Private _recentMenuItems As New List(Of ToolStripMenuItem)()
        Private _mnuRecentParent As ToolStripMenuItem

        Public Sub New()
            InitializeApp()
            _baseAutoCompleteList = SyntaxConfig.GetAutoCompleteList()
            _autoCompleteList = _baseAutoCompleteList

            _acTimer = New Timer() With {.Interval = 600, .Enabled = False}
            AddHandler _acTimer.Tick, Sub(s, ev)
                                         _acTimer.Stop()
                                         RebuildAutoCompleteList()
                                     End Sub

            _foldTimer = New Timer() With {.Interval = 500, .Enabled = False}
            AddHandler _foldTimer.Tick, Sub(s, ev)
                                            _foldTimer.Stop()
                                            If Settings.ShowFolding Then FoldingManager.UpdateFoldLevels(scintilla)
                                        End Sub

            InitializeComponent()
            ApplyCurrentTheme()
            SetupScintilla()

            If String.IsNullOrEmpty(Build.FBCPath) Then Build.FBCPath = FindFBCPath()
            If String.IsNullOrEmpty(Build.GDBPath) Then Build.GDBPath = GDBDebugger.FindGDBPath()

            DoNewFile()
            UpdateRecentFilesMenu()
            UpdateDebugUI()
            _initialized = True
        End Sub

        Private Sub InitializeComponent()
            Me.SuspendLayout()
            Me.AutoScaleMode = AutoScaleMode.Dpi
            Me.Text = APP_NAME
            Me.Size = New Size(1400, 900)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.WindowState = FormWindowState.Maximized
            Me.KeyPreview = True
            Me.Font = New Font("Segoe UI", 9)
            Me.AllowDrop = True    ' BUG FIX: enable drag & drop

            ' ---- Menu ----
            menuMain = New MenuStrip()
            BuildMenus()
            Me.MainMenuStrip = menuMain

            ' ---- Main Toolbar ----
            toolMain = New ToolStrip() With {.ImageScalingSize = New Size(16, 16), .GripStyle = ToolStripGripStyle.Hidden}
            BuildToolbar()

            ' ---- Debug Toolbar ----
            toolDebug = New ToolStrip() With {.ImageScalingSize = New Size(16, 16), .GripStyle = ToolStripGripStyle.Hidden}
            BuildDebugToolbar()

            ' ---- Status Bar ----
            statusMain = New StatusStrip()
            lblStatus = New ToolStripStatusLabel("Ready") With {.Spring = True, .TextAlign = ContentAlignment.MiddleLeft}
            lblPosition = New ToolStripStatusLabel("Ln: 1  Col: 1") With {.AutoSize = False, .Width = 120}
            lblInsOvr = New ToolStripStatusLabel("INS") With {.AutoSize = False, .Width = 50}
            lblLineCount = New ToolStripStatusLabel("") With {.AutoSize = False, .Width = 80}
            lblEncoding = New ToolStripStatusLabel("UTF-8") With {.AutoSize = False, .Width = 80}
            lblLineEnding = New ToolStripStatusLabel("CRLF") With {.AutoSize = False, .Width = 50}
            lblDebugState = New ToolStripStatusLabel("") With {.AutoSize = False, .Width = 100, .ForeColor = Color.DodgerBlue}
            statusMain.Items.AddRange(New ToolStripItem() {lblStatus, lblDebugState, lblPosition, lblInsOvr, lblLineCount, lblEncoding, lblLineEnding})

            ' ---- Left Panel ----
            tvProject = New TreeView() With {.Dock = DockStyle.Fill, .HideSelection = False, .ShowLines = True, .ShowPlusMinus = True, .ShowRootLines = True}
            tvOutline = New TreeView() With {.Dock = DockStyle.Fill, .HideSelection = False, .ShowLines = True, .ShowPlusMinus = True, .ShowRootLines = True}

            splitLeft = New SplitContainer() With {.Dock = DockStyle.Fill, .Orientation = Orientation.Horizontal, .SplitterWidth = 4, .SplitterDistance = 200}
            Dim lblFiles As New Label() With {.Text = "Open Files", .Dock = DockStyle.Top, .Height = 20, .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(4, 0, 0, 0)}
            Dim lblOutline As New Label() With {.Text = "Code Outline", .Dock = DockStyle.Top, .Height = 20, .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(4, 0, 0, 0)}
            splitLeft.Panel1.Controls.Add(tvProject)
            splitLeft.Panel1.Controls.Add(lblFiles)
            splitLeft.Panel2.Controls.Add(tvOutline)
            splitLeft.Panel2.Controls.Add(lblOutline)

            ' ---- Editor Area ----
            cboOpenFiles = New ComboBox() With {.Dock = DockStyle.Top, .DropDownStyle = ComboBoxStyle.DropDownList, .Font = New Font("Segoe UI", 9)}
            scintilla = New Scintilla() With {.Dock = DockStyle.Fill}

            Dim pnlEditor As New Panel() With {.Dock = DockStyle.Fill}
            pnlEditor.Controls.Add(scintilla)
            pnlEditor.Controls.Add(cboOpenFiles)

            ' ---- Output Panel ----
            BuildOutputPanel()

            ' ---- Main Splitters ----
            splitEditor = New SplitContainer() With {.Dock = DockStyle.Fill, .Orientation = Orientation.Horizontal, .SplitterWidth = 4, .SplitterDistance = 500}
            splitEditor.Panel1.Controls.Add(pnlEditor)
            splitEditor.Panel2.Controls.Add(tabOutput)

            splitMain = New SplitContainer() With {.Dock = DockStyle.Fill, .SplitterWidth = 4, .SplitterDistance = 250}
            splitMain.Panel1.Controls.Add(splitLeft)
            splitMain.Panel2.Controls.Add(splitEditor)

            ' ---- Add to form (order matters!) ----
            Me.Controls.Add(splitMain)
            Me.Controls.Add(statusMain)
            Me.Controls.Add(toolDebug)
            Me.Controls.Add(toolMain)
            Me.Controls.Add(menuMain)

            Me.ResumeLayout()
        End Sub

        Private Sub BuildOutputPanel()
            tabOutput = New TabControl() With {.Dock = DockStyle.Fill}

            ' Output tab
            tabPageOutput = New TabPage("Output")
            txtOutput = New TextBox() With {
                .Dock = DockStyle.Fill, .Multiline = True, .ScrollBars = ScrollBars.Both,
                .ReadOnly = True, .Font = New Font("Consolas", 10), .WordWrap = False
            }
            AddHandler txtOutput.MouseDoubleClick, AddressOf TxtOutput_DoubleClick
            tabPageOutput.Controls.Add(txtOutput)

            ' Debug Output tab
            tabPageDebugOutput = New TabPage("Debug Output")
            txtDebugOutput = New TextBox() With {
                .Dock = DockStyle.Fill, .Multiline = True, .ScrollBars = ScrollBars.Both,
                .ReadOnly = True, .Font = New Font("Consolas", 9), .WordWrap = False
            }
            txtGDBCommand = New TextBox() With {
                .Dock = DockStyle.Bottom, .Font = New Font("Consolas", 9), .Height = 24
            }
            AddHandler txtGDBCommand.KeyDown, AddressOf TxtGDBCommand_KeyDown
            tabPageDebugOutput.Controls.Add(txtDebugOutput)
            tabPageDebugOutput.Controls.Add(txtGDBCommand)

            ' Locals/Watch tab - split into locals (top) and watches (bottom)
            tabPageLocals = New TabPage("Locals / Watch")
            Dim splitLocalsWatch As New SplitContainer() With {
                .Dock = DockStyle.Fill, .Orientation = Orientation.Horizontal,
                .SplitterDistance = 120, .SplitterWidth = 4
            }

            ' Locals panel (top)
            Dim lblLocals As New Label() With {.Text = " Locals", .Dock = DockStyle.Top, .Height = 18,
                .Font = New Font("Segoe UI", 8, FontStyle.Bold), .BackColor = Color.FromArgb(220, 220, 220)}
            lvLocals = New ListView() With {
                .Dock = DockStyle.Fill, .View = View.Details, .FullRowSelect = True,
                .GridLines = True, .Font = New Font("Consolas", 9)
            }
            lvLocals.Columns.Add("Name", 160)
            lvLocals.Columns.Add("Value", 250)
            lvLocals.Columns.Add("Type", 120)
            splitLocalsWatch.Panel1.Controls.Add(lvLocals)
            splitLocalsWatch.Panel1.Controls.Add(lblLocals)

            ' Watch panel (bottom) with add/remove toolbar
            Dim lblWatch As New Label() With {.Text = " Watch", .Dock = DockStyle.Top, .Height = 18,
                .Font = New Font("Segoe UI", 8, FontStyle.Bold), .BackColor = Color.FromArgb(220, 220, 220)}
            Dim pnlWatchButtons As New FlowLayoutPanel() With {.Dock = DockStyle.Top, .Height = 28,
                .FlowDirection = FlowDirection.LeftToRight, .WrapContents = False, .Padding = New Padding(2)}
            Dim btnAddWatch As New Button() With {.Text = "+ Add", .Width = 65, .Height = 24, .FlatStyle = FlatStyle.System}
            Dim btnRemWatch As New Button() With {.Text = "- Remove", .Width = 75, .Height = 24, .FlatStyle = FlatStyle.System}
            Dim btnRefreshWatch As New Button() With {.Text = "Refresh", .Width = 65, .Height = 24, .FlatStyle = FlatStyle.System}
            AddHandler btnAddWatch.Click, Sub(s, ev) DoAddWatch()
            AddHandler btnRemWatch.Click, Sub(s, ev) DoRemoveSelectedWatch()
            AddHandler btnRefreshWatch.Click, Sub(s, ev)
                                                  If _debugger.IsRunning AndAlso _debugger.IsPaused Then
                                                      _debugger.RequestLocals()
                                                      _debugger.RefreshWatches()
                                                  End If
                                              End Sub
            pnlWatchButtons.Controls.AddRange(New Control() {btnAddWatch, btnRemWatch, btnRefreshWatch})

            lvWatch = New ListView() With {
                .Dock = DockStyle.Fill, .View = View.Details, .FullRowSelect = True,
                .GridLines = True, .Font = New Font("Consolas", 9), .LabelEdit = True
            }
            lvWatch.Columns.Add("Expression", 160)
            lvWatch.Columns.Add("Value", 250)
            lvWatch.Columns.Add("Type", 120)
            AddHandler lvWatch.KeyDown, Sub(s, ev)
                                            If ev.KeyCode = Keys.Delete Then DoRemoveSelectedWatch()
                                            If ev.KeyCode = Keys.Insert Then DoAddWatch()
                                        End Sub
            splitLocalsWatch.Panel2.Controls.Add(lvWatch)
            splitLocalsWatch.Panel2.Controls.Add(pnlWatchButtons)
            splitLocalsWatch.Panel2.Controls.Add(lblWatch)

            tabPageLocals.Controls.Add(splitLocalsWatch)

            ' Call Stack tab
            tabPageCallStack = New TabPage("Call Stack")
            lvCallStack = New ListView() With {
                .Dock = DockStyle.Fill, .View = View.Details, .FullRowSelect = True,
                .GridLines = True, .Font = New Font("Consolas", 9)
            }
            lvCallStack.Columns.Add("#", 30)
            lvCallStack.Columns.Add("Function", 200)
            lvCallStack.Columns.Add("File", 200)
            lvCallStack.Columns.Add("Line", 60)
            lvCallStack.Columns.Add("Address", 120)
            AddHandler lvCallStack.DoubleClick, AddressOf LvCallStack_DoubleClick
            tabPageCallStack.Controls.Add(lvCallStack)

            ' AI Chat tab
            tabPageAIChat = New TabPage("AI Chat")
            txtAIChat = New RichTextBox() With {
                .Dock = DockStyle.Fill, .ReadOnly = True, .Font = New Font("Consolas", 10),
                .WordWrap = True, .Text = "Welcome to AI Chat! Ask Claude about FreeBASIC programming." & vbCrLf &
                    "Use 'Send' to ask a question, or 'Send Code' to include your code." & vbCrLf &
                    "Press Ctrl+Enter in the input box to send quickly." & vbCrLf & vbCrLf
            }

            pnlAIActions = New FlowLayoutPanel() With {.Dock = DockStyle.Bottom, .Height = 32, .FlowDirection = FlowDirection.LeftToRight, .WrapContents = False}
            btnAIInsertCode = New Button() With {.Text = "Insert Code", .Width = 95, .Height = 28}
            btnAICopyReply = New Button() With {.Text = "Copy Reply", .Width = 90, .Height = 28}
            btnAIReplaceAll = New Button() With {.Text = "Replace All", .Width = 90, .Height = 28}
            btnAINewFile = New Button() With {.Text = "New File", .Width = 80, .Height = 28}
            btnAIClearChat = New Button() With {.Text = "Clear Chat", .Width = 85, .Height = 28}
            pnlAIActions.Controls.AddRange(New Control() {btnAIInsertCode, btnAICopyReply, btnAIReplaceAll, btnAINewFile, btnAIClearChat})

            pnlAIInput = New Panel() With {.Dock = DockStyle.Bottom, .Height = 32}
            txtAIInput = New TextBox() With {.Dock = DockStyle.Fill, .Font = New Font("Consolas", 10)}
            btnAISend = New Button() With {.Text = "Send", .Width = 65, .Dock = DockStyle.Right, .Height = 28}
            btnAISendCode = New Button() With {.Text = "Send Code", .Width = 85, .Dock = DockStyle.Right, .Height = 28}
            pnlAIInput.Controls.Add(txtAIInput)
            pnlAIInput.Controls.Add(btnAISendCode)
            pnlAIInput.Controls.Add(btnAISend)

            tabPageAIChat.Controls.Add(txtAIChat)
            tabPageAIChat.Controls.Add(pnlAIActions)
            tabPageAIChat.Controls.Add(pnlAIInput)

            tabOutput.TabPages.Add(tabPageOutput)
            tabOutput.TabPages.Add(tabPageDebugOutput)
            tabOutput.TabPages.Add(tabPageLocals)
            tabOutput.TabPages.Add(tabPageCallStack)
            tabOutput.TabPages.Add(tabPageAIChat)
        End Sub

#Region "Menu Building"
        Private Shared Function AddMenu(parent As ToolStripMenuItem, text As String, handler As EventHandler, Optional shortcut As Keys = Keys.None) As ToolStripMenuItem
            Dim item As New ToolStripMenuItem(text)
            If handler IsNot Nothing Then AddHandler item.Click, handler
            If shortcut <> Keys.None Then item.ShortcutKeys = shortcut
            parent.DropDownItems.Add(item)
            Return item
        End Function

        Private Sub BuildMenus()
            ' ---- FILE ----
            Dim mnuFile = New ToolStripMenuItem("&File")
            AddMenu(mnuFile, "&New", Sub(s, e) DoNewFile(), Keys.Control Or Keys.N)
            AddMenu(mnuFile, "&Open...", Sub(s, e) DoOpenFile(), Keys.Control Or Keys.O)
            mnuFile.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuFile, "&Save", Sub(s, e) DoSaveFile(), Keys.Control Or Keys.S)
            AddMenu(mnuFile, "Save &As...", Sub(s, e) DoSaveFileAs())
            AddMenu(mnuFile, "Save A&ll", Sub(s, e) DoSaveAll())
            mnuFile.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuFile, "&Close File", Sub(s, e) DoCloseFile(), Keys.Control Or Keys.W)
            AddMenu(mnuFile, "Close All", Sub(s, e) DoCloseAll())
            mnuFile.DropDownItems.Add(New ToolStripSeparator())
            _mnuRecentParent = New ToolStripMenuItem("Recent &Files")
            mnuFile.DropDownItems.Add(_mnuRecentParent)
            mnuFile.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuFile, "E&xit", Sub(s, e) Me.Close())

            ' ---- EDIT ----
            Dim mnuEdit = New ToolStripMenuItem("&Edit")
            AddMenu(mnuEdit, "&Undo", Sub(s, e) scintilla.Undo(), Keys.Control Or Keys.Z)
            AddMenu(mnuEdit, "&Redo", Sub(s, e) scintilla.Redo(), Keys.Control Or Keys.Y)
            mnuEdit.DropDownItems.Add(New ToolStripSeparator())
            ' No ShortcutKeys for Cut/Copy/Paste/SelectAll - let the focused control handle them natively
            ' This allows Ctrl+C/V/X/A to work in AI Chat, output panel, etc.
            Dim mnuCut = AddMenu(mnuEdit, "Cu&t", Sub(s, e) scintilla.Cut())
            mnuCut.ShortcutKeyDisplayString = "Ctrl+X"
            Dim mnuCopy = AddMenu(mnuEdit, "&Copy", Sub(s, e) scintilla.Copy())
            mnuCopy.ShortcutKeyDisplayString = "Ctrl+C"
            Dim mnuPaste = AddMenu(mnuEdit, "&Paste", Sub(s, e) scintilla.Paste())
            mnuPaste.ShortcutKeyDisplayString = "Ctrl+V"
            AddMenu(mnuEdit, "&Delete", Sub(s, e) scintilla.ReplaceSelection(""))
            mnuEdit.DropDownItems.Add(New ToolStripSeparator())
            Dim mnuSelAll = AddMenu(mnuEdit, "Select &All", Sub(s, e) scintilla.SelectAll())
            mnuSelAll.ShortcutKeyDisplayString = "Ctrl+A"
            mnuEdit.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuEdit, "Co&mment Block", Sub(s, e) CommentBlock())
            AddMenu(mnuEdit, "U&ncomment Block", Sub(s, e) UncommentBlock())
            mnuEdit.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuEdit, "Toggle Book&mark", Sub(s, e) ToggleBookmark())
            AddMenu(mnuEdit, "Next Bookmark", Sub(s, e) NextBookmark())
            AddMenu(mnuEdit, "Previous Bookmark", Sub(s, e) PrevBookmark())
            AddMenu(mnuEdit, "Clear All Bookmarks", Sub(s, e) scintilla.MarkerDeleteAll(MARKER_BOOKMARK))
            mnuEdit.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuEdit, "&Auto-Complete", Sub(s, e) ShowAutoComplete(True), Keys.Control Or Keys.Space)
            mnuEditAutoCompleteEnable = New ToolStripMenuItem("Enable Auto-Complete") With {.Checked = Settings.AutoComplete, .CheckOnClick = True}
            AddHandler mnuEditAutoCompleteEnable.CheckedChanged, Sub() Settings.AutoComplete = mnuEditAutoCompleteEnable.Checked
            mnuEdit.DropDownItems.Add(mnuEditAutoCompleteEnable)

            ' ---- SEARCH ----
            Dim mnuSearch = New ToolStripMenuItem("&Search")
            AddMenu(mnuSearch, "&Find...", Sub(s, e) ShowFindReplace(False), Keys.Control Or Keys.F)
            AddMenu(mnuSearch, "Find &Next", Sub(s, e) FindNext(), Keys.F3)
            AddMenu(mnuSearch, "&Replace...", Sub(s, e) ShowFindReplace(True), Keys.Control Or Keys.H)
            mnuSearch.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuSearch, "&Go To Line...", Sub(s, e) GoToLine(), Keys.Control Or Keys.G)

            ' ---- VIEW ----
            Dim mnuView = New ToolStripMenuItem("&View")
            mnuViewProjectExplorer = New ToolStripMenuItem("&Project Explorer") With {.Checked = True, .CheckOnClick = True}
            AddHandler mnuViewProjectExplorer.CheckedChanged, Sub() splitMain.Panel1Collapsed = Not mnuViewProjectExplorer.Checked
            mnuViewOutputPanel = New ToolStripMenuItem("&Output Panel") With {.Checked = True, .CheckOnClick = True}
            AddHandler mnuViewOutputPanel.CheckedChanged, Sub() splitEditor.Panel2Collapsed = Not mnuViewOutputPanel.Checked
            mnuView.DropDownItems.Add(mnuViewProjectExplorer)
            mnuView.DropDownItems.Add(mnuViewOutputPanel)
            AddMenu(mnuView, "&Refresh Code Outline", Sub(s, e) UpdateOutline(), Keys.F4)
            mnuView.DropDownItems.Add(New ToolStripSeparator())
            mnuViewLineNumbers = New ToolStripMenuItem("&Line Numbers") With {.Checked = Settings.ShowLineNumbers, .CheckOnClick = True}
            AddHandler mnuViewLineNumbers.CheckedChanged, Sub()
                                                             Settings.ShowLineNumbers = mnuViewLineNumbers.Checked
                                                             SetupMargins()
                                                         End Sub
            mnuViewWordWrap = New ToolStripMenuItem("&Word Wrap") With {.Checked = Settings.WordWrap, .CheckOnClick = True}
            AddHandler mnuViewWordWrap.CheckedChanged, Sub()
                                                          Settings.WordWrap = mnuViewWordWrap.Checked
                                                          scintilla.WrapMode = If(Settings.WordWrap, WrapMode.Word, WrapMode.None)
                                                      End Sub
            mnuViewWhitespace = New ToolStripMenuItem("White&space") With {.Checked = Settings.ShowWhitespace, .CheckOnClick = True}
            AddHandler mnuViewWhitespace.CheckedChanged, Sub()
                                                            Settings.ShowWhitespace = mnuViewWhitespace.Checked
                                                            scintilla.ViewWhitespace = If(Settings.ShowWhitespace, WhitespaceMode.VisibleAlways, WhitespaceMode.Invisible)
                                                        End Sub
            mnuViewIndentGuides = New ToolStripMenuItem("&Indentation Guides") With {.Checked = Settings.ShowIndentGuides, .CheckOnClick = True}
            AddHandler mnuViewIndentGuides.CheckedChanged, Sub()
                                                              Settings.ShowIndentGuides = mnuViewIndentGuides.Checked
                                                              scintilla.IndentationGuides = If(Settings.ShowIndentGuides, IndentView.LookBoth, IndentView.None)
                                                          End Sub
            mnuView.DropDownItems.AddRange(New ToolStripItem() {mnuViewLineNumbers, mnuViewWordWrap, mnuViewWhitespace, mnuViewIndentGuides})
            mnuView.DropDownItems.Add(New ToolStripSeparator())

            ' ---- Folding submenu ----
            Dim mnuViewFolding = New ToolStripMenuItem("Code &Folding") With {.Checked = Settings.ShowFolding, .CheckOnClick = True}
            AddHandler mnuViewFolding.CheckedChanged, Sub()
                                                          Settings.ShowFolding = mnuViewFolding.Checked
                                                          If mnuViewFolding.Checked Then
                                                              FoldingManager.SetupFoldMargin(scintilla, Settings.DarkTheme)
                                                              FoldingManager.UpdateFoldLevels(scintilla)
                                                          Else
                                                              FoldingManager.HideFoldMargin(scintilla)
                                                          End If
                                                      End Sub
            mnuView.DropDownItems.Add(mnuViewFolding)
            AddMenu(mnuView, "&Toggle Fold", Sub(s, e) FoldingManager.ToggleFold(scintilla), Keys.Control Or Keys.Shift Or Keys.OemOpenBrackets)
            AddMenu(mnuView, "Fold &All", Sub(s, e) FoldingManager.FoldAll(scintilla))
            AddMenu(mnuView, "&Unfold All", Sub(s, e) FoldingManager.UnfoldAll(scintilla))
            AddMenu(mnuView, "Fold to &Level 1", Sub(s, e) FoldingManager.FoldToLevel(scintilla, 1))

            mnuView.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuView, "Zoom &In", Sub(s, e) scintilla.ZoomIn())
            AddMenu(mnuView, "Zoom &Out", Sub(s, e) scintilla.ZoomOut())
            AddMenu(mnuView, "&Reset Zoom", Sub(s, e) scintilla.Zoom = 0)
            mnuView.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuView, "Editor &Font...", Sub(s, e) ChangeEditorFont())
            mnuView.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuView, "Toggle &Dark/Light Theme", Sub(s, e) ToggleTheme())
            mnuView.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuView, "&AI Chat Panel", Sub(s, e)
                                                 mnuViewOutputPanel.Checked = True
                                                 tabOutput.SelectedTab = tabPageAIChat
                                             End Sub)

            ' ---- BUILD ----
            Dim mnuBuild = New ToolStripMenuItem("&Build")
            AddMenu(mnuBuild, "&Compile", Sub(s, e) DoCompile(), Keys.Control Or Keys.F5)
            AddMenu(mnuBuild, "Compile && &Run", Sub(s, e) DoCompileAndRun(), Keys.F6)
            AddMenu(mnuBuild, "Run (&No Compile)", Sub(s, e) DoRunOnly(), Keys.Control Or Keys.F6)
            mnuBuild.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuBuild, "&Quick Run", Sub(s, e) DoQuickRun())
            AddMenu(mnuBuild, "&Syntax Check Only", Sub(s, e) DoSyntaxCheck())
            mnuBuild.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuBuild, "Build &Options...", Sub(s, e) ShowBuildOptions())

            ' ---- DEBUG ----
            Dim mnuDebug = New ToolStripMenuItem("&Debug")
            AddMenu(mnuDebug, "&Start / Continue", Sub(s, e) DoDebugStart(), Keys.F5)
            AddMenu(mnuDebug, "S&top Debugging", Sub(s, e) DoDebugStop(), Keys.Shift Or Keys.F5)
            AddMenu(mnuDebug, "&Pause", Sub(s, e) DoDebugPause())
            mnuDebug.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuDebug, "Step &Over", Sub(s, e) _debugger.StepOver(), Keys.F10)
            AddMenu(mnuDebug, "Step &Into", Sub(s, e) _debugger.StepInto(), Keys.F11)
            AddMenu(mnuDebug, "Step O&ut", Sub(s, e) _debugger.StepOut(), Keys.Shift Or Keys.F11)
            AddMenu(mnuDebug, "Run to &Cursor", Sub(s, e) DoRunToCursor(), Keys.Control Or Keys.F10)
            mnuDebug.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuDebug, "Toggle &Breakpoint", Sub(s, e) DoToggleBreakpoint(), Keys.F9)
            AddMenu(mnuDebug, "Clear All Breakpoints", Sub(s, e) DoClearAllBreakpoints())
            mnuDebug.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuDebug, "&Add Watch Expression...", Sub(s, e) DoAddWatch())
            AddMenu(mnuDebug, "Set &GDB Path...", Sub(s, e) SetGDBPath())

            ' ---- TOOLS ----
            Dim mnuTools = New ToolStripMenuItem("&Tools")
            AddMenu(mnuTools, "&FreeBASIC Compiler Path...", Sub(s, e) SetFBCPath())
            mnuTools.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuTools, "Open &Settings Folder", Sub(s, e) SafeProcessStart(SettingsPath))
            mnuTools.DropDownItems.Add(New ToolStripSeparator())
            Dim mnuEncoding = New ToolStripMenuItem("File &Encoding")
            mnuEncANSI = New ToolStripMenuItem("&ANSI (Windows-1252)", Nothing, Sub(s, e) ChangeFileEncoding(FileEncoding.ANSI))
            mnuEncUTF8 = New ToolStripMenuItem("UTF-&8", Nothing, Sub(s, e) ChangeFileEncoding(FileEncoding.UTF8))
            mnuEncUTF8BOM = New ToolStripMenuItem("UTF-8 with &BOM", Nothing, Sub(s, e) ChangeFileEncoding(FileEncoding.UTF8_BOM))
            mnuEncDefaultANSI = New ToolStripMenuItem("Default: ANSI", Nothing, Sub(s, e)
                                                                                   Settings.DefaultEncoding = FileEncoding.ANSI
                                                                                   SaveSettings()
                                                                                   UpdateEncodingUI()
                                                                               End Sub)
            mnuEncDefaultUTF8 = New ToolStripMenuItem("Default: UTF-8", Nothing, Sub(s, e)
                                                                                   Settings.DefaultEncoding = FileEncoding.UTF8
                                                                                   SaveSettings()
                                                                                   UpdateEncodingUI()
                                                                               End Sub)
            mnuEncoding.DropDownItems.AddRange(New ToolStripItem() {mnuEncANSI, mnuEncUTF8, mnuEncUTF8BOM, New ToolStripSeparator(), mnuEncDefaultANSI, mnuEncDefaultUTF8})
            mnuTools.DropDownItems.Add(mnuEncoding)

            ' ---- HELP ----
            Dim mnuHelp = New ToolStripMenuItem("&Help")
            AddMenu(mnuHelp, "&Local FreeBASIC Help", Sub(s, e) OpenFBHelp(), Keys.F1)
            AddMenu(mnuHelp, "FreeBASIC &Online Documentation", Sub(s, e) SafeProcessStart("https://www.freebasic.net/wiki"))
            mnuHelp.DropDownItems.Add(New ToolStripSeparator())
            AddMenu(mnuHelp, "&About " & APP_NAME & "...", Sub(s, e) ShowAbout())

            menuMain.Items.AddRange(New ToolStripItem() {mnuFile, mnuEdit, mnuSearch, mnuView, mnuBuild, mnuDebug, mnuTools, mnuHelp})
        End Sub

        Private Sub BuildToolbar()
            toolMain.Items.Add(New ToolStripButton("New", Nothing, Sub() DoNewFile()) With {.ToolTipText = "New File (Ctrl+N)"})
            toolMain.Items.Add(New ToolStripButton("Open", Nothing, Sub() DoOpenFile()) With {.ToolTipText = "Open File (Ctrl+O)"})
            toolMain.Items.Add(New ToolStripButton("Save", Nothing, Sub() DoSaveFile()) With {.ToolTipText = "Save File (Ctrl+S)"})
            toolMain.Items.Add(New ToolStripButton("All", Nothing, Sub() DoSaveAll()) With {.ToolTipText = "Save All"})
            toolMain.Items.Add(New ToolStripSeparator())
            toolMain.Items.Add(New ToolStripButton("Undo", Nothing, Sub() scintilla.Undo()) With {.ToolTipText = "Undo"})
            toolMain.Items.Add(New ToolStripButton("Redo", Nothing, Sub() scintilla.Redo()) With {.ToolTipText = "Redo"})
            toolMain.Items.Add(New ToolStripSeparator())
            toolMain.Items.Add(New ToolStripButton("Cut", Nothing, Sub() scintilla.Cut()) With {.ToolTipText = "Cut"})
            toolMain.Items.Add(New ToolStripButton("Copy", Nothing, Sub() scintilla.Copy()) With {.ToolTipText = "Copy"})
            toolMain.Items.Add(New ToolStripButton("Paste", Nothing, Sub() scintilla.Paste()) With {.ToolTipText = "Paste"})
            toolMain.Items.Add(New ToolStripSeparator())
            toolMain.Items.Add(New ToolStripButton("Find", Nothing, Sub() ShowFindReplace(False)) With {.ToolTipText = "Find (Ctrl+F)"})
            toolMain.Items.Add(New ToolStripSeparator())
            toolMain.Items.Add(New ToolStripButton("Compile", Nothing, Sub() DoCompile()) With {.ToolTipText = "Compile (Ctrl+F5)"})
            toolMain.Items.Add(New ToolStripButton("Run", Nothing, Sub() DoCompileAndRun()) With {.ToolTipText = "Compile & Run (F6)"})
            toolMain.Items.Add(New ToolStripButton("Opt", Nothing, Sub() ShowBuildOptions()) With {.ToolTipText = "Build Options"})
            toolMain.Items.Add(New ToolStripSeparator())
            toolMain.Items.Add(New ToolStripButton("AI", Nothing, Sub()
                                                                     mnuViewOutputPanel.Checked = True
                                                                     tabOutput.SelectedTab = tabPageAIChat
                                                                 End Sub) With {.ToolTipText = "AI Chat (Claude)"})
        End Sub

        Private Sub BuildDebugToolbar()
            btnDebugStart = New ToolStripButton("▶ Debug", Nothing, Sub() DoDebugStart()) With {.ToolTipText = "Start/Continue Debugging (F5)"}
            btnDebugStop = New ToolStripButton("■ Stop", Nothing, Sub() DoDebugStop()) With {.ToolTipText = "Stop Debugging (Shift+F5)"}
            btnDebugPause = New ToolStripButton("❚❚ Pause", Nothing, Sub() DoDebugPause()) With {.ToolTipText = "Pause Debugging"}
            toolDebug.Items.Add(btnDebugStart)
            toolDebug.Items.Add(btnDebugStop)
            toolDebug.Items.Add(btnDebugPause)
            toolDebug.Items.Add(New ToolStripSeparator())
            btnDebugStepOver = New ToolStripButton("Step Over", Nothing, Sub() _debugger.StepOver()) With {.ToolTipText = "Step Over (F10)"}
            btnDebugStepInto = New ToolStripButton("Step Into", Nothing, Sub() _debugger.StepInto()) With {.ToolTipText = "Step Into (F11)"}
            btnDebugStepOut = New ToolStripButton("Step Out", Nothing, Sub() _debugger.StepOut()) With {.ToolTipText = "Step Out (Shift+F11)"}
            toolDebug.Items.Add(btnDebugStepOver)
            toolDebug.Items.Add(btnDebugStepInto)
            toolDebug.Items.Add(btnDebugStepOut)
            toolDebug.Items.Add(New ToolStripSeparator())
            toolDebug.Items.Add(New ToolStripButton("BP", Nothing, Sub() DoToggleBreakpoint()) With {.ToolTipText = "Toggle Breakpoint (F9)"})
        End Sub
#End Region

#Region "Scintilla Setup"
        Private Sub SetupScintilla()
            scintilla.Lexer = Lexer.FreeBasic
            scintilla.SetKeywords(0, SyntaxConfig.FB_KEYWORDS)
            scintilla.SetKeywords(1, SyntaxConfig.FB_TYPES)
            scintilla.SetKeywords(2, SyntaxConfig.FB_PREPROCESSOR)
            scintilla.SetKeywords(3, SyntaxConfig.FB_FUNCTIONS)

            scintilla.TabWidth = Settings.TabWidth
            scintilla.UseTabs = Settings.UseTabs
            scintilla.IndentWidth = Settings.TabWidth
            scintilla.WrapMode = If(Settings.WordWrap, WrapMode.Word, WrapMode.None)
            scintilla.ViewWhitespace = If(Settings.ShowWhitespace, WhitespaceMode.VisibleAlways, WhitespaceMode.Invisible)
            scintilla.IndentationGuides = If(Settings.ShowIndentGuides, IndentView.LookBoth, IndentView.None)

            scintilla.AutoCIgnoreCase = True
            scintilla.AutoCMaxHeight = 12

            ' Bookmark margin (margin 1) - holds bookmarks + breakpoints
            Dim bmMargin = scintilla.Margins(1)
            bmMargin.Width = 20
            bmMargin.Sensitive = True
            bmMargin.Type = MarginType.Symbol
            bmMargin.Mask = (1 << MARKER_BOOKMARK) Or (1 << MARKER_BREAKPOINT) Or (1 << MARKER_DEBUGLINE)

            ' Marker 1: Bookmark (blue circle)
            scintilla.Markers(MARKER_BOOKMARK).Symbol = MarkerSymbol.Circle
            scintilla.Markers(MARKER_BOOKMARK).SetBackColor(Color.DodgerBlue)
            scintilla.Markers(MARKER_BOOKMARK).SetForeColor(Color.White)

            ' Marker 2: Breakpoint (red circle)
            scintilla.Markers(MARKER_BREAKPOINT).Symbol = MarkerSymbol.Circle
            scintilla.Markers(MARKER_BREAKPOINT).SetBackColor(Color.Red)
            scintilla.Markers(MARKER_BREAKPOINT).SetForeColor(Color.White)

            ' Marker 3: Current debug line (yellow arrow)
            scintilla.Markers(MARKER_DEBUGLINE).Symbol = MarkerSymbol.ShortArrow
            scintilla.Markers(MARKER_DEBUGLINE).SetBackColor(Color.Yellow)
            scintilla.Markers(MARKER_DEBUGLINE).SetForeColor(Color.Black)

            SetupMargins()
            ApplyEditorTheme()
            scintilla.Technology = Technology.DirectWrite

            ' ---- Folding ----
            If Settings.ShowFolding Then
                FoldingManager.SetupFoldMargin(scintilla, Settings.DarkTheme)
            Else
                FoldingManager.HideFoldMargin(scintilla)
            End If
        End Sub

        Private Sub SetupMargins()
            If Settings.ShowLineNumbers Then
                Dim lineCount = Math.Max(scintilla.Lines.Count, 10000)
                scintilla.Margins(0).Width = scintilla.TextWidth(Style.LineNumber, New String("9"c, lineCount.ToString().Length + 1))
            Else
                scintilla.Margins(0).Width = 0
            End If
        End Sub

        Private Sub ApplyEditorTheme()
            Dim dark = Settings.DarkTheme
            scintilla.StyleResetDefault()
            scintilla.Styles(Style.Default).Font = Settings.EditorFont
            scintilla.Styles(Style.Default).Size = Settings.EditorFontSize
            scintilla.Styles(Style.Default).BackColor = If(dark, Color.FromArgb(30, 30, 30), Color.White)
            scintilla.Styles(Style.Default).ForeColor = If(dark, Color.FromArgb(212, 212, 212), Color.FromArgb(30, 30, 30))
            scintilla.StyleClearAll()

            scintilla.Styles(Style.FreeBasic.Default).ForeColor = If(dark, Color.FromArgb(212, 212, 212), Color.Black)
            scintilla.Styles(Style.FreeBasic.Comment).ForeColor = If(dark, Color.FromArgb(106, 153, 85), Color.Green)
            scintilla.Styles(Style.FreeBasic.CommentBlock).ForeColor = If(dark, Color.FromArgb(106, 153, 85), Color.Green)
            scintilla.Styles(Style.FreeBasic.Number).ForeColor = If(dark, Color.FromArgb(181, 206, 168), Color.Teal)
            scintilla.Styles(Style.FreeBasic.Keyword).ForeColor = If(dark, Color.FromArgb(86, 156, 214), Color.Blue)
            scintilla.Styles(Style.FreeBasic.Keyword).Bold = True
            scintilla.Styles(Style.FreeBasic.String).ForeColor = If(dark, Color.FromArgb(206, 145, 120), Color.FromArgb(163, 21, 21))
            scintilla.Styles(Style.FreeBasic.Preprocessor).ForeColor = If(dark, Color.FromArgb(155, 155, 155), Color.Gray)
            scintilla.Styles(Style.FreeBasic.Operator).ForeColor = If(dark, Color.FromArgb(212, 212, 212), Color.Black)
            scintilla.Styles(Style.FreeBasic.Keyword2).ForeColor = If(dark, Color.FromArgb(78, 201, 176), Color.FromArgb(128, 0, 128))
            scintilla.Styles(Style.FreeBasic.Keyword3).ForeColor = If(dark, Color.FromArgb(155, 155, 155), Color.Gray)
            scintilla.Styles(Style.FreeBasic.Keyword4).ForeColor = If(dark, Color.FromArgb(220, 220, 170), Color.FromArgb(0, 128, 128))

            scintilla.Styles(Style.LineNumber).ForeColor = If(dark, Color.FromArgb(133, 133, 133), Color.Gray)
            scintilla.Styles(Style.LineNumber).BackColor = If(dark, Color.FromArgb(30, 30, 30), Color.White)

            scintilla.CaretForeColor = If(dark, Color.FromArgb(174, 175, 173), Color.Black)
            scintilla.SetSelectionBackColor(True, If(dark, Color.FromArgb(38, 79, 120), Color.FromArgb(173, 214, 255)))

            If Settings.HighlightCurrentLine Then
                scintilla.CaretLineVisible = True
                scintilla.CaretLineBackColor = If(dark, Color.FromArgb(42, 45, 46), Color.FromArgb(255, 255, 224))
            End If

            scintilla.Styles(Style.IndentGuide).ForeColor = If(dark, Color.FromArgb(64, 64, 64), Color.FromArgb(192, 192, 192))

            ' Brace matching styles
            scintilla.Styles(Style.BraceLight).ForeColor = If(dark, Color.FromArgb(220, 220, 100), Color.Blue)
            scintilla.Styles(Style.BraceLight).BackColor = If(dark, Color.FromArgb(60, 60, 60), Color.FromArgb(220, 220, 255))
            scintilla.Styles(Style.BraceLight).Bold = True
            scintilla.Styles(Style.BraceBad).ForeColor = Color.Red
            scintilla.Styles(Style.BraceBad).BackColor = If(dark, Color.FromArgb(60, 30, 30), Color.FromArgb(255, 220, 220))
            scintilla.Styles(Style.BraceBad).Bold = True
        End Sub
#End Region

#Region "Theme Management"
        Private Sub ApplyCurrentTheme()
            ThemeManager.ApplyTheme(Me, Settings.DarkTheme)
            If _initialized Then ApplyEditorTheme()
        End Sub

        Private Sub ToggleTheme()
            Settings.DarkTheme = Not Settings.DarkTheme
            ApplyCurrentTheme()
            ApplyEditorTheme()
            scintilla.Colorize(0, -1)
            SaveSettings()
        End Sub
#End Region

#Region "File Management"
        Public Sub DoNewFile()
            For i = 0 To _files.Count - 1
                If _files(i).IsNew AndAlso Not _files(i).IsModified AndAlso String.IsNullOrEmpty(_files(i).Content) Then
                    If i = _activeFile AndAlso scintilla.TextLength = 0 Then
                        SwitchToFile(i) : Return
                    ElseIf i <> _activeFile Then
                        SwitchToFile(i) : Return
                    End If
                End If
            Next

            Dim fi As New OpenFileInfo() With {.FileName = NewUntitledName(), .IsNew = True, .FileEnc = Settings.DefaultEncoding}
            _files.Add(fi)
            UpdateFileList()
            UpdateTreeView()
            SwitchToFile(_files.Count - 1)
        End Sub

        Public Sub DoOpenFile()
            Using dlg As New OpenFileDialog()
                dlg.Filter = "FreeBASIC Files (*.bas;*.bi;*.rc)|*.bas;*.bi;*.rc|BASIC Files (*.bas)|*.bas|Include Files (*.bi)|*.bi|All Files (*.*)|*.*"
                dlg.Title = "Open File"
                dlg.Multiselect = True    ' ENHANCEMENT: multi-file open
                If dlg.ShowDialog() = DialogResult.OK Then
                    For Each f In dlg.FileNames
                        OpenFileByPath(f)
                    Next
                End If
            End Using
        End Sub

        Public Sub OpenFileByPath(filePath As String)
            For i = 0 To _files.Count - 1
                If _files(i).FilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase) Then
                    SwitchToFile(i) : Return
                End If
            Next

            Dim enc As FileEncoding
            Dim content = ReadFileWithEncoding(filePath, enc)

            If _files.Count = 1 AndAlso _files(0).IsNew AndAlso Not _files(0).IsModified Then
                If _activeFile = 0 AndAlso scintilla.TextLength = 0 Then
                    _activeFile = -1
                    _files.RemoveAt(0)
                End If
            End If

            Dim fi As New OpenFileInfo() With {.FilePath = filePath, .FileName = Path.GetFileName(filePath), .IsNew = False, .Content = content, .FileEnc = enc}
            _files.Add(fi)
            UpdateFileList()
            UpdateTreeView()
            SwitchToFile(_files.Count - 1)
            AddRecentFile(filePath)
            UpdateRecentFilesMenu()
        End Sub

        Public Function DoSaveFile() As Boolean
            If _activeFile < 0 OrElse _activeFile >= _files.Count Then Return False
            If _files(_activeFile).IsNew OrElse String.IsNullOrEmpty(_files(_activeFile).FilePath) Then Return DoSaveFileAs()

            Dim content = scintilla.Text
            Try
                WriteFileWithEncoding(_files(_activeFile).FilePath, content, _files(_activeFile).FileEnc)
                _files(_activeFile).IsModified = False
                _files(_activeFile).IsNew = False
                scintilla.SetSavePoint()
                UpdateTitle() : UpdateFileList() : UpdateTreeView()
                lblStatus.Text = "Saved: " & _files(_activeFile).FileName
                Return True
            Catch ex As Exception
                MessageBox.Show("Error saving: " & ex.Message, APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

        Public Function DoSaveFileAs() As Boolean
            Using dlg As New SaveFileDialog()
                dlg.Filter = "FreeBASIC Files (*.bas)|*.bas|Include Files (*.bi)|*.bi|All Files (*.*)|*.*"
                dlg.Title = "Save File As"
                If _activeFile >= 0 Then dlg.FileName = _files(_activeFile).FileName
                If dlg.ShowDialog() = DialogResult.OK Then
                    _files(_activeFile).FilePath = dlg.FileName
                    _files(_activeFile).FileName = Path.GetFileName(dlg.FileName)
                    _files(_activeFile).IsNew = False
                    Dim result = DoSaveFile()
                    If result Then
                        AddRecentFile(dlg.FileName)
                        UpdateRecentFilesMenu()
                    End If
                    Return result
                End If
            End Using
            Return False
        End Function

        Public Sub DoSaveAll()
            Dim orig = _activeFile
            For i = 0 To _files.Count - 1
                If _files(i).IsModified Then
                    SwitchToFile(i)
                    DoSaveFile()
                End If
            Next
            If orig >= 0 AndAlso orig < _files.Count Then SwitchToFile(orig)
        End Sub

        Public Sub DoCloseFile()
            If _activeFile < 0 OrElse _files.Count = 0 Then Return
            If _files(_activeFile).IsModified Then
                Dim r = MessageBox.Show("Save changes to " & _files(_activeFile).FileName & "?",
                    APP_NAME, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                If r = DialogResult.Yes Then
                    If Not DoSaveFile() Then Return
                ElseIf r = DialogResult.Cancel Then
                    Return
                End If
            End If

            Dim closed = _activeFile
            _activeFile = -1
            _files.RemoveAt(closed)
            UpdateFileList() : UpdateTreeView()

            If _files.Count = 0 Then DoNewFile() Else SwitchToFile(Math.Min(closed, _files.Count - 1))
        End Sub

        Private Sub DoCloseAll()
            ' Prompt for unsaved files first (don't modify collection while iterating)
            For i = 0 To _files.Count - 1
                If _files(i).IsModified Then
                    Dim r = MessageBox.Show("Save changes to " & _files(i).FileName & "?",
                        APP_NAME, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                    If r = DialogResult.Yes Then
                        SwitchToFile(i)
                        If Not DoSaveFile() Then Return
                    ElseIf r = DialogResult.Cancel Then
                        Return
                    End If
                End If
            Next
            _activeFile = -1 : _files.Clear() : UpdateFileList() : UpdateTreeView() : DoNewFile()
        End Sub

        Private Sub SwitchToFile(index As Integer)
            If index < 0 OrElse index >= _files.Count Then Return
            ' BUG FIX: removed early return when index = _activeFile (allows refresh)

            ' Save current state
            If _activeFile >= 0 AndAlso _activeFile < _files.Count Then
                _files(_activeFile).Content = scintilla.Text
                _files(_activeFile).FirstVisibleLine = scintilla.FirstVisibleLine
                _files(_activeFile).CursorPos = scintilla.CurrentPosition
            End If

            Dim wasModified = _files(index).IsModified
            _switchingFile = True
            _activeFile = index

            scintilla.Text = _files(index).Content
            scintilla.GotoPosition(_files(index).CursorPos)
            scintilla.FirstVisibleLine = _files(index).FirstVisibleLine
            scintilla.EmptyUndoBuffer()
            scintilla.SetSavePoint()

            _files(index).IsModified = wasModified
            _switchingFile = False

            UpdateFileList() : UpdateTreeView() : UpdateEncodingUI() : UpdateTitle() : UpdateOutline() : SetupMargins()

            ' Restore breakpoint markers for this file
            RestoreBreakpointMarkers()
            scintilla.Focus()
        End Sub

        Private Sub RestoreBreakpointMarkers()
            ' Clear all breakpoint markers first
            scintilla.MarkerDeleteAll(MARKER_BREAKPOINT)
            If _activeFile < 0 OrElse _activeFile >= _files.Count Then Return
            Dim filePath = _files(_activeFile).FilePath
            If String.IsNullOrEmpty(filePath) Then Return
            Dim bps = _debugger.GetBreakpointsForFile(filePath)
            For Each bp In bps
                If bp.LineNumber > 0 AndAlso bp.LineNumber <= scintilla.Lines.Count Then
                    scintilla.Lines(bp.LineNumber - 1).MarkerAdd(MARKER_BREAKPOINT)
                End If
            Next
        End Sub

        Private Sub UpdateFileList()
            Dim prev = _switchingFile : _switchingFile = True
            cboOpenFiles.Items.Clear()
            For i = 0 To _files.Count - 1
                Dim label = _files(i).FileName
                If _files(i).IsModified Then label &= " *"
                cboOpenFiles.Items.Add(label)
            Next
            If _activeFile >= 0 AndAlso _activeFile < cboOpenFiles.Items.Count Then cboOpenFiles.SelectedIndex = _activeFile
            _switchingFile = prev
        End Sub

        Private Sub UpdateTreeView()
            tvProject.Nodes.Clear()
            Dim root = tvProject.Nodes.Add("root", "Open Files")
            For i = 0 To _files.Count - 1
                Dim label = If(String.IsNullOrEmpty(_files(i).FileName), "Untitled", _files(i).FileName)
                If _files(i).IsModified Then label &= " *"
                root.Nodes.Add("file" & i, label)
            Next
            root.Expand()
        End Sub

        Private Sub UpdateOutline()
            tvOutline.Nodes.Clear()
            If _activeFile < 0 OrElse _activeFile >= _files.Count Then Return
            Dim code = scintilla.Text
            If String.IsNullOrEmpty(code) Then Return

            Dim items = CodeOutline.ParseOutline(code)
            If items.Count = 0 Then Return

            Dim categories As New Dictionary(Of String, TreeNode)()
            For Each item In items
                If Not categories.ContainsKey(item.Category) Then
                    categories(item.Category) = tvOutline.Nodes.Add("cat_" & item.Category, item.Category)
                End If
                Dim label = $"[{item.Icon}] {item.Name}"
                If Not String.IsNullOrEmpty(item.DataType) Then label &= " As " & item.DataType
                categories(item.Category).Nodes.Add("item_" & item.LineNumber, label)
            Next
            tvOutline.ExpandAll()
            RebuildAutoCompleteList()
        End Sub

        Private Sub UpdateTitle()
            Dim title = APP_NAME
            If _activeFile >= 0 AndAlso _activeFile < _files.Count Then
                title = _files(_activeFile).FileName
                If _files(_activeFile).IsModified Then title &= " *"
                title &= " - " & APP_NAME
            End If
            Me.Text = title
        End Sub

        Private Sub UpdateEncodingUI()
            If _activeFile >= 0 AndAlso _activeFile < _files.Count Then
                lblEncoding.Text = GetEncodingName(_files(_activeFile).FileEnc)
                mnuEncANSI.Checked = (_files(_activeFile).FileEnc = FileEncoding.ANSI)
                mnuEncUTF8.Checked = (_files(_activeFile).FileEnc = FileEncoding.UTF8)
                mnuEncUTF8BOM.Checked = (_files(_activeFile).FileEnc = FileEncoding.UTF8_BOM)
            End If
            mnuEncDefaultANSI.Checked = (Settings.DefaultEncoding = FileEncoding.ANSI)
            mnuEncDefaultUTF8.Checked = (Settings.DefaultEncoding = FileEncoding.UTF8)
        End Sub

        Private Sub UpdateRecentFilesMenu()
            _mnuRecentParent.DropDownItems.Clear()
            If RecentFiles.Count = 0 Then
                _mnuRecentParent.DropDownItems.Add("(empty)").Enabled = False
            Else
                For i = 0 To RecentFiles.Count - 1
                    ' BUG FIX: capture loop variable properly with local copy
                    Dim capturedPath = RecentFiles(i)
                    Dim item = _mnuRecentParent.DropDownItems.Add($"&{i + 1} {capturedPath}")
                    AddHandler item.Click, Sub() OpenFileByPath(capturedPath)
                Next
            End If
        End Sub

        Private Sub ChangeFileEncoding(enc As FileEncoding)
            If _activeFile < 0 Then Return
            If _files(_activeFile).FileEnc = enc Then Return
            _files(_activeFile).FileEnc = enc
            If Not _files(_activeFile).IsModified Then
                _files(_activeFile).IsModified = True
                UpdateFileList() : UpdateTreeView() : UpdateTitle()
            End If
            UpdateEncodingUI()
            lblStatus.Text = "Encoding: " & GetEncodingName(enc)
        End Sub
#End Region

#Region "Scintilla Events"
        Private Sub Scintilla_SavePointLeft(sender As Object, e As EventArgs) Handles scintilla.SavePointLeft
            If _switchingFile Then Return
            If _activeFile >= 0 AndAlso _activeFile < _files.Count Then
                _files(_activeFile).IsModified = True
                UpdateTitle() : UpdateFileList() : UpdateTreeView()
            End If
        End Sub

        Private Sub Scintilla_SavePointReached(sender As Object, e As EventArgs) Handles scintilla.SavePointReached
            If _switchingFile Then Return
            If _activeFile >= 0 AndAlso _activeFile < _files.Count Then
                _files(_activeFile).IsModified = False
                UpdateTitle() : UpdateFileList() : UpdateTreeView()
            End If
        End Sub

        Private Sub Scintilla_UpdateUI(sender As Object, e As UpdateUIEventArgs) Handles scintilla.UpdateUI
            Dim line = scintilla.CurrentLine + 1
            Dim col = scintilla.GetColumn(scintilla.CurrentPosition) + 1
            lblPosition.Text = $"Ln: {line}  Col: {col}"
            lblLineCount.Text = $"Lines: {scintilla.Lines.Count}"

            ' Brace matching
            Dim pos = scintilla.CurrentPosition
            Dim braceFound = False
            If pos > 0 AndAlso IsBraceChar(scintilla.GetCharAt(pos - 1)) Then
                Dim match = scintilla.BraceMatch(pos - 1)
                If match >= 0 Then
                    scintilla.BraceHighlight(pos - 1, match)
                Else
                    scintilla.BraceBadLight(pos - 1)
                End If
                braceFound = True
            ElseIf pos < scintilla.TextLength AndAlso IsBraceChar(scintilla.GetCharAt(pos)) Then
                Dim match = scintilla.BraceMatch(pos)
                If match >= 0 Then
                    scintilla.BraceHighlight(pos, match)
                Else
                    scintilla.BraceBadLight(pos)
                End If
                braceFound = True
            End If
            If Not braceFound Then
                scintilla.BraceHighlight(ScintillaNET.Scintilla.InvalidPosition, ScintillaNET.Scintilla.InvalidPosition)
            End If
        End Sub

        Private Shared Function IsBraceChar(c As Integer) As Boolean
            Dim ch = ChrW(c)
            Return ch = "("c OrElse ch = ")"c OrElse ch = "["c OrElse ch = "]"c OrElse ch = "{"c OrElse ch = "}"c
        End Function

        Private Sub Scintilla_CharAdded(sender As Object, e As CharAddedEventArgs) Handles scintilla.CharAdded
            If Settings.AutoIndent AndAlso e.Char = 10 Then
                Dim curLine As Integer = scintilla.CurrentLine
                If curLine > 0 Then
                    Dim prevIndent As Integer = scintilla.Lines(curLine - 1).Indentation
                    If prevIndent > 0 Then
                        scintilla.Lines(curLine).Indentation = prevIndent
                        Dim linePos = scintilla.Lines(curLine).Position
                        Dim lineText = scintilla.Lines(curLine).Text
                        Dim indentChars = 0
                        While indentChars < lineText.Length
                            If lineText(indentChars) = " "c OrElse lineText(indentChars) = vbTab(0) Then indentChars += 1 Else Exit While
                        End While
                        scintilla.GotoPosition(linePos + indentChars)
                    End If
                End If
                Return
            End If
            If e.Char = 13 Then Return

            If Settings.AutoComplete Then
                Dim ch = ChrW(e.Char)
                If Char.IsLetterOrDigit(ch) OrElse ch = "_"c OrElse ch = "#"c Then ShowAutoComplete(False)
            End If
            _acTimer.Stop() : _acTimer.Start()
            If Settings.ShowFolding Then _foldTimer.Stop() : _foldTimer.Start()
        End Sub

        Private Sub Scintilla_MarginClick(sender As Object, e As MarginClickEventArgs) Handles scintilla.MarginClick
            Dim line = scintilla.LineFromPosition(e.Position)
            If e.Margin = 1 Then
                ' Check if Ctrl is held - toggle breakpoint, otherwise toggle bookmark
                If (Control.ModifierKeys And Keys.Control) <> 0 Then
                    ToggleBookmarkOnLine(line)
                Else
                    ' Toggle breakpoint
                    ToggleBreakpointOnLine(line)
                End If
            ElseIf e.Margin = 3 Then
                ' Fold margin click — toggle fold
                scintilla.Lines(line).ToggleFold()
            End If
        End Sub

        Private Sub Scintilla_ZoomChanged(sender As Object, e As EventArgs) Handles scintilla.ZoomChanged
            SetupMargins()
        End Sub
#End Region

#Region "Editor Operations"
        Private Sub ShowAutoComplete(force As Boolean)
            Dim pos = scintilla.CurrentPosition
            Dim wordStart = scintilla.WordStartPosition(pos, True)
            Dim lenEntered = pos - wordStart
            If force Then
                scintilla.AutoCShow(lenEntered, _autoCompleteList)
            ElseIf lenEntered >= 2 Then
                scintilla.AutoCShow(lenEntered, _autoCompleteList)
            End If
        End Sub

        Private Sub RebuildAutoCompleteList()
            Try
                If _activeFile < 0 OrElse _activeFile >= _files.Count Then Return
                Dim code As String = scintilla.Text
                If String.IsNullOrEmpty(code) Then
                    _autoCompleteList = _baseAutoCompleteList
                    Return
                End If

                Dim all As New HashSet(Of String)(_baseAutoCompleteList.Split({" "c}, StringSplitOptions.RemoveEmptyEntries), StringComparer.OrdinalIgnoreCase)

                ' Add outline items (subs, functions, types, enums, consts, variables)
                Dim items = CodeOutline.ParseOutline(code)
                For Each item In items
                    If item.Name.Length > 1 Then all.Add(item.Name)
                    ' Also add the data type if present
                    If item.DataType.Length > 1 Then all.Add(item.DataType)
                Next

                ' Scan ALL identifiers in the code for comprehensive auto-complete
                ' This catches: type members, enum values, #define names, local variables, etc.
                Dim identRegex As New System.Text.RegularExpressions.Regex("\b([A-Za-z_]\w{1,})\b")
                For Each m As System.Text.RegularExpressions.Match In identRegex.Matches(code)
                    Dim word = m.Groups(1).Value
                    ' Skip short words and pure numbers
                    If word.Length >= 2 Then all.Add(word)
                Next

                Dim sorted As New List(Of String)(all)
                sorted.Sort(StringComparer.OrdinalIgnoreCase)
                _autoCompleteList = String.Join(" ", sorted)
            Catch
            End Try
        End Sub

        Private Sub CommentBlock()
            Dim startLine = scintilla.LineFromPosition(scintilla.SelectionStart)
            Dim endLine = scintilla.LineFromPosition(scintilla.SelectionEnd)
            For i = startLine To endLine
                scintilla.InsertText(scintilla.Lines(i).Position, "'")
            Next
        End Sub

        Private Sub UncommentBlock()
            Dim startLine = scintilla.LineFromPosition(scintilla.SelectionStart)
            Dim endLine = scintilla.LineFromPosition(scintilla.SelectionEnd)
            For i = startLine To endLine
                Dim line = scintilla.Lines(i)
                Dim text = line.Text.TrimStart()
                If text.StartsWith("'") Then
                    Dim apostrophePos = line.Text.IndexOf("'"c)
                    If apostrophePos >= 0 Then
                        scintilla.TargetStart = line.Position + apostrophePos
                        scintilla.TargetEnd = line.Position + apostrophePos + 1
                        scintilla.ReplaceTarget("")
                    End If
                End If
            Next
        End Sub

        Private Sub ToggleBookmark()
            ToggleBookmarkOnLine(scintilla.CurrentLine)
        End Sub

        Private Sub ToggleBookmarkOnLine(line As Integer)
            Dim mask = scintilla.Lines(line).MarkerGet()
            If (mask And (1 << MARKER_BOOKMARK)) <> 0 Then
                scintilla.Lines(line).MarkerDelete(MARKER_BOOKMARK)
            Else
                scintilla.Lines(line).MarkerAdd(MARKER_BOOKMARK)
            End If
        End Sub

        ' BUG FIX: bounds-checked bookmark navigation
        Private Sub NextBookmark()
            Dim line = scintilla.CurrentLine
            Dim searchFrom = If(line + 1 < scintilla.Lines.Count, line + 1, 0)
            Dim next_ = scintilla.Lines(searchFrom).MarkerNext(1 << MARKER_BOOKMARK)
            If next_ < 0 Then next_ = scintilla.Lines(0).MarkerNext(1 << MARKER_BOOKMARK)
            If next_ >= 0 Then scintilla.GotoPosition(scintilla.Lines(next_).Position)
        End Sub

        Private Sub PrevBookmark()
            Dim line = scintilla.CurrentLine
            Dim searchFrom = If(line - 1 >= 0, line - 1, scintilla.Lines.Count - 1)
            Dim prev = scintilla.Lines(searchFrom).MarkerPrevious(1 << MARKER_BOOKMARK)
            If prev < 0 Then prev = scintilla.Lines(scintilla.Lines.Count - 1).MarkerPrevious(1 << MARKER_BOOKMARK)
            If prev >= 0 Then scintilla.GotoPosition(scintilla.Lines(prev).Position)
        End Sub

        ' BUG FIX: reuse FindReplaceForm to prevent memory leak
        Private Sub ShowFindReplace(showReplace As Boolean)
            If _findReplaceForm Is Nothing OrElse _findReplaceForm.IsDisposed Then
                _findReplaceForm = New FindReplaceForm(Me, scintilla, showReplace)
            End If
            _findReplaceForm.ShowWithText(scintilla.SelectedText, showReplace)
        End Sub

        Private Sub FindNext()
            If Not String.IsNullOrEmpty(_lastFindText) Then
                scintilla.SearchFlags = SearchFlags.None
                scintilla.TargetStart = scintilla.SelectionEnd
                scintilla.TargetEnd = scintilla.TextLength
                Dim pos = scintilla.SearchInTarget(_lastFindText)
                If pos >= 0 Then
                    scintilla.SetSel(scintilla.TargetStart, scintilla.TargetEnd)
                    scintilla.ScrollCaret()
                End If
            Else
                ShowFindReplace(False)
            End If
        End Sub

        Public Property LastFindText As String
            Get
                Return _lastFindText
            End Get
            Set(value As String)
                _lastFindText = value
            End Set
        End Property

        Private Sub GoToLine()
            Dim dlg As New GoToLineForm(scintilla.CurrentLine + 1)
            If dlg.ShowDialog(Me) = DialogResult.OK Then
                Dim lineNum = dlg.LineNumber - 1
                If lineNum >= 0 AndAlso lineNum < scintilla.Lines.Count Then
                    scintilla.GotoPosition(scintilla.Lines(lineNum).Position)
                    scintilla.ScrollCaret()
                End If
            End If
        End Sub

        Private Sub ChangeEditorFont()
            Using dlg As New FontDialog()
                dlg.FixedPitchOnly = True : dlg.MinSize = 6 : dlg.MaxSize = 72
                dlg.Font = New Font(Settings.EditorFont, Settings.EditorFontSize)
                If dlg.ShowDialog() = DialogResult.OK Then
                    Settings.EditorFont = dlg.Font.Name
                    Settings.EditorFontSize = CInt(dlg.Font.Size)
                    ApplyEditorTheme() : scintilla.Colorize(0, -1) : SaveSettings()
                    lblStatus.Text = $"Font: {Settings.EditorFont} {Settings.EditorFontSize}pt"
                End If
            End Using
        End Sub
#End Region

#Region "Build Operations"
        Private Sub DoCompile()
            If Not EnsureSavedForBuild() Then Return
            lblStatus.Text = "Compiling..."
            tabOutput.SelectedTab = tabPageOutput : txtOutput.Text = "" : Application.DoEvents()
            Dim result = BuildFile(_files(_activeFile).FilePath)
            txtOutput.Text = result.Output
            txtOutput.SelectionStart = txtOutput.TextLength

            ' Parse errors for double-click navigation
            Dim baseDir = If(_activeFile >= 0, Path.GetDirectoryName(_files(_activeFile).FilePath), "")
            _compilerErrors = ParseCompilerErrors(result.Output, baseDir)

            Dim errCount = _compilerErrors.FindAll(Function(e) e.ErrorType = "error").Count
            Dim warnCount = _compilerErrors.FindAll(Function(e) e.ErrorType = "warning").Count
            If result.Success Then
                lblStatus.Text = $"Build successful ({warnCount} warning(s))"
            Else
                lblStatus.Text = $"Build FAILED - {errCount} error(s), {warnCount} warning(s)"
                mnuViewOutputPanel.Checked = True
            End If
        End Sub

        Private Sub DoCompileAndRun()
            If Not EnsureSavedForBuild() Then Return
            lblStatus.Text = "Compiling and Running..."
            tabOutput.SelectedTab = tabPageOutput : txtOutput.Text = "" : Application.DoEvents()
            Dim result = BuildFile(_files(_activeFile).FilePath, True)
            txtOutput.Text = result.Output
            txtOutput.SelectionStart = txtOutput.TextLength

            Dim baseDir = If(_activeFile >= 0, Path.GetDirectoryName(_files(_activeFile).FilePath), "")
            _compilerErrors = ParseCompilerErrors(result.Output, baseDir)
            lblStatus.Text = If(result.Success, "Running...", $"Build FAILED (exit code {result.ExitCode})")
        End Sub

        Private Sub DoRunOnly()
            If _activeFile < 0 OrElse String.IsNullOrEmpty(_files(_activeFile).FilePath) Then Return
            Dim exePath = GetOutputExePath(_files(_activeFile).FilePath)
            If File.Exists(exePath) Then
                lblStatus.Text = "Running: " & Path.GetFileName(exePath)
                RunExecutable(exePath, Path.GetDirectoryName(_files(_activeFile).FilePath))
            Else
                MessageBox.Show("Executable not found. Please compile first.", APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

        Private Sub DoQuickRun()
            If Not EnsureSavedForBuild() Then Return
            lblStatus.Text = "Quick Run..."
            tabOutput.SelectedTab = tabPageOutput : txtOutput.Text = "" : Application.DoEvents()
            Dim result = QuickRun(_files(_activeFile).FilePath)
            txtOutput.Text = result.Output
            lblStatus.Text = If(result.Success, "Quick Run complete", "Quick Run FAILED")
        End Sub

        Private Sub DoSyntaxCheck()
            If Not EnsureSavedForBuild() Then Return
            lblStatus.Text = "Syntax Check..."
            tabOutput.SelectedTab = tabPageOutput : txtOutput.Text = "" : Application.DoEvents()
            Dim result = BuildFile(_files(_activeFile).FilePath, False, True)
            txtOutput.Text = result.Output
            lblStatus.Text = If(result.Success, "Syntax OK", "Syntax errors found")
        End Sub

        Private Function EnsureSavedForBuild() As Boolean
            If _activeFile < 0 Then Return False
            If _files(_activeFile).IsModified OrElse _files(_activeFile).IsNew Then
                If Not DoSaveFile() Then Return False
            End If
            Return Not String.IsNullOrEmpty(_files(_activeFile).FilePath)
        End Function

        Private Sub ShowBuildOptions()
            Dim dlg As New BuildOptionsForm()
            dlg.ShowDialog(Me)
        End Sub

        ''' <summary>Double-click output to navigate to error line</summary>
        Private Sub TxtOutput_DoubleClick(sender As Object, e As MouseEventArgs)
            If _compilerErrors.Count = 0 Then Return
            ' Find which line was clicked
            Dim charIdx = txtOutput.GetCharIndexFromPosition(e.Location)
            Dim lineIdx = txtOutput.GetLineFromCharIndex(charIdx)
            Dim lineText = ""
            If lineIdx >= 0 AndAlso lineIdx < txtOutput.Lines.Length Then
                lineText = txtOutput.Lines(lineIdx)
            End If
            If String.IsNullOrEmpty(lineText) Then Return

            ' Match against compiler errors
            For Each compErr In _compilerErrors
                If lineText.Contains(compErr.Message) OrElse lineText.Contains($"({compErr.LineNumber})") Then
                    ' Navigate to the error
                    If Not String.IsNullOrEmpty(compErr.FilePath) AndAlso File.Exists(compErr.FilePath) Then
                        OpenFileByPath(compErr.FilePath)
                    End If
                    If compErr.LineNumber > 0 AndAlso compErr.LineNumber <= scintilla.Lines.Count Then
                        scintilla.GotoPosition(scintilla.Lines(compErr.LineNumber - 1).Position)
                        scintilla.ScrollCaret() : scintilla.Focus()
                    End If
                    Exit For
                End If
            Next
        End Sub
#End Region

#Region "Debug Operations"
        Private Sub DoDebugStart()
            If _debugger.IsRunning Then
                If _debugger.IsPaused Then _debugger.Continue()
                Return
            End If

            ' Ensure file is saved and compiled with debug info
            If Not EnsureSavedForBuild() Then Return

            ' Auto-enable debug info if not set
            Dim origDebug = Build.DebugInfo
            Build.DebugInfo = True

            lblStatus.Text = "Compiling with debug info..."
            tabOutput.SelectedTab = tabPageOutput : txtOutput.Text = "" : Application.DoEvents()
            Dim result = BuildFile(_files(_activeFile).FilePath)
            txtOutput.Text = result.Output

            Build.DebugInfo = origDebug

            If Not result.Success Then
                Dim baseDir = Path.GetDirectoryName(_files(_activeFile).FilePath)
                _compilerErrors = ParseCompilerErrors(result.Output, baseDir)
                lblStatus.Text = "Build failed - cannot start debugger"
                Return
            End If

            ' Find GDB
            Dim gdbPath = Build.GDBPath
            If String.IsNullOrEmpty(gdbPath) Then gdbPath = GDBDebugger.FindGDBPath()
            If String.IsNullOrEmpty(gdbPath) Then
                MessageBox.Show("GDB not found. Please configure the GDB path in Debug > Set GDB Path.", APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim exePath = GetOutputExePath(_files(_activeFile).FilePath)
            Dim workDir = Path.GetDirectoryName(_files(_activeFile).FilePath)

            tabOutput.SelectedTab = tabPageDebugOutput
            txtDebugOutput.Text = ""
            mnuViewOutputPanel.Checked = True

            If _debugger.StartDebugging(gdbPath, exePath, _files(_activeFile).FilePath, workDir) Then
                _debugger.Run()
            End If
        End Sub

        Private Sub DoDebugStop()
            If _debugger.IsRunning Then
                _debugger.StopDebugging()
            End If
        End Sub

        Private Sub DoDebugPause()
            If _debugger.IsRunning AndAlso Not _debugger.IsPaused Then
                _debugger.Pause()
            End If
        End Sub

        Private Sub DoRunToCursor()
            If _debugger.IsRunning AndAlso _debugger.IsPaused AndAlso _activeFile >= 0 Then
                Dim fp = _files(_activeFile).FilePath
                Dim ln = scintilla.CurrentLine + 1
                _debugger.RunToCursor(fp, ln)
            End If
        End Sub

        Private Sub DoToggleBreakpoint()
            If _activeFile < 0 OrElse _activeFile >= _files.Count Then Return
            Dim line = scintilla.CurrentLine
            ToggleBreakpointOnLine(line)
        End Sub

        Private Sub ToggleBreakpointOnLine(line As Integer)
            If _activeFile < 0 OrElse _activeFile >= _files.Count Then Return
            Dim filePath = _files(_activeFile).FilePath
            If String.IsNullOrEmpty(filePath) Then filePath = _files(_activeFile).FileName

            Dim lineNum = line + 1  ' GDB uses 1-based
            Dim added = _debugger.ToggleBreakpoint(filePath, lineNum)
            Dim mask = scintilla.Lines(line).MarkerGet()

            If added Then
                scintilla.Lines(line).MarkerAdd(MARKER_BREAKPOINT)
            Else
                scintilla.Lines(line).MarkerDelete(MARKER_BREAKPOINT)
            End If
        End Sub

        Private Sub DoClearAllBreakpoints()
            _debugger.ClearAllBreakpoints()
            scintilla.MarkerDeleteAll(MARKER_BREAKPOINT)
        End Sub

        Private Sub DoAddWatch()
            Dim expr = InputBox("Enter watch expression:", "Add Watch", "")
            If Not String.IsNullOrEmpty(expr) Then
                _debugger.AddWatch(expr)
                ' Add to watch ListView immediately (value will update when paused)
                Dim lvi As New ListViewItem(expr)
                lvi.SubItems.Add(If(_debugger.IsPaused, "(evaluating...)", "(not running)"))
                lvi.SubItems.Add("")
                lvWatch.Items.Add(lvi)
                ' Request evaluation if debugger is paused
                If _debugger.IsRunning AndAlso _debugger.IsPaused Then
                    _debugger.RefreshWatches()
                End If
            End If
        End Sub

        Private Sub DoRemoveSelectedWatch()
            If lvWatch.SelectedItems.Count = 0 Then Return
            Dim item = lvWatch.SelectedItems(0)
            Dim expr = item.Text
            _debugger.RemoveWatch(expr)
            lvWatch.Items.Remove(item)
        End Sub

        Private Sub SetGDBPath()
            Using dlg As New OpenFileDialog()
                dlg.Filter = "GDB Debugger (gdb.exe)|gdb.exe|All Executables (*.exe)|*.exe"
                dlg.Title = "Locate GDB Debugger"
                If dlg.ShowDialog() = DialogResult.OK Then
                    Build.GDBPath = dlg.FileName
                    SaveSettings()
                    MessageBox.Show("GDB path set to:" & vbCrLf & dlg.FileName, APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Sub

        Private Sub UpdateDebugUI()
            Dim running = _debugger.IsRunning
            Dim paused = _debugger.IsPaused

            btnDebugStart.Text = If(running AndAlso paused, "▶ Continue", "▶ Debug")
            btnDebugStop.Enabled = running
            btnDebugPause.Enabled = running AndAlso Not paused
            btnDebugStepOver.Enabled = running AndAlso paused
            btnDebugStepInto.Enabled = running AndAlso paused
            btnDebugStepOut.Enabled = running AndAlso paused

            If running Then
                lblDebugState.Text = If(paused, "⏸ PAUSED", "▶ RUNNING")
                lblDebugState.ForeColor = If(paused, Color.Orange, Color.Green)
            Else
                lblDebugState.Text = ""
            End If
        End Sub

        Private Sub TxtGDBCommand_KeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Enter AndAlso _debugger.IsRunning Then
                Dim cmd = txtGDBCommand.Text.Trim()
                If Not String.IsNullOrEmpty(cmd) Then
                    _debugger.SendRawCommand(cmd)
                    txtGDBCommand.Text = ""
                End If
                e.Handled = True : e.SuppressKeyPress = True
            End If
        End Sub

        Private Sub LvCallStack_DoubleClick(sender As Object, e As EventArgs)
            If lvCallStack.SelectedItems.Count = 0 Then Return
            Dim item = lvCallStack.SelectedItems(0)
            Dim level = 0
            Integer.TryParse(item.SubItems(0).Text, level)
            _debugger.SelectFrame(level)

            ' Navigate to file/line
            Dim filePath = item.SubItems(2).Text
            Dim lineNum = 0
            Integer.TryParse(item.SubItems(3).Text, lineNum)
            If Not String.IsNullOrEmpty(filePath) AndAlso File.Exists(filePath) Then
                OpenFileByPath(filePath)
            End If
            If lineNum > 0 AndAlso lineNum <= scintilla.Lines.Count Then
                scintilla.GotoPosition(scintilla.Lines(lineNum - 1).Position)
                scintilla.ScrollCaret()
            End If
        End Sub
#End Region

#Region "Debugger Event Handlers"
        Private Sub Debugger_DebugStarted() Handles _debugger.DebugStarted
            UpdateDebugUI()
        End Sub

        Private Sub Debugger_DebugStopped() Handles _debugger.DebugStopped
            scintilla.MarkerDeleteAll(MARKER_DEBUGLINE)
            UpdateDebugUI()
        End Sub

        Private Sub Debugger_DebugPaused(filePath As String, lineNumber As Integer) Handles _debugger.DebugPaused
            ' Clear previous debug line marker
            scintilla.MarkerDeleteAll(MARKER_DEBUGLINE)

            ' Open file if needed and navigate
            If Not String.IsNullOrEmpty(filePath) AndAlso File.Exists(filePath) Then
                OpenFileByPath(filePath)
            End If

            If lineNumber > 0 AndAlso lineNumber <= scintilla.Lines.Count Then
                scintilla.Lines(lineNumber - 1).MarkerAdd(MARKER_DEBUGLINE)
                scintilla.GotoPosition(scintilla.Lines(lineNumber - 1).Position)
                scintilla.ScrollCaret()
            End If

            UpdateDebugUI()
        End Sub

        Private Sub Debugger_DebugResumed() Handles _debugger.DebugResumed
            scintilla.MarkerDeleteAll(MARKER_DEBUGLINE)
            UpdateDebugUI()
        End Sub

        Private Sub Debugger_DebugOutput(text As String) Handles _debugger.DebugOutput
            If txtDebugOutput.TextLength > 200000 Then txtDebugOutput.Text = txtDebugOutput.Text.Substring(100000)
            txtDebugOutput.AppendText(text.TrimEnd() & vbCrLf)
        End Sub

        Private Sub Debugger_DebugError(text As String) Handles _debugger.DebugError
            txtDebugOutput.AppendText("[ERROR] " & text & vbCrLf)
            lblStatus.Text = text
        End Sub

        Private Sub Debugger_LocalsUpdated(locals As List(Of GDBDebugger.VariableInfo)) Handles _debugger.LocalsUpdated
            lvLocals.BeginUpdate()
            lvLocals.Items.Clear()
            For Each v In locals
                Dim lvi As New ListViewItem(v.Name)
                lvi.SubItems.Add(v.Value)
                lvi.SubItems.Add(v.DataType)
                lvLocals.Items.Add(lvi)
            Next
            lvLocals.EndUpdate()
        End Sub

        Private Sub Debugger_WatchUpdated(watches As List(Of GDBDebugger.VariableInfo)) Handles _debugger.WatchUpdated
            lvWatch.BeginUpdate()
            lvWatch.Items.Clear()
            For Each v In watches
                Dim lvi As New ListViewItem(v.Name)
                lvi.SubItems.Add(v.Value)
                lvi.SubItems.Add(v.DataType)
                lvWatch.Items.Add(lvi)
            Next
            lvWatch.EndUpdate()
        End Sub

        Private Sub Debugger_CallStackUpdated(frames As List(Of GDBDebugger.StackFrameInfo)) Handles _debugger.CallStackUpdated
            lvCallStack.Items.Clear()
            For Each f In frames
                Dim lvi As New ListViewItem(f.Level.ToString())
                lvi.SubItems.Add(f.FunctionName)
                lvi.SubItems.Add(If(String.IsNullOrEmpty(f.FilePath), "", Path.GetFileName(f.FilePath)))
                lvi.SubItems.Add(If(f.LineNumber > 0, f.LineNumber.ToString(), ""))
                lvi.SubItems.Add(f.Address)
                lvCallStack.Items.Add(lvi)
            Next
        End Sub

        Private Sub Debugger_BreakpointHit(bpNumber As Integer, filePath As String, lineNumber As Integer) Handles _debugger.BreakpointHit
            tabOutput.SelectedTab = tabPageLocals
        End Sub
#End Region

#Region "AI Chat"
        Private Async Sub BtnAISend_Click(sender As Object, e As EventArgs) Handles btnAISend.Click
            Await SendAIMessage(False)
        End Sub

        Private Async Sub BtnAISendCode_Click(sender As Object, e As EventArgs) Handles btnAISendCode.Click
            Await SendAIMessage(True)
        End Sub

        Private Sub TxtAIInput_KeyDown(sender As Object, e As KeyEventArgs) Handles txtAIInput.KeyDown
            If e.KeyCode = Keys.Enter AndAlso e.Control Then
                BtnAISend_Click(Nothing, Nothing) : e.Handled = True : e.SuppressKeyPress = True
            End If
        End Sub

        Private Async Function SendAIMessage(includeCode As Boolean) As Task
            Dim userMsg = txtAIInput.Text.Trim()
            If String.IsNullOrEmpty(userMsg) Then Return

            Dim code = "" : Dim fileName = "Untitled.bas"
            If includeCode Then
                code = scintilla.SelectedText
                If String.IsNullOrEmpty(code) Then code = scintilla.Text
                If _activeFile >= 0 Then fileName = _files(_activeFile).FileName
            End If

            Dim displayMsg = userMsg
            If includeCode AndAlso code.Length > 0 Then displayMsg &= $"  [+ {code.Length} chars from {fileName}]"
            AppendAIChat("You", displayMsg)
            txtAIInput.Text = ""

            btnAISend.Enabled = False : btnAISendCode.Enabled = False : btnAISend.Text = "..."
            lblStatus.Text = "AI: Connecting to Claude..."
            AppendAIChat("", "Claude is thinking...")

            Dim response = Await _aiChat.SendMessageAsync(userMsg, includeCode, code, fileName)

            RemoveLastLine()
            AppendAIChat("Claude", response)
            btnAISend.Enabled = True : btnAISendCode.Enabled = True : btnAISend.Text = "Send"
            lblStatus.Text = "AI: Response received"
        End Function

        Private Sub AppendAIChat(sender As String, message As String)
            If String.IsNullOrEmpty(sender) Then
                txtAIChat.AppendText(message & vbCrLf)
            Else
                txtAIChat.AppendText($"--- {sender} ---" & vbCrLf & message & vbCrLf & vbCrLf)
            End If
            txtAIChat.ScrollToCaret()
        End Sub

        Private Sub RemoveLastLine()
            Dim text = txtAIChat.Text
            Dim lastNewline = text.LastIndexOf(vbCrLf, text.Length - 3)
            If lastNewline >= 0 Then txtAIChat.Text = text.Substring(0, lastNewline + 2)
        End Sub

        Private Sub BtnAIInsertCode_Click(sender As Object, e As EventArgs) Handles btnAIInsertCode.Click
            Dim code = AIChatManager.ExtractCodeFromResponse(_aiChat.LastResponse)
            If String.IsNullOrEmpty(code) Then
                MessageBox.Show("No code block found.", APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            scintilla.ReplaceSelection(code) : lblStatus.Text = $"AI: Code inserted ({code.Length} chars)"
        End Sub

        Private Sub BtnAICopyReply_Click(sender As Object, e As EventArgs) Handles btnAICopyReply.Click
            If String.IsNullOrEmpty(_aiChat.LastResponse) Then Return
            Clipboard.SetText(_aiChat.LastResponse) : lblStatus.Text = "AI: Response copied"
        End Sub

        Private Sub BtnAIReplaceAll_Click(sender As Object, e As EventArgs) Handles btnAIReplaceAll.Click
            Dim code = AIChatManager.ExtractCodeFromResponse(_aiChat.LastResponse)
            If String.IsNullOrEmpty(code) Then code = _aiChat.LastResponse
            If MessageBox.Show("Replace ALL code in editor with AI-generated code?", APP_NAME, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                scintilla.Text = code : lblStatus.Text = "AI: Editor content replaced"
            End If
        End Sub

        Private Sub BtnAINewFile_Click(sender As Object, e As EventArgs) Handles btnAINewFile.Click
            Dim code = AIChatManager.ExtractCodeFromResponse(_aiChat.LastResponse)
            If String.IsNullOrEmpty(code) Then code = _aiChat.LastResponse
            DoNewFile() : scintilla.Text = code : lblStatus.Text = "AI: Code placed in new file"
        End Sub

        Private Sub BtnAIClearChat_Click(sender As Object, e As EventArgs) Handles btnAIClearChat.Click
            _aiChat.ClearHistory() : txtAIChat.Text = "" : lblStatus.Text = "AI: Chat cleared"
        End Sub
#End Region

#Region "Tools / Help"
        Private Sub SetFBCPath()
            Using dlg As New OpenFileDialog()
                dlg.Filter = "FreeBASIC Compiler (fbc*.exe)|fbc*.exe|All Executables (*.exe)|*.exe"
                dlg.Title = "Locate FreeBASIC Compiler"
                If dlg.ShowDialog() = DialogResult.OK Then
                    Build.FBCPath = dlg.FileName : SaveSettings()
                    MessageBox.Show("Compiler set to:" & vbCrLf & dlg.FileName, APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Sub

        Private Sub OpenFBHelp()
            If Not String.IsNullOrEmpty(Build.FBDocPath) AndAlso File.Exists(Build.FBDocPath) Then
                SafeProcessStart(Build.FBDocPath) : Return
            End If
            If Not String.IsNullOrEmpty(Build.FBCPath) Then
                Dim docDir = Path.GetDirectoryName(Path.GetDirectoryName(Build.FBCPath))
                For Each helpName In {"fbhelp.chm", "FreeBASIC-Manual.chm"}
                    Dim p = Path.Combine(docDir, "doc", helpName)
                    If File.Exists(p) Then
                        SafeProcessStart(p)
                        Return
                    End If
                Next
            End If
            If MessageBox.Show("Local help not found. Open online documentation?", APP_NAME,
                              MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                SafeProcessStart("https://www.freebasic.net/wiki")
            End If
        End Sub

        Private Sub ShowAbout()
            Dim dlg As New AboutForm()
            dlg.ShowDialog(Me)
        End Sub
#End Region

#Region "Form Events"
        Private Sub CboOpenFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboOpenFiles.SelectedIndexChanged
            If _switchingFile Then Return
            If cboOpenFiles.SelectedIndex <> _activeFile Then SwitchToFile(cboOpenFiles.SelectedIndex)
        End Sub

        Private Sub TvProject_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles tvProject.NodeMouseDoubleClick
            If e.Node.Name.StartsWith("file") Then
                Dim idx As Integer
                If Integer.TryParse(e.Node.Name.Substring(4), idx) Then SwitchToFile(idx)
            End If
        End Sub

        Private Sub TvOutline_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles tvOutline.NodeMouseDoubleClick
            If e.Node.Name.StartsWith("item_") Then
                Dim lineNum As Integer
                If Integer.TryParse(e.Node.Name.Substring(5), lineNum) Then
                    lineNum -= 1
                    If lineNum >= 0 AndAlso lineNum < scintilla.Lines.Count Then
                        scintilla.GotoPosition(scintilla.Lines(lineNum).Position)
                        scintilla.ScrollCaret() : scintilla.Focus()
                    End If
                End If
            End If
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            ' Stop debugger first
            If _debugger.IsRunning Then _debugger.StopDebugging()

            For Each f In _files
                If f.IsModified Then
                    Dim r = MessageBox.Show($"Save changes to {f.FileName}?", APP_NAME, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                    If r = DialogResult.Yes Then
                        Dim idx = _files.IndexOf(f) : SwitchToFile(idx)
                        If Not DoSaveFile() Then
                            e.Cancel = True
                            Return
                        End If
                    ElseIf r = DialogResult.Cancel Then
                        e.Cancel = True
                        Return
                    End If
                End If
            Next
            SaveSettings() : _acTimer.Dispose() : _debugger.Dispose()
            If _findReplaceForm IsNot Nothing Then _findReplaceForm.Dispose()
            MyBase.OnFormClosing(e)
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            Try
                splitMain.SplitterDistance = 250
                splitEditor.SplitterDistance = CInt(splitEditor.Height * 0.75)
            Catch
            End Try
        End Sub

        ' ENHANCEMENT: Ctrl+Tab / Ctrl+Shift+Tab file switching
        Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
            ' Route clipboard shortcuts to the focused control (not always Scintilla)
            ' This fixes copy/paste in AI Chat, output panel, etc.
            If keyData = (Keys.Control Or Keys.C) OrElse
               keyData = (Keys.Control Or Keys.V) OrElse
               keyData = (Keys.Control Or Keys.X) OrElse
               keyData = (Keys.Control Or Keys.A) Then
                Dim focused = GetFocusedControl(Me)
                If focused IsNot Nothing AndAlso Not TypeOf focused Is ScintillaNET.Scintilla Then
                    ' Let the focused non-Scintilla control handle the shortcut natively
                    Return MyBase.ProcessCmdKey(msg, keyData)
                End If
            End If

            If keyData = (Keys.Control Or Keys.Tab) Then
                If _files.Count > 1 Then
                    Dim next_ = (_activeFile + 1) Mod _files.Count
                    SwitchToFile(next_)
                End If
                Return True
            ElseIf keyData = (Keys.Control Or Keys.Shift Or Keys.Tab) Then
                If _files.Count > 1 Then
                    Dim prev = If(_activeFile - 1 < 0, _files.Count - 1, _activeFile - 1)
                    SwitchToFile(prev)
                End If
                Return True
            End If
            Return MyBase.ProcessCmdKey(msg, keyData)
        End Function

        ''' <summary>Walk the control tree to find the actually focused control.</summary>
        Private Shared Function GetFocusedControl(container As Control) As Control
            For Each child As Control In container.Controls
                If child.Focused Then Return child
                If child.ContainsFocus Then
                    Dim inner = GetFocusedControl(child)
                    If inner IsNot Nothing Then Return inner
                End If
            Next
            Return Nothing
        End Function

        Protected Overrides Sub OnDragEnter(e As DragEventArgs)
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then e.Effect = DragDropEffects.Copy
        End Sub

        Protected Overrides Sub OnDragDrop(e As DragEventArgs)
            Dim files = CType(e.Data.GetData(DataFormats.FileDrop), String())
            If files IsNot Nothing Then
                For Each f In files
                    OpenFileByPath(f)
                Next
            End If
        End Sub
#End Region
    End Class

