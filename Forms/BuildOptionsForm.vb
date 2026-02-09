Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

    ''' <summary>
    ''' Build Options / Configuration Dialog
    ''' Ported from VB6 frmBuildOptions.frm
    ''' </summary>
    Public Class BuildOptionsForm
        Inherits Form

        ' Tab control & pages
        Private tabOptions As TabControl
        Private tabCompiler As TabPage
        Private tabDebug As TabPage
        Private tabPaths As TabPage

        ' Compiler tab controls
        Private cboTargetType As ComboBox
        Private cboOptimization As ComboBox
        Private cboErrorCheck As ComboBox
        Private cboDialect As ComboBox
        Private cboCodeGen As ComboBox
        Private cboWarnings As ComboBox
        Private cboArch As ComboBox
        Private cboFPU As ComboBox
        Private txtStackSize As TextBox
        Private txtExtraCompiler As TextBox

        ' Debug tab controls
        Private chkDebugInfo As CheckBox
        Private chkVerbose As CheckBox
        Private chkShowCommands As CheckBox
        Private chkGenMap As CheckBox
        Private chkEmitASM As CheckBox
        Private chkKeepIntermediate As CheckBox

        ' Paths tab controls
        Private txtFBCPath As TextBox
        Private txtFBC32Path As TextBox
        Private txtFBC64Path As TextBox
        Private txtFBDocPath As TextBox
        Private txtAPIKeyPath As TextBox
        Private txtOutputFile As TextBox
        Private txtIncludePaths As TextBox
        Private txtLibraryPaths As TextBox
        Private txtExtraLinker As TextBox

        ' Bottom buttons & preview
        Private lblCmdPreview As Label
        Private btnOK As Button
        Private btnCancel As Button
        Private btnApply As Button

        Public Sub New()
            InitializeComponent()
            PopulateCombos()
            LoadFromSettings()
        End Sub

        Private Sub InitializeComponent()
            Me.Text = "Build Options"
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowInTaskbar = False
            Me.StartPosition = FormStartPosition.CenterParent
            Me.ClientSize = New Size(540, 530)
            Me.Font = New Font("Segoe UI", 9)

            ' === Tab Control ===
            tabOptions = New TabControl() With {
                .Location = New Point(8, 8),
                .Size = New Size(524, 430),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
            }

            ' === Compiler Tab ===
            tabCompiler = New TabPage("Compiler")
            CreateCompilerTab()
            tabOptions.TabPages.Add(tabCompiler)

            ' === Debug/Output Tab ===
            tabDebug = New TabPage("Debug / Output")
            CreateDebugTab()
            tabOptions.TabPages.Add(tabDebug)

            ' === Paths Tab ===
            tabPaths = New TabPage("Paths")
            CreatePathsTab()
            tabOptions.TabPages.Add(tabPaths)

            Me.Controls.Add(tabOptions)

            ' === Command Preview ===
            lblCmdPreview = New Label() With {
                .Location = New Point(8, 444),
                .Size = New Size(300, 36),
                .Font = New Font("Consolas", 8.25F),
                .BackColor = Color.FromArgb(240, 240, 240),
                .BorderStyle = BorderStyle.FixedSingle,
                .AutoEllipsis = True,
                .Text = "fbc <source.bas>"
            }
            Me.Controls.Add(lblCmdPreview)

            ' === Buttons ===
            btnOK = New Button() With {
                .Text = "OK",
                .Size = New Size(80, 30),
                .Location = New Point(310, 448),
                .DialogResult = DialogResult.OK
            }
            AddHandler btnOK.Click, AddressOf BtnOK_Click
            Me.Controls.Add(btnOK)

            btnCancel = New Button() With {
                .Text = "Cancel",
                .Size = New Size(80, 30),
                .Location = New Point(395, 448),
                .DialogResult = DialogResult.Cancel
            }
            Me.Controls.Add(btnCancel)

            btnApply = New Button() With {
                .Text = "Apply",
                .Size = New Size(70, 30),
                .Location = New Point(480, 448)
            }
            AddHandler btnApply.Click, AddressOf BtnApply_Click
            Me.Controls.Add(btnApply)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub

        Private Sub CreateCompilerTab()
            Dim y As Integer = 12
            Dim lblW As Integer = 110
            Dim cboW As Integer = 220
            Dim cboX As Integer = 124

            ' Target Type
            tabCompiler.Controls.Add(MakeLabel("Target Type:", 12, y + 3))
            cboTargetType = MakeCombo(cboX, y, cboW)
            tabCompiler.Controls.Add(cboTargetType)
            y += 32

            ' Optimization
            tabCompiler.Controls.Add(MakeLabel("Optimization:", 12, y + 3))
            cboOptimization = MakeCombo(cboX, y, cboW)
            tabCompiler.Controls.Add(cboOptimization)
            y += 32

            ' Error Checking
            tabCompiler.Controls.Add(MakeLabel("Error Checking:", 12, y + 3))
            cboErrorCheck = MakeCombo(cboX, y, cboW)
            tabCompiler.Controls.Add(cboErrorCheck)
            y += 32

            ' Dialect
            tabCompiler.Controls.Add(MakeLabel("Dialect:", 12, y + 3))
            cboDialect = MakeCombo(cboX, y, cboW)
            tabCompiler.Controls.Add(cboDialect)
            y += 32

            ' Code Generator
            tabCompiler.Controls.Add(MakeLabel("Code Generator:", 12, y + 3))
            cboCodeGen = MakeCombo(cboX, y, cboW)
            tabCompiler.Controls.Add(cboCodeGen)
            y += 32

            ' Warnings
            tabCompiler.Controls.Add(MakeLabel("Warnings:", 12, y + 3))
            cboWarnings = MakeCombo(cboX, y, cboW)
            tabCompiler.Controls.Add(cboWarnings)
            y += 32

            ' Architecture
            tabCompiler.Controls.Add(MakeLabel("Architecture:", 12, y + 3))
            cboArch = MakeCombo(cboX, y, cboW)
            tabCompiler.Controls.Add(cboArch)
            y += 32

            ' FPU
            tabCompiler.Controls.Add(MakeLabel("FPU:", 12, y + 3))
            cboFPU = MakeCombo(cboX, y, cboW)
            tabCompiler.Controls.Add(cboFPU)
            y += 32

            ' Stack Size
            tabCompiler.Controls.Add(MakeLabel("Stack Size:", 12, y + 3))
            txtStackSize = New TextBox() With {
                .Location = New Point(cboX, y),
                .Size = New Size(cboW, 23),
                .Text = "0"
            }
            tabCompiler.Controls.Add(txtStackSize)
            y += 32

            ' Extra Options
            tabCompiler.Controls.Add(MakeLabel("Extra Options:", 12, y + 3))
            txtExtraCompiler = New TextBox() With {
                .Location = New Point(cboX, y),
                .Size = New Size(370, 52),
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical
            }
            tabCompiler.Controls.Add(txtExtraCompiler)

            ' Wire up change events for preview
            AddHandler cboTargetType.SelectedIndexChanged, AddressOf SettingChanged
            AddHandler cboOptimization.SelectedIndexChanged, AddressOf SettingChanged
            AddHandler cboErrorCheck.SelectedIndexChanged, AddressOf SettingChanged
            AddHandler cboDialect.SelectedIndexChanged, AddressOf SettingChanged
            AddHandler cboCodeGen.SelectedIndexChanged, AddressOf SettingChanged
            AddHandler cboWarnings.SelectedIndexChanged, AddressOf SettingChanged
            AddHandler cboArch.SelectedIndexChanged, AddressOf SettingChanged
            AddHandler cboFPU.SelectedIndexChanged, AddressOf SettingChanged
        End Sub

        Private Sub CreateDebugTab()
            Dim y As Integer = 16

            chkDebugInfo = MakeCheckBox("Generate debug information (-g)", 16, y) : y += 30
            chkVerbose = MakeCheckBox("Verbose compilation (-v)", 16, y) : y += 30
            chkShowCommands = MakeCheckBox("Show commands (-showincludes)", 16, y) : y += 30
            chkGenMap = MakeCheckBox("Generate map file (-map)", 16, y) : y += 30
            chkEmitASM = MakeCheckBox("Emit assembly output (-R)", 16, y) : y += 30
            chkKeepIntermediate = MakeCheckBox("Keep intermediate files (-C)", 16, y)

            tabDebug.Controls.AddRange({chkDebugInfo, chkVerbose, chkShowCommands,
                                         chkGenMap, chkEmitASM, chkKeepIntermediate})

            AddHandler chkDebugInfo.CheckedChanged, AddressOf SettingChanged
            AddHandler chkVerbose.CheckedChanged, AddressOf SettingChanged
        End Sub

        Private Sub CreatePathsTab()
            Dim y As Integer = 8
            Dim txW As Integer = 380
            Dim btnX As Integer = 400

            ' FBC Path
            tabPaths.Controls.Add(MakeLabel("FreeBASIC Compiler (fbc32.exe / fbc.exe):", 8, y))
            y += 18
            txtFBCPath = New TextBox() With {.Location = New Point(8, y), .Size = New Size(txW, 23)}
            tabPaths.Controls.Add(txtFBCPath)
            Dim btnBrowseFBC As New Button() With {.Text = "...", .Size = New Size(32, 23), .Location = New Point(btnX, y)}
            AddHandler btnBrowseFBC.Click, Sub() BrowseForExe(txtFBCPath, "FreeBASIC Compiler (fbc*.exe)|fbc*.exe;fbc32.exe;fbc64.exe|All (*.exe)|*.exe")
            tabPaths.Controls.Add(btnBrowseFBC)
            Dim btnAutoDetect As New Button() With {.Text = "Auto", .Size = New Size(52, 23), .Location = New Point(436, y)}
            AddHandler btnAutoDetect.Click, AddressOf AutoDetectFBC
            tabPaths.Controls.Add(btnAutoDetect)
            y += 30

            ' FBC32 Path
            tabPaths.Controls.Add(MakeLabel("FBC 32-bit compiler (fbc32.exe) - optional:", 8, y))
            y += 18
            txtFBC32Path = New TextBox() With {.Location = New Point(8, y), .Size = New Size(txW, 23)}
            tabPaths.Controls.Add(txtFBC32Path)
            Dim btnBrowse32 As New Button() With {.Text = "...", .Size = New Size(32, 23), .Location = New Point(btnX, y)}
            AddHandler btnBrowse32.Click, Sub() BrowseForExe(txtFBC32Path, "fbc32.exe|fbc32.exe|All (*.exe)|*.exe")
            tabPaths.Controls.Add(btnBrowse32)
            y += 30

            ' FBC64 Path
            tabPaths.Controls.Add(MakeLabel("FBC 64-bit compiler (fbc64.exe) - optional:", 8, y))
            y += 18
            txtFBC64Path = New TextBox() With {.Location = New Point(8, y), .Size = New Size(txW, 23)}
            tabPaths.Controls.Add(txtFBC64Path)
            Dim btnBrowse64 As New Button() With {.Text = "...", .Size = New Size(32, 23), .Location = New Point(btnX, y)}
            AddHandler btnBrowse64.Click, Sub() BrowseForExe(txtFBC64Path, "fbc64.exe|fbc64.exe|All (*.exe)|*.exe")
            tabPaths.Controls.Add(btnBrowse64)
            y += 30

            ' FB Documentation Path
            tabPaths.Controls.Add(MakeLabel("FreeBASIC Documentation (.chm) - F1 to open:", 8, y))
            y += 18
            txtFBDocPath = New TextBox() With {.Location = New Point(8, y), .Size = New Size(txW, 23)}
            tabPaths.Controls.Add(txtFBDocPath)
            Dim btnBrowseDoc As New Button() With {.Text = "...", .Size = New Size(32, 23), .Location = New Point(btnX, y)}
            AddHandler btnBrowseDoc.Click, Sub() BrowseForFile(txtFBDocPath, "CHM Files (*.chm)|*.chm|All Files (*.*)|*.*")
            tabPaths.Controls.Add(btnBrowseDoc)
            y += 30

            ' API Key File Path
            tabPaths.Controls.Add(MakeLabel("Anthropic API Key File (.txt):", 8, y))
            y += 18
            txtAPIKeyPath = New TextBox() With {.Location = New Point(8, y), .Size = New Size(txW, 23)}
            tabPaths.Controls.Add(txtAPIKeyPath)
            Dim btnBrowseAPI As New Button() With {.Text = "...", .Size = New Size(32, 23), .Location = New Point(btnX, y)}
            AddHandler btnBrowseAPI.Click, Sub() BrowseForFile(txtAPIKeyPath, "Text Files (*.txt)|*.txt|All Files (*.*)|*.*")
            tabPaths.Controls.Add(btnBrowseAPI)
            y += 34

            ' Output Filename
            tabPaths.Controls.Add(MakeLabel("Output Filename (leave blank for default):", 8, y))
            y += 18
            txtOutputFile = New TextBox() With {.Location = New Point(8, y), .Size = New Size(480, 23)}
            tabPaths.Controls.Add(txtOutputFile)
            y += 30

            ' Include Paths
            tabPaths.Controls.Add(MakeLabel("Include Paths (one per line):", 8, y))
            y += 18
            txtIncludePaths = New TextBox() With {
                .Location = New Point(8, y), .Size = New Size(480, 44),
                .Multiline = True, .ScrollBars = ScrollBars.Vertical
            }
            tabPaths.Controls.Add(txtIncludePaths)
            y += 50

            ' Library Paths
            tabPaths.Controls.Add(MakeLabel("Library Paths (one per line):", 8, y))
            y += 18
            txtLibraryPaths = New TextBox() With {
                .Location = New Point(8, y), .Size = New Size(480, 44),
                .Multiline = True, .ScrollBars = ScrollBars.Vertical
            }
            tabPaths.Controls.Add(txtLibraryPaths)
            y += 50

            ' Extra Linker Options
            tabPaths.Controls.Add(MakeLabel("Extra Linker Options:", 8, y))
            y += 18
            txtExtraLinker = New TextBox() With {
                .Location = New Point(8, y), .Size = New Size(480, 38),
                .Multiline = True
            }
            tabPaths.Controls.Add(txtExtraLinker)
        End Sub

        Private Sub PopulateCombos()
            cboTargetType.Items.AddRange({"Console Application (.exe)", "GUI Application (.exe)",
                                           "Dynamic Library (.dll)", "Static Library (.a)"})

            cboOptimization.Items.AddRange({"None (default)", "-O 1 (Basic)", "-O 2 (Standard)", "-O 3 (Maximum)"})

            cboErrorCheck.Items.AddRange({"None (default)", "-e (Basic)", "-ex (With RESUME support)",
                                           "-exx (Full - Array bounds + null ptr)"})

            cboDialect.Items.AddRange({"fb (FreeBASIC default)", "qb (QuickBASIC compatible)",
                                        "fblite (QB-like with FB features)", "deprecated (Old FB syntax)"})

            cboCodeGen.Items.AddRange({"gas (GNU Assembler - default)", "gcc (GNU C Compiler backend)",
                                        "llvm (LLVM backend)"})

            cboWarnings.Items.AddRange({"None", "All (-w all)", "Pedantic (-w pedantic)"})

            cboArch.Items.AddRange({"32-bit (x86)", "64-bit (x86_64)"})

            cboFPU.Items.AddRange({"x87 (default)", "SSE"})
        End Sub

        Private Sub LoadFromSettings()
            Dim b = Build

            SafeSetIndex(cboTargetType, b.TargetType)
            SafeSetIndex(cboOptimization, b.Optimization)
            SafeSetIndex(cboErrorCheck, b.ErrorChecking)
            SafeSetIndex(cboDialect, b.LangDialect)
            SafeSetIndex(cboCodeGen, b.CodeGen)
            SafeSetIndex(cboWarnings, b.Warnings)
            SafeSetIndex(cboArch, b.TargetArch)
            SafeSetIndex(cboFPU, b.FPU)

            txtStackSize.Text = b.StackSize.ToString()
            txtExtraCompiler.Text = b.ExtraCompilerOpts

            chkDebugInfo.Checked = b.DebugInfo
            chkVerbose.Checked = b.Verbose
            chkShowCommands.Checked = b.ShowCommands
            chkGenMap.Checked = b.GenerateMap
            chkEmitASM.Checked = b.EmitASM
            chkKeepIntermediate.Checked = b.KeepIntermediate

            txtFBCPath.Text = b.FBCPath
            txtFBC32Path.Text = b.FBC32Path
            txtFBC64Path.Text = b.FBC64Path
            txtFBDocPath.Text = b.FBDocPath
            txtAPIKeyPath.Text = b.APIKeyFilePath
            txtOutputFile.Text = b.OutputFile
            txtIncludePaths.Text = b.IncludePaths.Replace(";", Environment.NewLine)
            txtLibraryPaths.Text = b.LibraryPaths.Replace(";", Environment.NewLine)
            txtExtraLinker.Text = b.ExtraLinkerOpts

            UpdatePreview()
        End Sub

        Private Sub SaveToSettings()
            Dim b = Build

            b.TargetType = cboTargetType.SelectedIndex
            b.Optimization = cboOptimization.SelectedIndex
            b.ErrorChecking = cboErrorCheck.SelectedIndex
            b.LangDialect = cboDialect.SelectedIndex
            b.CodeGen = cboCodeGen.SelectedIndex
            b.Warnings = cboWarnings.SelectedIndex
            b.TargetArch = cboArch.SelectedIndex
            b.FPU = cboFPU.SelectedIndex

            Dim stack As Integer
            If Integer.TryParse(txtStackSize.Text, stack) Then b.StackSize = stack Else b.StackSize = 0

            b.ExtraCompilerOpts = txtExtraCompiler.Text

            b.DebugInfo = chkDebugInfo.Checked
            b.Verbose = chkVerbose.Checked
            b.ShowCommands = chkShowCommands.Checked
            b.GenerateMap = chkGenMap.Checked
            b.EmitASM = chkEmitASM.Checked
            b.KeepIntermediate = chkKeepIntermediate.Checked

            b.FBCPath = txtFBCPath.Text
            b.FBC32Path = txtFBC32Path.Text
            b.FBC64Path = txtFBC64Path.Text
            b.FBDocPath = txtFBDocPath.Text
            b.APIKeyFilePath = txtAPIKeyPath.Text
            b.OutputFile = txtOutputFile.Text
            b.IncludePaths = txtIncludePaths.Text.Replace(Environment.NewLine, ";").Replace(vbCr, ";")
            b.LibraryPaths = txtLibraryPaths.Text.Replace(Environment.NewLine, ";").Replace(vbCr, ";")
            b.ExtraLinkerOpts = txtExtraLinker.Text

            SaveSettings()
        End Sub

        Private Sub UpdatePreview()
            Dim cmd As String = "fbc"

            Select Case cboTargetType.SelectedIndex
                Case 1 : cmd &= " -s gui"
                Case 2 : cmd &= " -dll"
                Case 3 : cmd &= " -lib"
            End Select

            If cboOptimization.SelectedIndex > 0 Then cmd &= " -O " & cboOptimization.SelectedIndex
            Select Case cboErrorCheck.SelectedIndex
                Case 1 : cmd &= " -e"
                Case 2 : cmd &= " -ex"
                Case 3 : cmd &= " -exx"
            End Select

            Dim dialects() = {"fb", "qb", "fblite", "deprecated"}
            If cboDialect.SelectedIndex > 0 Then cmd &= " -lang " & dialects(cboDialect.SelectedIndex)

            If chkDebugInfo.Checked Then cmd &= " -g"
            If chkVerbose.Checked Then cmd &= " -v"

            cmd &= " <source.bas>"
            lblCmdPreview.Text = cmd
        End Sub

        Private Sub SettingChanged(sender As Object, e As EventArgs)
            UpdatePreview()
        End Sub

        Private Sub BtnOK_Click(sender As Object, e As EventArgs)
            SaveToSettings()
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub

        Private Sub BtnApply_Click(sender As Object, e As EventArgs)
            SaveToSettings()
        End Sub

        Private Sub AutoDetectFBC(sender As Object, e As EventArgs)
            Dim path = FindFBCPath()
            If Not String.IsNullOrEmpty(path) Then
                txtFBCPath.Text = path
                MessageBox.Show("Found FreeBASIC compiler at:" & Environment.NewLine & path,
                                APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Could not auto-detect FreeBASIC compiler." & Environment.NewLine &
                                "Please browse for fbc32.exe or fbc.exe manually.",
                                APP_NAME, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        End Sub

        Private Sub BrowseForExe(txt As TextBox, filter As String)
            Using dlg As New OpenFileDialog()
                dlg.Filter = filter
                dlg.Title = "Locate FreeBASIC Compiler"
                If Not String.IsNullOrEmpty(txt.Text) AndAlso Directory.Exists(Path.GetDirectoryName(txt.Text)) Then
                    dlg.InitialDirectory = Path.GetDirectoryName(txt.Text)
                End If
                If dlg.ShowDialog() = DialogResult.OK Then
                    txt.Text = dlg.FileName
                End If
            End Using
        End Sub

        Private Sub BrowseForFile(txt As TextBox, filter As String)
            Using dlg As New OpenFileDialog()
                dlg.Filter = filter
                If Not String.IsNullOrEmpty(txt.Text) AndAlso Directory.Exists(Path.GetDirectoryName(txt.Text)) Then
                    dlg.InitialDirectory = Path.GetDirectoryName(txt.Text)
                End If
                If dlg.ShowDialog() = DialogResult.OK Then
                    txt.Text = dlg.FileName
                End If
            End Using
        End Sub

        ' === UI Helpers ===

        Private Shared Function MakeLabel(text As String, x As Integer, y As Integer) As Label
            Return New Label() With {
                .Text = text,
                .Location = New Point(x, y),
                .AutoSize = True
            }
        End Function

        Private Shared Function MakeCombo(x As Integer, y As Integer, w As Integer) As ComboBox
            Dim cbo As New ComboBox() With {
                .Location = New Point(x, y),
                .Size = New Size(w, 23),
                .DropDownStyle = ComboBoxStyle.DropDownList
            }
            Return cbo
        End Function

        Private Shared Function MakeCheckBox(text As String, x As Integer, y As Integer) As CheckBox
            Return New CheckBox() With {
                .Text = text,
                .Location = New Point(x, y),
                .AutoSize = True
            }
        End Function

        Private Shared Sub SafeSetIndex(cbo As ComboBox, index As Integer)
            If index >= 0 AndAlso index < cbo.Items.Count Then
                cbo.SelectedIndex = index
            ElseIf cbo.Items.Count > 0 Then
                cbo.SelectedIndex = 0
            End If
        End Sub
    End Class
