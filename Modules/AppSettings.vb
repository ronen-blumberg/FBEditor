Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

    ''' <summary>
    ''' File encoding types
    ''' </summary>
    Public Enum FileEncoding
        ANSI = 0
        UTF8 = 1
        UTF8_BOM = 2
    End Enum

    ''' <summary>
    ''' Open file information for multi-tab support
    ''' </summary>
    Public Class OpenFileInfo
        Public FilePath As String = ""
        Public FileName As String = ""
        Public IsModified As Boolean = False
        Public Content As String = ""
        Public FirstVisibleLine As Integer = 0
        Public CursorPos As Integer = 0
        Public IsNew As Boolean = True
        Public FileEnc As FileEncoding = FileEncoding.UTF8
    End Class

    ''' <summary>
    ''' Parsed compiler error/warning
    ''' </summary>
    Public Class CompilerError
        Public FilePath As String = ""
        Public LineNumber As Integer = 0
        Public ErrorType As String = ""   ' "error" or "warning"
        Public ErrorCode As Integer = 0
        Public Message As String = ""

        Public Overrides Function ToString() As String
            Return $"{Path.GetFileName(FilePath)}({LineNumber}) {ErrorType} {ErrorCode}: {Message}"
        End Function
    End Class

    ''' <summary>
    ''' Build configuration settings
    ''' </summary>
    Public Class BuildSettings
        Public FBCPath As String = ""
        Public FBC32Path As String = ""
        Public FBC64Path As String = ""
        Public FBDocPath As String = ""
        Public APIKeyFilePath As String = ""
        Public GDBPath As String = ""          ' GDB debugger path
        Public TargetType As Integer = 0       ' 0=Console, 1=GUI, 2=DLL, 3=Static Lib
        Public Optimization As Integer = 0     ' 0=None, 1=O1, 2=O2, 3=O3
        Public ErrorChecking As Integer = 0    ' 0=None, 1=-e, 2=-ex, 3=-exx
        Public LangDialect As Integer = 0      ' 0=fb, 1=qb, 2=fblite, 3=deprecated
        Public CodeGen As Integer = 0          ' 0=gas, 1=gcc, 2=llvm
        Public Warnings As Integer = 0         ' 0=None, 1=All, 2=Pedantic
        Public DebugInfo As Boolean = False
        Public Verbose As Boolean = False
        Public ShowCommands As Boolean = False
        Public GenerateMap As Boolean = False
        Public EmitASM As Boolean = False
        Public KeepIntermediate As Boolean = False
        Public TargetArch As Integer = 0       ' 0=32bit, 1=64bit
        Public FPU As Integer = 0              ' 0=x87, 1=sse
        Public StackSize As Integer = 0
        Public OutputFile As String = ""
        Public ExtraCompilerOpts As String = ""
        Public ExtraLinkerOpts As String = ""
        Public IncludePaths As String = ""
        Public LibraryPaths As String = ""
    End Class

    ''' <summary>
    ''' Editor settings
    ''' </summary>
    Public Class EditorSettings
        Public EditorFont As String = "Consolas"
        Public EditorFontSize As Integer = 11
        Public TabWidth As Integer = 4
        Public UseTabs As Boolean = True
        Public ShowLineNumbers As Boolean = True
        Public ShowIndentGuides As Boolean = True
        Public WordWrap As Boolean = False
        Public ShowWhitespace As Boolean = False
        Public AutoIndent As Boolean = True
        Public AutoComplete As Boolean = True
        Public HighlightCurrentLine As Boolean = True
        Public ShowFolding As Boolean = True
        Public DefaultEncoding As FileEncoding = FileEncoding.UTF8
        Public DarkTheme As Boolean = False
    End Class

    ''' <summary>
    ''' Global application settings
    ''' </summary>
    Public Module AppGlobals
        Public Const APP_NAME As String = "FBEditor"
        Public Const APP_VERSION As String = "4.0.0"
        Public Const APP_AUTHOR As String = "Ronen Blumberg"
        Public Const APP_COPYRIGHT As String = "Copyright © 2026 Ronen Blumberg"
        Public Const MAX_RECENT_FILES As Integer = 10

        Public Settings As New EditorSettings()
        Public Build As New BuildSettings()
        Public RecentFiles As New List(Of String)()
        Public AppPath As String = ""
        Public SettingsPath As String = ""
        Public NewFileCounter As Integer = 0

        Private ReadOnly US As UserSettings = UserSettings.Default

        Public Sub InitializeApp()
            AppPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            If Not AppPath.EndsWith("\") Then AppPath &= "\"

            Try
                SettingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), APP_NAME)
                If Not Directory.Exists(SettingsPath) Then Directory.CreateDirectory(SettingsPath)
                If Not SettingsPath.EndsWith("\") Then SettingsPath &= "\"
            Catch
                SettingsPath = AppPath
            End Try

            LoadSettings()
        End Sub

        Public Sub LoadSettings()
            ' Editor settings
            Settings.EditorFont = US.EditorFont
            Settings.EditorFontSize = US.EditorFontSize
            Settings.TabWidth = US.TabWidth
            Settings.UseTabs = US.UseTabs
            Settings.ShowLineNumbers = US.ShowLineNumbers
            Settings.ShowIndentGuides = US.ShowIndentGuides
            Settings.WordWrap = US.WordWrap
            Settings.ShowWhitespace = US.ShowWhitespace
            Settings.AutoIndent = US.AutoIndent
            Settings.AutoComplete = US.AutoComplete
            Settings.HighlightCurrentLine = US.HighlightCurrentLine
            Settings.ShowFolding = US.ShowFolding
            Settings.DefaultEncoding = CType(US.DefaultEncoding, FileEncoding)
            Settings.DarkTheme = US.DarkTheme

            ' Build settings
            Build.FBCPath = US.FBCPath
            Build.FBC32Path = US.FBC32Path
            Build.FBC64Path = US.FBC64Path
            Build.FBDocPath = US.FBDocPath
            Build.APIKeyFilePath = US.APIKeyFilePath
            Build.GDBPath = US.GDBPath
            Build.TargetType = US.TargetType
            Build.Optimization = US.Optimization
            Build.ErrorChecking = US.ErrorChecking
            Build.LangDialect = US.LangDialect
            Build.CodeGen = US.CodeGen
            Build.Warnings = US.Warnings
            Build.DebugInfo = US.DebugInfo
            Build.Verbose = US.Verbose
            Build.ShowCommands = US.ShowCommands
            Build.GenerateMap = US.GenerateMap
            Build.EmitASM = US.EmitASM
            Build.KeepIntermediate = US.KeepIntermediate
            Build.TargetArch = US.TargetArch
            Build.FPU = US.FPU
            Build.StackSize = US.StackSize
            Build.OutputFile = US.OutputFile
            Build.ExtraCompilerOpts = US.ExtraCompilerOpts
            Build.ExtraLinkerOpts = US.ExtraLinkerOpts
            Build.IncludePaths = US.IncludePaths
            Build.LibraryPaths = US.LibraryPaths

            ' Recent files
            RecentFiles.Clear()
            Dim recentStr As String = US.RecentFilesList
            If recentStr IsNot Nothing AndAlso recentStr.Length > 0 Then
                For Each part In recentStr.Split("|"c)
                    If part.Length > 0 Then RecentFiles.Add(part)
                Next
            End If
        End Sub

        Public Sub SaveSettings()
            US.EditorFont = Settings.EditorFont
            US.EditorFontSize = Settings.EditorFontSize
            US.TabWidth = Settings.TabWidth
            US.UseTabs = Settings.UseTabs
            US.ShowLineNumbers = Settings.ShowLineNumbers
            US.ShowIndentGuides = Settings.ShowIndentGuides
            US.WordWrap = Settings.WordWrap
            US.ShowWhitespace = Settings.ShowWhitespace
            US.AutoIndent = Settings.AutoIndent
            US.AutoComplete = Settings.AutoComplete
            US.HighlightCurrentLine = Settings.HighlightCurrentLine
            US.ShowFolding = Settings.ShowFolding
            US.DefaultEncoding = CInt(Settings.DefaultEncoding)
            US.DarkTheme = Settings.DarkTheme

            US.FBCPath = Build.FBCPath
            US.FBC32Path = Build.FBC32Path
            US.FBC64Path = Build.FBC64Path
            US.FBDocPath = Build.FBDocPath
            US.APIKeyFilePath = Build.APIKeyFilePath
            US.GDBPath = Build.GDBPath
            US.TargetType = Build.TargetType
            US.Optimization = Build.Optimization
            US.ErrorChecking = Build.ErrorChecking
            US.LangDialect = Build.LangDialect
            US.CodeGen = Build.CodeGen
            US.Warnings = Build.Warnings
            US.DebugInfo = Build.DebugInfo
            US.Verbose = Build.Verbose
            US.ShowCommands = Build.ShowCommands
            US.GenerateMap = Build.GenerateMap
            US.EmitASM = Build.EmitASM
            US.KeepIntermediate = Build.KeepIntermediate
            US.TargetArch = Build.TargetArch
            US.FPU = Build.FPU
            US.StackSize = Build.StackSize
            US.OutputFile = Build.OutputFile
            US.ExtraCompilerOpts = Build.ExtraCompilerOpts
            US.ExtraLinkerOpts = Build.ExtraLinkerOpts
            US.IncludePaths = Build.IncludePaths
            US.LibraryPaths = Build.LibraryPaths

            If RecentFiles.Count > 0 Then
                US.RecentFilesList = String.Join("|", RecentFiles.ToArray())
            Else
                US.RecentFilesList = ""
            End If

            Try
                US.Save()
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(
                    "Failed to save settings:" & vbCrLf & ex.Message,
                    APP_NAME, System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning)
            End Try
        End Sub

        Public Sub AddRecentFile(filePath As String)
            RecentFiles.RemoveAll(Function(f) f.Equals(filePath, StringComparison.OrdinalIgnoreCase))
            RecentFiles.Insert(0, filePath)
            If RecentFiles.Count > MAX_RECENT_FILES Then
                RecentFiles.RemoveRange(MAX_RECENT_FILES, RecentFiles.Count - MAX_RECENT_FILES)
            End If
        End Sub

        Public Function NewUntitledName() As String
            NewFileCounter += 1
            If NewFileCounter = 1 Then Return "Untitled.bas"
            Return "Untitled" & NewFileCounter.ToString() & ".bas"
        End Function

        Public Function FindFBCPath() As String
            Dim searchPaths() As String = {
                "C:\FreeBASIC\", "C:\fbc\",
                "C:\Program Files\FreeBASIC\",
                "C:\Program Files (x86)\FreeBASIC\",
                "C:\fb_programming\",
                AppPath & "fbc\", AppPath & "FreeBASIC\"
            }
            For Each searchDir In searchPaths
                If File.Exists(searchDir & "fbc32.exe") Then Return searchDir & "fbc32.exe"
                If File.Exists(searchDir & "fbc.exe") Then Return searchDir & "fbc.exe"
            Next
            Dim pathEnv As String = Environment.GetEnvironmentVariable("PATH")
            If pathEnv IsNot Nothing Then
                For Each pathDir In pathEnv.Split(";"c)
                    Dim d = pathDir
                    If Not d.EndsWith("\") Then d &= "\"
                    If File.Exists(d & "fbc32.exe") Then Return d & "fbc32.exe"
                    If File.Exists(d & "fbc.exe") Then Return d & "fbc.exe"
                Next
            End If
            Return ""
        End Function

        Public Function GetEncodingName(enc As FileEncoding) As String
            Select Case enc
                Case FileEncoding.UTF8 : Return "UTF-8"
                Case FileEncoding.UTF8_BOM : Return "UTF-8 BOM"
                Case Else : Return "ANSI"
            End Select
        End Function

        ''' <summary>
        ''' Parse FBC compiler output for errors and warnings.
        ''' Format: filename.bas(line) error num: message
        ''' </summary>
        Public Function ParseCompilerErrors(output As String, Optional baseDir As String = "") As List(Of CompilerError)
            Dim errors As New List(Of CompilerError)()
            If String.IsNullOrEmpty(output) Then Return errors

            Dim pattern = "^(.+?)\((\d+)\)\s+(error|warning)\s+(\d+):\s+(.+)$"
            For Each line In output.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)
                Dim m = Regex.Match(line.Trim(), pattern, RegexOptions.IgnoreCase)
                If m.Success Then
                    Dim filePath = m.Groups(1).Value.Trim()
                    ' Resolve relative paths
                    If Not Path.IsPathRooted(filePath) AndAlso Not String.IsNullOrEmpty(baseDir) Then
                        Dim fullPath = Path.Combine(baseDir, filePath)
                        If File.Exists(fullPath) Then filePath = fullPath
                    End If
                    errors.Add(New CompilerError() With {
                        .FilePath = filePath,
                        .LineNumber = Integer.Parse(m.Groups(2).Value),
                        .ErrorType = m.Groups(3).Value.ToLower(),
                        .ErrorCode = Integer.Parse(m.Groups(4).Value),
                        .Message = m.Groups(5).Value
                    })
                End If
            Next
            Return errors
        End Function

        Public Function DetectFileEncoding(filePath As String) As FileEncoding
            Try
                Dim bytes() As Byte = File.ReadAllBytes(filePath)
                If bytes.Length >= 3 AndAlso bytes(0) = &HEF AndAlso bytes(1) = &HBB AndAlso bytes(2) = &HBF Then
                    Return FileEncoding.UTF8_BOM
                End If
                Dim ci As Integer = 0
                Dim multiByte As Integer = 0
                While ci < bytes.Length
                    If bytes(ci) < &H80 Then
                        ci += 1
                    ElseIf (bytes(ci) And &HE0) = &HC0 AndAlso ci + 1 < bytes.Length AndAlso (bytes(ci + 1) And &HC0) = &H80 Then
                        multiByte += 1 : ci += 2
                    ElseIf (bytes(ci) And &HF0) = &HE0 AndAlso ci + 2 < bytes.Length AndAlso
                           (bytes(ci + 1) And &HC0) = &H80 AndAlso (bytes(ci + 2) And &HC0) = &H80 Then
                        multiByte += 1 : ci += 3
                    ElseIf bytes(ci) >= &H80 Then
                        Return FileEncoding.ANSI
                    Else
                        ci += 1
                    End If
                End While
                If multiByte > 0 Then Return FileEncoding.UTF8
                Return Settings.DefaultEncoding
            Catch
                Return Settings.DefaultEncoding
            End Try
        End Function

        Public Function ReadFileWithEncoding(filePath As String, ByRef detectedEnc As FileEncoding) As String
            detectedEnc = DetectFileEncoding(filePath)
            Select Case detectedEnc
                Case FileEncoding.UTF8, FileEncoding.UTF8_BOM
                    Return File.ReadAllText(filePath, Encoding.UTF8)
                Case Else
                    Return File.ReadAllText(filePath, Encoding.GetEncoding(1252))
            End Select
        End Function

        Public Sub WriteFileWithEncoding(filePath As String, content As String, enc As FileEncoding)
            Select Case enc
                Case FileEncoding.UTF8
                    File.WriteAllText(filePath, content, New UTF8Encoding(False))
                Case FileEncoding.UTF8_BOM
                    File.WriteAllText(filePath, content, New UTF8Encoding(True))
                Case Else
                    File.WriteAllText(filePath, content, Encoding.GetEncoding(1252))
            End Select
        End Sub

        ''' <summary>Safe Process.Start wrapper for compatibility across .NET versions</summary>
        Public Sub SafeProcessStart(url As String)
            Try
                Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
            Catch
                Try : Process.Start(url) : Catch : End Try
            End Try
        End Sub
    End Module
