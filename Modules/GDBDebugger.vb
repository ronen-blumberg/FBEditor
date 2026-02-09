Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading

''' <summary>
''' GDB Debugger integration using GDB/MI (Machine Interface) protocol.
''' Provides breakpoint management, stepping, variable inspection, and call stack navigation.
''' </summary>
Public Class GDBDebugger
    Implements IDisposable

    ' ---- Events (all fire on the UI thread via SynchronizationContext) ----
    Public Event DebugStarted()
    Public Event DebugStopped()
    Public Event DebugPaused(filePath As String, lineNumber As Integer)
    Public Event DebugResumed()
    Public Event DebugOutput(text As String)
    Public Event DebugError(text As String)
    Public Event LocalsUpdated(locals As List(Of VariableInfo))
    Public Event WatchUpdated(watches As List(Of VariableInfo))
    Public Event CallStackUpdated(frames As List(Of StackFrameInfo))
    Public Event BreakpointHit(bpNumber As Integer, filePath As String, lineNumber As Integer)

    ' ---- State ----
    Private _process As Process
    Private _syncCtx As SynchronizationContext
    Private _isRunning As Boolean = False
    Private _isPaused As Boolean = False
    Private _disposed As Boolean = False
    Private _gdbPath As String = ""
    Private _sourceFile As String = ""
    Private _exePath As String = ""
    Private _workDir As String = ""
    Private _miTokenCounter As Integer = 0

    ' ---- Breakpoints ----
    Private _breakpoints As New Dictionary(Of String, List(Of BreakpointInfo))()
    Private _watchExpressions As New List(Of String)()
    Private _pendingWatchTokens As New Dictionary(Of Integer, String)()  ' token -> expression
    Private _watchResults As New List(Of VariableInfo)()
    Private _watchResultCount As Integer = 0
    Private _currentFile As String = ""
    Private _currentLine As Integer = 0

    ' ---- Locals collection (using text-based info locals/args for FreeBASIC compatibility) ----
    Private _localsToken As Integer = -1
    Private _argsToken As Integer = -1
    Private _localsLines As New List(Of String)()
    Private _argsLines As New List(Of String)()
    Private _localsCollected As Boolean = False
    Private _argsCollected As Boolean = False

    ' ---- Data classes ----
    Public Class BreakpointInfo
        Public Number As Integer = 0
        Public FilePath As String = ""
        Public LineNumber As Integer = 0
        Public Enabled As Boolean = True
        Public Condition As String = ""
        Public HitCount As Integer = 0
        Public Pending As Boolean = True
    End Class

    Public Class VariableInfo
        Public Name As String = ""
        Public Value As String = ""
        Public DataType As String = ""
    End Class

    Public Class StackFrameInfo
        Public Level As Integer = 0
        Public FunctionName As String = ""
        Public FilePath As String = ""
        Public LineNumber As Integer = 0
        Public Address As String = ""
    End Class

    ' ---- Properties ----
    Public ReadOnly Property IsRunning As Boolean
        Get
            Return _isRunning
        End Get
    End Property

    Public ReadOnly Property IsPaused As Boolean
        Get
            Return _isPaused
        End Get
    End Property

    Public ReadOnly Property CurrentFile As String
        Get
            Return _currentFile
        End Get
    End Property

    Public ReadOnly Property CurrentLine As Integer
        Get
            Return _currentLine
        End Get
    End Property

    Public ReadOnly Property Breakpoints As Dictionary(Of String, List(Of BreakpointInfo))
        Get
            Return _breakpoints
        End Get
    End Property

    Public ReadOnly Property WatchExpressions As List(Of String)
        Get
            Return _watchExpressions
        End Get
    End Property

    Public Sub New()
        _syncCtx = SynchronizationContext.Current
        If _syncCtx Is Nothing Then _syncCtx = New SynchronizationContext()
    End Sub

    ' ---- GDB Path Detection ----
    Public Shared Function FindGDBPath() As String
        Dim searchPaths() As String = {
            "C:\MinGW\bin\gdb.exe",
            "C:\msys64\mingw64\bin\gdb.exe",
            "C:\msys64\mingw32\bin\gdb.exe",
            "C:\TDM-GCC-64\bin\gdb.exe",
            "C:\TDM-GCC-32\bin\gdb.exe",
            "C:\Program Files\MinGW\bin\gdb.exe",
            "C:\Program Files (x86)\MinGW\bin\gdb.exe"
        }
        For Each p In searchPaths
            If File.Exists(p) Then Return p
        Next

        ' Check near FreeBASIC compiler
        If Not String.IsNullOrEmpty(Build.FBCPath) Then
            Dim fbcDir = Path.GetDirectoryName(Build.FBCPath)
            If fbcDir IsNot Nothing Then
                Dim gdb1 = Path.Combine(fbcDir, "gdb.exe")
                If File.Exists(gdb1) Then Return gdb1
                Dim parentDir = Path.GetDirectoryName(fbcDir)
                If parentDir IsNot Nothing Then
                    Dim gdb2 = Path.Combine(parentDir, "bin", "gdb.exe")
                    If File.Exists(gdb2) Then Return gdb2
                End If
            End If
        End If

        ' Search PATH
        Dim pathEnv = Environment.GetEnvironmentVariable("PATH")
        If pathEnv IsNot Nothing Then
            For Each d In pathEnv.Split(";"c)
                Dim p = Path.Combine(d.Trim(), "gdb.exe")
                If File.Exists(p) Then Return p
            Next
        End If
        Return ""
    End Function

    ' ---- Start / Stop ----
    Public Function StartDebugging(gdbPath As String, exePath As String, sourceFile As String, workDir As String) As Boolean
        If _isRunning Then
            FireOnUI(Sub() RaiseEvent DebugError("Debugger is already running."))
            Return False
        End If
        If String.IsNullOrEmpty(gdbPath) OrElse Not File.Exists(gdbPath) Then
            FireOnUI(Sub() RaiseEvent DebugError("GDB not found: " & If(gdbPath, "(empty)")))
            Return False
        End If
        If Not File.Exists(exePath) Then
            FireOnUI(Sub() RaiseEvent DebugError("Executable not found: " & exePath))
            Return False
        End If

        _gdbPath = gdbPath
        _exePath = exePath
        _sourceFile = sourceFile
        _workDir = If(String.IsNullOrEmpty(workDir), Path.GetDirectoryName(exePath), workDir)

        Try
            Dim psi As New ProcessStartInfo() With {
                .FileName = gdbPath,
                .Arguments = $"--interpreter=mi2 ""{exePath}""",
                .WorkingDirectory = _workDir,
                .UseShellExecute = False,
                .RedirectStandardInput = True,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True,
                .StandardOutputEncoding = Encoding.UTF8,
                .StandardErrorEncoding = Encoding.UTF8
            }

            _process = New Process() With {.StartInfo = psi}
            _process.Start()

            Dim outThread As New Thread(AddressOf ReadGDBOutput)
            outThread.IsBackground = True : outThread.Name = "GDB_Out" : outThread.Start()

            Dim errThread As New Thread(
                Sub()
                    Try
                        While Not _disposed AndAlso _process IsNot Nothing AndAlso Not _process.HasExited
                            Dim line = _process.StandardError.ReadLine()
                            If line IsNot Nothing Then FireOnUI(Sub() RaiseEvent DebugOutput("[GDB ERR] " & line))
                        End While
                    Catch : End Try
                End Sub)
            errThread.IsBackground = True : errThread.Name = "GDB_Err" : errThread.Start()

            _isRunning = True
            _isPaused = False

            SendCmd("-gdb-set new-console on")
            SendCmd("-gdb-set print pretty on")
            SendCmd("-gdb-set pagination off")
            SendAllBreakpoints()

            FireOnUI(Sub()
                         RaiseEvent DebugStarted()
                         RaiseEvent DebugOutput("Debugger started: " & exePath)
                     End Sub)
            Return True
        Catch ex As Exception
            FireOnUI(Sub() RaiseEvent DebugError("Failed to start GDB: " & ex.Message))
            Return False
        End Try
    End Function

    Public Sub StopDebugging()
        If Not _isRunning Then Return
        Try
            If _process IsNot Nothing AndAlso Not _process.HasExited Then
                SendCmdDirect("-gdb-exit")
                If Not _process.WaitForExit(2000) Then _process.Kill()
            End If
        Catch : End Try
        CleanupProcess()
        _isRunning = False : _isPaused = False : _currentFile = "" : _currentLine = 0
        FireOnUI(Sub()
                     RaiseEvent DebugStopped()
                     RaiseEvent DebugOutput("Debugger stopped.")
                 End Sub)
    End Sub

    ' ---- Execution Control ----
    Public Sub Run()
        If Not _isRunning Then Return
        SendCmd("-exec-run") : _isPaused = False
        FireOnUI(Sub() RaiseEvent DebugResumed())
    End Sub

    Public Sub [Continue]()
        If Not _isRunning OrElse Not _isPaused Then Return
        SendCmd("-exec-continue") : _isPaused = False
        FireOnUI(Sub() RaiseEvent DebugResumed())
    End Sub

    Public Sub Pause()
        If Not _isRunning OrElse _isPaused Then Return
        SendCmd("-exec-interrupt")
    End Sub

    Public Sub StepOver()
        If Not _isRunning OrElse Not _isPaused Then Return
        SendCmd("-exec-next") : _isPaused = False
        FireOnUI(Sub() RaiseEvent DebugResumed())
    End Sub

    Public Sub StepInto()
        If Not _isRunning OrElse Not _isPaused Then Return
        SendCmd("-exec-step") : _isPaused = False
        FireOnUI(Sub() RaiseEvent DebugResumed())
    End Sub

    Public Sub StepOut()
        If Not _isRunning OrElse Not _isPaused Then Return
        SendCmd("-exec-finish") : _isPaused = False
        FireOnUI(Sub() RaiseEvent DebugResumed())
    End Sub

    Public Sub RunToCursor(filePath As String, lineNumber As Integer)
        If Not _isRunning OrElse Not _isPaused Then Return
        SendCmd($"-exec-until ""{Norm(filePath)}:{lineNumber}""") : _isPaused = False
        FireOnUI(Sub() RaiseEvent DebugResumed())
    End Sub

    ' ---- Breakpoint Management ----
    Public Function AddBreakpoint(filePath As String, lineNumber As Integer) As BreakpointInfo
        Dim key = Norm(filePath).ToLowerInvariant()
        If Not _breakpoints.ContainsKey(key) Then _breakpoints(key) = New List(Of BreakpointInfo)()
        For Each bp In _breakpoints(key)
            If bp.LineNumber = lineNumber Then Return bp
        Next
        Dim newBP As New BreakpointInfo() With {.FilePath = filePath, .LineNumber = lineNumber, .Enabled = True, .Pending = True}
        _breakpoints(key).Add(newBP)
        If _isRunning Then
            SendCmd($"-break-insert ""{Norm(filePath)}:{lineNumber}""")
            newBP.Pending = False
        End If
        Return newBP
    End Function

    Public Sub RemoveBreakpoint(filePath As String, lineNumber As Integer)
        Dim key = Norm(filePath).ToLowerInvariant()
        If Not _breakpoints.ContainsKey(key) Then Return
        Dim toRemove As BreakpointInfo = Nothing
        For Each bp In _breakpoints(key)
            If bp.LineNumber = lineNumber Then toRemove = bp : Exit For
        Next
        If toRemove IsNot Nothing Then
            If _isRunning AndAlso toRemove.Number > 0 Then SendCmd($"-break-delete {toRemove.Number}")
            _breakpoints(key).Remove(toRemove)
        End If
    End Sub

    Public Function ToggleBreakpoint(filePath As String, lineNumber As Integer) As Boolean
        Dim key = Norm(filePath).ToLowerInvariant()
        If _breakpoints.ContainsKey(key) Then
            For Each bp In _breakpoints(key)
                If bp.LineNumber = lineNumber Then
                    RemoveBreakpoint(filePath, lineNumber) : Return False
                End If
            Next
        End If
        AddBreakpoint(filePath, lineNumber) : Return True
    End Function

    Public Function HasBreakpoint(filePath As String, lineNumber As Integer) As Boolean
        Dim key = Norm(filePath).ToLowerInvariant()
        If Not _breakpoints.ContainsKey(key) Then Return False
        For Each bp In _breakpoints(key)
            If bp.LineNumber = lineNumber Then Return True
        Next
        Return False
    End Function

    Public Function GetBreakpointsForFile(filePath As String) As List(Of BreakpointInfo)
        Dim key = Norm(filePath).ToLowerInvariant()
        If _breakpoints.ContainsKey(key) Then Return _breakpoints(key)
        Return New List(Of BreakpointInfo)()
    End Function

    Public Sub ClearAllBreakpoints()
        If _isRunning Then SendCmd("-break-delete")
        _breakpoints.Clear()
    End Sub

    ' ---- Watch / Inspect ----
    Public Sub AddWatch(expression As String)
        If Not _watchExpressions.Contains(expression) Then
            _watchExpressions.Add(expression)
            If _isRunning AndAlso _isPaused Then RefreshWatches()
        End If
    End Sub

    Public Sub RemoveWatch(expression As String)
        _watchExpressions.Remove(expression)
    End Sub

    Public Sub RefreshWatches()
        If Not _isRunning OrElse Not _isPaused Then Return
        If _watchExpressions.Count = 0 Then
            FireOnUI(Sub() RaiseEvent WatchUpdated(New List(Of VariableInfo)()))
            Return
        End If
        SyncLock _pendingWatchTokens
            _pendingWatchTokens.Clear()
            _watchResults.Clear()
            _watchResultCount = _watchExpressions.Count
        End SyncLock
        For Each expr In _watchExpressions
            Dim token = SendCmd($"-data-evaluate-expression ""{EscGDB(expr)}""")
            If token > 0 Then
                SyncLock _pendingWatchTokens
                    _pendingWatchTokens(token) = expr
                End SyncLock
            End If
        Next
    End Sub

    Public Sub RequestLocals()
        If Not _isRunning OrElse Not _isPaused Then Return
        ' Use text-based "info locals" and "info args" instead of MI -stack-list-locals
        ' because FreeBASIC debug info works much better with text commands
        SyncLock _localsLines
            _localsLines.Clear()
            _argsLines.Clear()
            _localsCollected = False
            _argsCollected = False
        End SyncLock
        _localsToken = SendCmd("-interpreter-exec console ""info locals""")
        _argsToken = SendCmd("-interpreter-exec console ""info args""")
    End Sub

    Public Sub RequestCallStack()
        If Not _isRunning OrElse Not _isPaused Then Return
        SendCmd("-stack-list-frames")
    End Sub

    Public Sub SelectFrame(level As Integer)
        If Not _isRunning OrElse Not _isPaused Then Return
        SendCmd($"-stack-select-frame {level}")
        RequestLocals()
    End Sub

    Public Sub EvaluateExpression(expression As String)
        If Not _isRunning OrElse Not _isPaused Then Return
        SendCmd($"-data-evaluate-expression ""{EscGDB(expression)}""")
    End Sub

    Public Sub SendRawCommand(command As String)
        If Not _isRunning Then Return
        SendCmd(command)
    End Sub

    ' ---- GDB Communication ----
    Private Function SendCmd(command As String) As Integer
        If _process Is Nothing OrElse _process.HasExited Then Return -1
        Try
            _miTokenCounter += 1
            Dim token = _miTokenCounter
            Dim fullCmd = $"{token}{command}"
            _process.StandardInput.WriteLine(fullCmd)
            _process.StandardInput.Flush()
            Return token
        Catch
            Return -1
        End Try
    End Function

    Private Sub SendCmdDirect(command As String)
        If _process Is Nothing OrElse _process.HasExited Then Return
        Try
            _process.StandardInput.WriteLine(command)
            _process.StandardInput.Flush()
        Catch : End Try
    End Sub

    Private Sub SendAllBreakpoints()
        For Each kvp In _breakpoints
            For Each bp In kvp.Value
                If bp.Pending OrElse bp.Number = 0 Then
                    SendCmd($"-break-insert ""{Norm(bp.FilePath)}:{bp.LineNumber}""")
                    bp.Pending = False
                End If
            Next
        Next
    End Sub

    ' ---- GDB Output Reader ----
    Private Sub ReadGDBOutput()
        Try
            While Not _disposed AndAlso _process IsNot Nothing AndAlso Not _process.HasExited
                Dim line = _process.StandardOutput.ReadLine()
                If line Is Nothing Then Exit While
                ProcessLine(line)
            End While
        Catch : End Try
        If _isRunning Then
            _isRunning = False : _isPaused = False
            FireOnUI(Sub() RaiseEvent DebugStopped())
        End If
    End Sub

    Private Sub ProcessLine(line As String)
        If String.IsNullOrEmpty(line) Then Return
        Dim c = line(0)
        Select Case c
            Case "~"c
                Dim text = ExtractQuoted(line, 1)
                ' Check if this console output belongs to a pending locals/args request
                Dim captured = False
                SyncLock _localsLines
                    If _localsToken > 0 AndAlso Not _localsCollected Then
                        _localsLines.Add(text)
                        captured = True
                    ElseIf _argsToken > 0 AndAlso Not _argsCollected AndAlso _localsCollected Then
                        _argsLines.Add(text)
                        captured = True
                    End If
                End SyncLock
                If Not captured Then
                    FireOnUI(Sub() RaiseEvent DebugOutput(text))
                End If
            Case "@"c : FireOnUI(Sub() RaiseEvent DebugOutput("[TARGET] " & ExtractQuoted(line, 1)))
            Case "&"c ' Suppress noisy log
            Case "*"c : ParseAsync(line.Substring(1))
            Case "="c : ParseNotify(line.Substring(1))
            Case "^"c : ParseResult(line, 0)
            Case Else
                ' Extract token number from prefix
                Dim tokenStr = ""
                Dim idx = 0
                While idx < line.Length AndAlso Char.IsDigit(line(idx))
                    tokenStr &= line(idx)
                    idx += 1
                End While
                Dim token = 0
                Integer.TryParse(tokenStr, token)
                If idx < line.Length Then
                    Select Case line(idx)
                        Case "^"c : ParseResult(line.Substring(idx), token)
                        Case "*"c : ParseAsync(line.Substring(idx + 1))
                        Case "="c : ParseNotify(line.Substring(idx + 1))
                    End Select
                End If
        End Select
    End Sub

    Private Sub ParseAsync(data As String)
        If data.StartsWith("stopped") Then
            _isPaused = True
            Dim reason = GetField(data, "reason")
            Dim fp = GetField(data, "fullname") : If String.IsNullOrEmpty(fp) Then fp = GetField(data, "file")
            Dim ln = 0 : Integer.TryParse(GetField(data, "line"), ln)
            _currentFile = If(String.IsNullOrEmpty(fp), _sourceFile, fp)
            _currentLine = ln

            Select Case reason
                Case "breakpoint-hit"
                    Dim bpn = 0 : Integer.TryParse(GetField(data, "bkptno"), bpn)
                    FireOnUI(Sub()
                                 RaiseEvent DebugPaused(_currentFile, _currentLine)
                                 RaiseEvent BreakpointHit(bpn, _currentFile, _currentLine)
                                 RaiseEvent DebugOutput($"Breakpoint {bpn} hit at {Path.GetFileName(_currentFile)}:{_currentLine}")
                             End Sub)
                Case "end-stepping-range", "function-finished"
                    FireOnUI(Sub()
                                 RaiseEvent DebugPaused(_currentFile, _currentLine)
                                 RaiseEvent DebugOutput($"Stopped at {Path.GetFileName(_currentFile)}:{_currentLine}")
                             End Sub)
                Case "signal-received"
                    Dim sn = GetField(data, "signal-name")
                    Dim sm = GetField(data, "signal-meaning")
                    FireOnUI(Sub()
                                 RaiseEvent DebugPaused(_currentFile, _currentLine)
                                 RaiseEvent DebugError($"Signal: {sn} - {sm}")
                             End Sub)
                Case "exited-normally"
                    _isRunning = False : _isPaused = False
                    FireOnUI(Sub()
                                 RaiseEvent DebugOutput("Program exited normally.")
                                 RaiseEvent DebugStopped()
                             End Sub) : Return
                Case "exited"
                    Dim ec = GetField(data, "exit-code")
                    _isRunning = False : _isPaused = False
                    FireOnUI(Sub()
                                 RaiseEvent DebugOutput($"Program exited with code {ec}.")
                                 RaiseEvent DebugStopped()
                             End Sub) : Return
                Case "exited-signalled"
                    Dim sn = GetField(data, "signal-name")
                    _isRunning = False : _isPaused = False
                    FireOnUI(Sub()
                                 RaiseEvent DebugError($"Program terminated by signal: {sn}")
                                 RaiseEvent DebugStopped()
                             End Sub) : Return
                Case Else
                    FireOnUI(Sub()
                                 RaiseEvent DebugPaused(_currentFile, _currentLine)
                                 RaiseEvent DebugOutput($"Stopped ({reason}) at {Path.GetFileName(_currentFile)}:{_currentLine}")
                             End Sub)
            End Select
            If _isPaused Then
                RequestLocals() : RequestCallStack()
                If _watchExpressions.Count > 0 Then RefreshWatches()
            End If
        ElseIf data.StartsWith("running") Then
            _isPaused = False
            FireOnUI(Sub() RaiseEvent DebugResumed())
        End If
    End Sub

    Private Sub ParseNotify(data As String)
        If data.StartsWith("breakpoint-created") OrElse data.StartsWith("breakpoint-modified") Then
            Dim bpn = 0 : Integer.TryParse(GetField(data, "number"), bpn)
            Dim fp = GetField(data, "fullname") : If String.IsNullOrEmpty(fp) Then fp = GetField(data, "file")
            Dim ln = 0 : Integer.TryParse(GetField(data, "line"), ln)
            If bpn > 0 Then UpdateBPNumber(fp, ln, bpn)
        End If
    End Sub

    Private Sub ParseResult(data As String, token As Integer)
        Dim s = data.TrimStart("^"c)
        If s.StartsWith("done") Then
            ' Check if this is a locals/args completion
            If token > 0 AndAlso token = _localsToken Then
                SyncLock _localsLines
                    _localsCollected = True
                End SyncLock
                _localsToken = -1
                ' If args already collected (or no args pending), fire event now
                If _argsToken <= 0 OrElse _argsCollected Then
                    FireLocalsFromTextOutput()
                End If
                Return
            ElseIf token > 0 AndAlso token = _argsToken Then
                SyncLock _localsLines
                    _argsCollected = True
                End SyncLock
                _argsToken = -1
                FireLocalsFromTextOutput()
                Return
            End If

            ' Check if this is a watch result
            Dim isWatch = False
            Dim watchExpr = ""
            SyncLock _pendingWatchTokens
                If token > 0 AndAlso _pendingWatchTokens.ContainsKey(token) Then
                    isWatch = True
                    watchExpr = _pendingWatchTokens(token)
                    _pendingWatchTokens.Remove(token)
                End If
            End SyncLock

            If isWatch Then
                Dim v = ExtractMIValue(s, "value")
                SyncLock _pendingWatchTokens
                    _watchResults.Add(New VariableInfo() With {.Name = watchExpr, .Value = If(v, "(unknown)")})
                    ' Fire event when all watches have been evaluated
                    If _watchResults.Count >= _watchResultCount OrElse _pendingWatchTokens.Count = 0 Then
                        Dim results = New List(Of VariableInfo)(_watchResults)
                        FireOnUI(Sub() RaiseEvent WatchUpdated(results))
                    End If
                End SyncLock
            ElseIf s.Contains("stack=") Then
                ParseStack(s)
            ElseIf s.Contains("value=") Then
                Dim v = ExtractMIValue(s, "value")
                FireOnUI(Sub() RaiseEvent DebugOutput("[EVAL] " & v))
            End If
            If s.Contains("bkpt=") Then
                Dim bpn = 0
                Integer.TryParse(GetField(s, "number"), bpn)
                Dim fp = GetField(s, "fullname")
                If String.IsNullOrEmpty(fp) Then fp = GetField(s, "file")
                Dim ln = 0
                Integer.TryParse(GetField(s, "line"), ln)
                If bpn > 0 Then UpdateBPNumber(fp, ln, bpn)
            End If
        ElseIf s.StartsWith("error") Then
            ' Check if this is a locals/args error (still mark as collected)
            If token > 0 AndAlso token = _localsToken Then
                SyncLock _localsLines
                    _localsCollected = True
                End SyncLock
                _localsToken = -1
                If _argsToken <= 0 OrElse _argsCollected Then FireLocalsFromTextOutput()
                Return
            ElseIf token > 0 AndAlso token = _argsToken Then
                SyncLock _localsLines
                    _argsCollected = True
                End SyncLock
                _argsToken = -1
                FireLocalsFromTextOutput()
                Return
            End If
            ' Check if this is a watch error
            Dim isWatch = False
            Dim watchExpr = ""
            SyncLock _pendingWatchTokens
                If token > 0 AndAlso _pendingWatchTokens.ContainsKey(token) Then
                    isWatch = True
                    watchExpr = _pendingWatchTokens(token)
                    _pendingWatchTokens.Remove(token)
                End If
            End SyncLock
            Dim msg = GetField(s, "msg")
            If isWatch Then
                SyncLock _pendingWatchTokens
                    _watchResults.Add(New VariableInfo() With {.Name = watchExpr, .Value = "<error: " & msg & ">"})
                    If _watchResults.Count >= _watchResultCount OrElse _pendingWatchTokens.Count = 0 Then
                        Dim results = New List(Of VariableInfo)(_watchResults)
                        FireOnUI(Sub() RaiseEvent WatchUpdated(results))
                    End If
                End SyncLock
            Else
                FireOnUI(Sub() RaiseEvent DebugError("[GDB Error] " & msg))
            End If
        ElseIf s.StartsWith("running") Then
            _isPaused = False
        End If
    End Sub

    ''' <summary>Parse collected text output from "info locals" and "info args" and fire LocalsUpdated.</summary>
    Private Sub FireLocalsFromTextOutput()
        Dim locals As New List(Of VariableInfo)()
        Dim allLines As New List(Of String)()
        SyncLock _localsLines
            allLines.AddRange(_argsLines)   ' args first (function parameters)
            allLines.AddRange(_localsLines) ' then locals
        End SyncLock

        ' GDB "info locals"/"info args" output format:
        '   varname = value
        '   arrname = {1, 2, 3}
        '   strvar = 0x12345 "hello"
        ' Multi-line values (structs) may span multiple lines
        Dim currentName = ""
        Dim currentValue = ""

        For Each rawLine In allLines
            ' Each line may end with \n, clean it up
            Dim ln = rawLine.TrimEnd(vbLf(0), vbCr(0), " "c)
            If String.IsNullOrEmpty(ln) Then Continue For
            If ln = "No locals." OrElse ln = "No arguments." Then Continue For

            ' Check if this line starts a new variable (has " = ")
            Dim eqIdx = ln.IndexOf(" = ")
            Dim startsNewVar = False
            If eqIdx > 0 Then
                ' Ensure the part before " = " looks like a variable name (no spaces except at start)
                Dim namePart = ln.Substring(0, eqIdx).Trim()
                If namePart.Length > 0 AndAlso Not namePart.Contains(" ") Then
                    startsNewVar = True
                End If
            End If

            If startsNewVar Then
                ' Save previous variable
                If Not String.IsNullOrEmpty(currentName) Then
                    locals.Add(New VariableInfo() With {.Name = currentName, .Value = currentValue.Trim()})
                End If
                currentName = ln.Substring(0, eqIdx).Trim()
                currentValue = ln.Substring(eqIdx + 3)
            Else
                ' Continuation of previous variable's value (multi-line struct/array)
                If Not String.IsNullOrEmpty(currentName) Then
                    currentValue &= " " & ln.Trim()
                End If
            End If
        Next

        ' Don't forget the last variable
        If Not String.IsNullOrEmpty(currentName) Then
            locals.Add(New VariableInfo() With {.Name = currentName, .Value = currentValue.Trim()})
        End If

        ' Try to extract type info and clean up pointer-prefixed string values
        Dim dummyInt As Integer
        Dim dummyDbl As Double
        For Each v In locals
            ' FreeBASIC strings show as: 0x12345 "actual string"
            Dim strMatch = Regex.Match(v.Value, "^0x[0-9a-fA-F]+ ""(.*)""$")
            If strMatch.Success Then
                v.Value = """" & strMatch.Groups(1).Value & """"
                v.DataType = "STRING"
            ElseIf v.Value.StartsWith("{") Then
                v.DataType = "ARRAY/UDT"
            ElseIf Integer.TryParse(v.Value, dummyInt) Then
                v.DataType = "INTEGER"
            ElseIf Double.TryParse(v.Value, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, dummyDbl) Then
                v.DataType = "DOUBLE"
            End If
        Next

        FireOnUI(Sub() RaiseEvent LocalsUpdated(locals))
    End Sub

    ''' <summary>Read a GDB/MI field value starting at position i. Returns (value, newPosition).</summary>
    Private Shared Function ReadMIFieldValue(data As String, startPos As Integer) As Tuple(Of String, Integer)
        If startPos >= data.Length Then Return Tuple.Create("", startPos)

        Dim i = startPos
        If data(i) = """"c Then
            ' Quoted string - read until matching unescaped quote
            i += 1
            Dim sb As New StringBuilder()
            While i < data.Length
                If data(i) = "\"c AndAlso i + 1 < data.Length Then
                    sb.Append(data(i))
                    sb.Append(data(i + 1))
                    i += 2
                ElseIf data(i) = """"c Then
                    i += 1 ' skip closing quote
                    Return Tuple.Create(sb.ToString(), i)
                Else
                    sb.Append(data(i))
                    i += 1
                End If
            End While
            Return Tuple.Create(sb.ToString(), i)
        ElseIf data(i) = "{"c OrElse data(i) = "["c Then
            ' Nested structure - track brace depth
            Dim closeCh = If(data(i) = "{"c, "}"c, "]"c)
            Dim depth = 1
            Dim sb As New StringBuilder()
            sb.Append(data(i))
            i += 1
            While i < data.Length AndAlso depth > 0
                If data(i) = "\"c AndAlso i + 1 < data.Length Then
                    sb.Append(data(i))
                    sb.Append(data(i + 1))
                    i += 2
                    Continue While
                ElseIf data(i) = """"c Then
                    ' Skip quoted strings inside nested structures
                    sb.Append(data(i))
                    i += 1
                    While i < data.Length
                        If data(i) = "\"c AndAlso i + 1 < data.Length Then
                            sb.Append(data(i))
                            sb.Append(data(i + 1))
                            i += 2
                        ElseIf data(i) = """"c Then
                            sb.Append(data(i))
                            i += 1
                            Exit While
                        Else
                            sb.Append(data(i))
                            i += 1
                        End If
                    End While
                    Continue While
                End If
                If data(i) = "{"c OrElse data(i) = "["c Then depth += 1
                If data(i) = "}"c OrElse data(i) = "]"c Then depth -= 1
                sb.Append(data(i))
                i += 1
            End While
            Return Tuple.Create(sb.ToString(), i)
        Else
            ' Unquoted value (number, etc.)
            Dim sb As New StringBuilder()
            While i < data.Length AndAlso data(i) <> ","c AndAlso data(i) <> "}"c AndAlso data(i) <> "]"c
                sb.Append(data(i))
                i += 1
            End While
            Return Tuple.Create(sb.ToString(), i)
        End If
    End Function

    ''' <summary>Extract a named value from GDB/MI response data.</summary>
    Private Shared Function ExtractMIValue(data As String, fieldName As String) As String
        Dim searchStr = fieldName & "="""
        Dim idx = data.IndexOf(searchStr)
        If idx < 0 Then Return ""
        idx += searchStr.Length
        Dim sb As New StringBuilder()
        While idx < data.Length
            If data(idx) = "\"c AndAlso idx + 1 < data.Length Then
                sb.Append(data(idx))
                sb.Append(data(idx + 1))
                idx += 2
            ElseIf data(idx) = """"c Then
                Exit While
            Else
                sb.Append(data(idx))
                idx += 1
            End If
        End While
        Return sb.ToString()
    End Function

    Private Sub ParseStack(data As String)
        Dim frames As New List(Of StackFrameInfo)()
        For Each m As Match In Regex.Matches(data, "frame=\{([^}]+)\}")
            Dim fd = m.Groups(1).Value
            Dim fr As New StackFrameInfo()
            Integer.TryParse(GetField(fd, "level"), fr.Level)
            fr.Address = GetField(fd, "addr")
            fr.FunctionName = GetField(fd, "func")
            fr.FilePath = GetField(fd, "fullname") : If String.IsNullOrEmpty(fr.FilePath) Then fr.FilePath = GetField(fd, "file")
            Integer.TryParse(GetField(fd, "line"), fr.LineNumber)
            frames.Add(fr)
        Next
        FireOnUI(Sub() RaiseEvent CallStackUpdated(frames))
    End Sub

    ' ---- Helpers ----
    Private Sub UpdateBPNumber(filePath As String, lineNumber As Integer, gdbNumber As Integer)
        If String.IsNullOrEmpty(filePath) Then Return
        Dim key = Norm(filePath).ToLowerInvariant()
        If _breakpoints.ContainsKey(key) Then
            For Each bp In _breakpoints(key)
                If bp.LineNumber = lineNumber OrElse bp.Number = 0 Then bp.Number = gdbNumber : bp.Pending = False : Exit For
            Next
        End If
    End Sub

    Private Shared Function ExtractQuoted(line As String, startIdx As Integer) As String
        If startIdx >= line.Length OrElse line(startIdx) <> """"c Then Return If(startIdx < line.Length, line.Substring(startIdx), "")
        Dim sb As New StringBuilder()
        Dim i = startIdx + 1
        While i < line.Length
            If line(i) = "\"c AndAlso i + 1 < line.Length Then
                Select Case line(i + 1)
                    Case "n"c : sb.Append(vbLf)
                    Case "t"c : sb.Append(vbTab)
                    Case "\"c : sb.Append("\")
                    Case """"c : sb.Append("""")
                    Case Else : sb.Append(line(i))
                End Select
                i += 2
            ElseIf line(i) = """"c Then 
                Exit While
            Else
                sb.Append(line(i))
                i += 1
            End If
        End While
        Return sb.ToString()
    End Function

    Private Shared Function GetField(data As String, name As String) As String
        Dim m = Regex.Match(data, name & "=""([^""]*?)""")
        Return If(m.Success, m.Groups(1).Value, "")
    End Function

    Private Shared Function Norm(path As String) As String
        Return If(String.IsNullOrEmpty(path), "", path.Replace("\", "/"))
    End Function

    Private Shared Function EscGDB(text As String) As String
        Return text.Replace("\", "\\").Replace("""", "\""")
    End Function

    Private Shared Function UnEsc(text As String) As String
        Return text.Replace("\n", vbLf).Replace("\t", vbTab).Replace("\""", """").Replace("\\", "\")
    End Function

    Private Sub FireOnUI(action As Action)
        Try : _syncCtx.Post(Sub(s) action(), Nothing) : Catch : End Try
    End Sub

    Private Sub CleanupProcess()
        Try
            If _process IsNot Nothing Then
                If Not _process.HasExited Then _process.Kill()
                _process.Dispose() : _process = Nothing
            End If
        Catch : End Try
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If _disposed Then Return
        _disposed = True : StopDebugging()
    End Sub
End Class
