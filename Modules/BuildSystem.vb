Imports System.Diagnostics
Imports System.IO
Imports System.Text


    Public Class BuildResult
        Public Success As Boolean = False
        Public ExitCode As Integer = -1
        Public Output As String = ""
        Public CommandLine As String = ""
        Public Duration As Double = 0
    End Class

    Public Module BuildSystem

        Public Function BuildFile(sourceFile As String, Optional runAfter As Boolean = False,
                                  Optional syntaxOnly As Boolean = False) As BuildResult
            Dim result As New BuildResult()

            ' Validate FBC path
            Dim fbcPath = GetActiveFBCPath()
            If String.IsNullOrEmpty(fbcPath) Then
                result.Output = "ERROR: FreeBASIC compiler path not set." & vbCrLf &
                               "Please configure the compiler path in Build > Build Options..."
                Return result
            End If
            If Not File.Exists(fbcPath) Then
                result.Output = "ERROR: FreeBASIC compiler not found at: " & fbcPath
                Return result
            End If

            ' Build command line
            Dim args = BuildCommandLine(sourceFile, syntaxOnly)
            result.CommandLine = $"""{fbcPath}"" {args}"

            ' Execute compiler
            Dim sw = Stopwatch.StartNew()
            result.Output = ExecuteProcess(fbcPath, args, Path.GetDirectoryName(sourceFile), result.ExitCode)
            sw.Stop()
            result.Duration = sw.Elapsed.TotalSeconds
            result.Success = (result.ExitCode = 0)

            ' Header
            Dim header As New StringBuilder()
            header.AppendLine("Compiler: " & fbcPath)
            header.AppendLine("Command:  " & result.CommandLine)
            header.AppendLine("Source:   " & sourceFile)
            header.AppendLine(New String("-"c, 60))
            result.Output = header.ToString() & result.Output

            ' Summary
            If result.Success Then
                result.Output &= vbCrLf & New String("-"c, 60) & vbCrLf &
                                $"Build successful! ({result.Duration:F2}s)"
                If runAfter AndAlso Not syntaxOnly Then
                    Dim exePath = GetOutputExePath(sourceFile)
                    If File.Exists(exePath) Then
                        result.Output &= vbCrLf & "Running: " & exePath
                        RunExecutable(exePath, Path.GetDirectoryName(sourceFile))
                    Else
                        result.Output &= vbCrLf & "WARNING: Output not found: " & exePath
                    End If
                End If
            Else
                result.Output &= vbCrLf & New String("-"c, 60) & vbCrLf &
                                $"Build FAILED with exit code {result.ExitCode} ({result.Duration:F2}s)"
            End If

            Return result
        End Function

        Public Function GetActiveFBCPath() As String
            ' Use primary path, or fall back to 32/64 based on target
            If Not String.IsNullOrEmpty(Build.FBCPath) Then Return Build.FBCPath
            If Build.TargetArch = 0 AndAlso Not String.IsNullOrEmpty(Build.FBC32Path) Then Return Build.FBC32Path
            If Build.TargetArch = 1 AndAlso Not String.IsNullOrEmpty(Build.FBC64Path) Then Return Build.FBC64Path
            If Not String.IsNullOrEmpty(Build.FBC32Path) Then Return Build.FBC32Path
            If Not String.IsNullOrEmpty(Build.FBC64Path) Then Return Build.FBC64Path
            Return ""
        End Function

        Public Function BuildCommandLine(sourceFile As String, Optional syntaxOnly As Boolean = False) As String
            Dim sb As New StringBuilder()

            If syntaxOnly Then
                sb.Append($" -pp ""{sourceFile}""")
                Return sb.ToString()
            End If

            ' Target type
            Select Case Build.TargetType
                Case 1 : sb.Append(" -s gui")
                Case 2 : sb.Append(" -dll")
                Case 3 : sb.Append(" -lib")
            End Select

            ' Dialect
            Select Case Build.LangDialect
                Case 1 : sb.Append(" -lang qb")
                Case 2 : sb.Append(" -lang fblite")
                Case 3 : sb.Append(" -lang deprecated")
            End Select

            ' Optimization
            If Build.Optimization > 0 Then sb.Append($" -O {Build.Optimization}")

            ' Error checking
            Select Case Build.ErrorChecking
                Case 1 : sb.Append(" -e")
                Case 2 : sb.Append(" -ex")
                Case 3 : sb.Append(" -exx")
            End Select

            ' Code generator
            Select Case Build.CodeGen
                Case 1 : sb.Append(" -gen gcc")
                Case 2 : sb.Append(" -gen llvm")
            End Select

            ' Warnings
            Select Case Build.Warnings
                Case 1 : sb.Append(" -w all")
                Case 2 : sb.Append(" -w pedantic")
            End Select

            ' Architecture
            If Build.TargetArch = 1 Then sb.Append(" -arch x86_64")

            ' FPU
            If Build.FPU = 1 Then sb.Append(" -fpu sse")

            ' Boolean flags
            If Build.DebugInfo Then sb.Append(" -g")
            If Build.Verbose Then sb.Append(" -v")
            If Build.ShowCommands Then sb.Append(" -showincludes")
            If Build.GenerateMap Then sb.Append(" -map")
            If Build.EmitASM Then sb.Append(" -R")
            If Build.KeepIntermediate Then sb.Append(" -C")

            ' Stack size
            If Build.StackSize > 0 Then sb.Append($" -t {Build.StackSize}")

            ' Include paths
            If Not String.IsNullOrEmpty(Build.IncludePaths) Then
                For Each p In Build.IncludePaths.Split(";"c)
                    If p.Trim() <> "" Then sb.Append($" -i ""{p.Trim()}""")
                Next
            End If

            ' Library paths
            If Not String.IsNullOrEmpty(Build.LibraryPaths) Then
                For Each p In Build.LibraryPaths.Split(";"c)
                    If p.Trim() <> "" Then sb.Append($" -p ""{p.Trim()}""")
                Next
            End If

            ' Output
            If Not String.IsNullOrEmpty(Build.OutputFile) Then sb.Append($" -x ""{Build.OutputFile}""")

            ' Extra options
            If Not String.IsNullOrEmpty(Build.ExtraCompilerOpts?.Trim()) Then sb.Append(" " & Build.ExtraCompilerOpts.Trim())
            If Not String.IsNullOrEmpty(Build.ExtraLinkerOpts?.Trim()) Then sb.Append($" -Wl ""{Build.ExtraLinkerOpts.Trim()}""")

            ' Source file last
            sb.Append($" ""{sourceFile}""")
            Return sb.ToString()
        End Function

        Public Function GetOutputExePath(sourceFile As String) As String
            If Not String.IsNullOrEmpty(Build.OutputFile) Then Return Build.OutputFile
            Dim baseName = Path.ChangeExtension(sourceFile, Nothing)
            Select Case Build.TargetType
                Case 2 : Return baseName & ".dll"
                Case 3 : Return baseName & ".a"
                Case Else : Return baseName & ".exe"
            End Select
        End Function

        Public Function ExecuteProcess(exePath As String, arguments As String,
                                       workDir As String, ByRef exitCode As Integer) As String
            Try
                Dim psi As New ProcessStartInfo() With {
                    .FileName = exePath,
                    .Arguments = arguments,
                    .WorkingDirectory = If(String.IsNullOrEmpty(workDir), AppPath, workDir),
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True,
                    .StandardOutputEncoding = Encoding.UTF8,
                    .StandardErrorEncoding = Encoding.UTF8
                }

                Dim output As New StringBuilder()
                Using proc = Process.Start(psi)
                    AddHandler proc.OutputDataReceived, Sub(s, e)
                                                            If e.Data IsNot Nothing Then output.AppendLine(e.Data)
                                                        End Sub
                    AddHandler proc.ErrorDataReceived, Sub(s, e)
                                                           If e.Data IsNot Nothing Then output.AppendLine(e.Data)
                                                       End Sub
                    proc.BeginOutputReadLine()
                    proc.BeginErrorReadLine()

                    If proc.WaitForExit(60000) Then
                        proc.WaitForExit() ' Ensure streams flushed
                        exitCode = proc.ExitCode
                    Else
                        proc.Kill()
                        exitCode = -99
                        output.AppendLine("*** BUILD TIMED OUT after 60 seconds ***")
                    End If
                End Using
                Return output.ToString()
            Catch ex As Exception
                exitCode = -1
                Return "ERROR: " & ex.Message
            End Try
        End Function

        Public Sub RunExecutable(exePath As String, workDir As String)
            Try
                Process.Start(New ProcessStartInfo() With {
                    .FileName = "cmd.exe",
                    .Arguments = $"/k ""{exePath}""",
                    .WorkingDirectory = workDir,
                    .UseShellExecute = True
                })
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show("Error running: " & ex.Message, APP_NAME,
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Sub

        Public Function QuickRun(sourceFile As String) As BuildResult
            Dim origOutput = Build.OutputFile
            Build.OutputFile = Path.Combine(Path.GetTempPath(), "rbfbide_quickrun.exe")
            Dim result = BuildFile(sourceFile, True)
            Build.OutputFile = origOutput
            Return result
        End Function
    End Module

