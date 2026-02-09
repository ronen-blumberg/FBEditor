Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


    Public Class AIChatManager
        Private Shared ReadOnly _httpClient As New HttpClient() With {
            .Timeout = TimeSpan.FromSeconds(120)
        }
        Private _chatHistory As New List(Of ChatMessage)()
        Private _lastResponse As String = ""
        Private _isBusy As Boolean = False

        Public Property IsBusy As Boolean
            Get
                Return _isBusy
            End Get
            Set(value As Boolean)
                _isBusy = value
            End Set
        End Property

        Public Property LastResponse As String
            Get
                Return _lastResponse
            End Get
            Set(value As String)
                _lastResponse = value
            End Set
        End Property

        Public Class ChatMessage
            Private _role As String = ""
            Private _content As String = ""

            Public Property Role As String
                Get
                    Return _role
                End Get
                Set(value As String)
                    _role = value
                End Set
            End Property

            Public Property Content As String
                Get
                    Return _content
                End Get
                Set(value As String)
                    _content = value
                End Set
            End Property
        End Class

        Public Function LoadAPIKey() As String
            Try
                ' First check configured path
                If Not String.IsNullOrEmpty(Build.APIKeyFilePath) AndAlso File.Exists(Build.APIKeyFilePath) Then
                    Return File.ReadAllText(Build.APIKeyFilePath).Trim()
                End If
                ' Check settings directory (%APPDATA%\FBEditor\)
                Dim settingsKeyPath = Path.Combine(SettingsPath, "api_key.txt")
                If File.Exists(settingsKeyPath) Then
                    Return File.ReadAllText(settingsKeyPath).Trim()
                End If
                ' Fall back to app directory
                Dim defaultPath = Path.Combine(AppPath, "api_key.txt")
                If File.Exists(defaultPath) Then
                    Return File.ReadAllText(defaultPath).Trim()
                End If
                Return ""
            Catch
                Return ""
            End Try
        End Function

        Public Async Function SendMessageAsync(userMessage As String, Optional includeCode As Boolean = False,
                                                Optional code As String = "", Optional fileName As String = "") As Task(Of String)
            If IsBusy Then Return "Please wait for the current response to complete."

            Dim apiKey = LoadAPIKey()
            If String.IsNullOrEmpty(apiKey) Then
                Return "Error: No API key found." & vbCrLf &
                       "Configure the API key file path in Build > Build Options," & vbCrLf &
                       "or place 'api_key.txt' in: " & SettingsPath
            End If

            ' Build full message
            Dim fullMsg = userMessage
            If includeCode AndAlso Not String.IsNullOrEmpty(code) Then
                fullMsg &= vbCrLf & vbCrLf &
                           $"Here is my FreeBASIC code ({fileName}):" & vbCrLf &
                           "```freebasic" & vbCrLf & code & vbCrLf & "```"
            End If

            ' Add to history
            _chatHistory.Add(New ChatMessage() With {.Role = "user", .Content = fullMsg})

            ' Keep only last 10 exchanges to avoid token limits
            While _chatHistory.Count > 20
                _chatHistory.RemoveAt(0)
            End While

            IsBusy = True
            Try
                Dim response = Await CallClaudeAPIAsync(apiKey, _chatHistory)
                _lastResponse = response

                ' Add assistant response to history
                _chatHistory.Add(New ChatMessage() With {.Role = "assistant", .Content = response})

                Return response
            Catch ex As Exception
                Return "Error: " & ex.Message
            Finally
                IsBusy = False
            End Try
        End Function

        Private Async Function CallClaudeAPIAsync(apiKey As String, messages As List(Of ChatMessage)) As Task(Of String)
            Dim systemPrompt = "You are a helpful FreeBASIC programming assistant integrated into the FBEditor IDE. " &
                "Help with FreeBASIC code, debugging, syntax, and programming concepts. " &
                "Keep responses concise and focused. When showing code, use FreeBASIC syntax." & vbLf & vbLf &
                "CRITICAL FreeBASIC syntax rules you MUST follow:" & vbLf &
                "- Comments use a single apostrophe: ' This is a comment" & vbLf &
                "- Or REM keyword: REM This is a comment" & vbLf &
                "- NEVER use // for comments. FreeBASIC does NOT support // comments." & vbLf &
                "- NEVER use /* */ block comments. FreeBASIC does NOT support them." & vbLf &
                "- String literals use double quotes: ""hello"" not 'hello'" & vbLf &
                "- Variable declaration: Dim x As Integer" & vbLf &
                "- Always wrap code in ```freebasic code blocks." & vbLf &
                "- If you define Sub Main(), you MUST call Main at the end of the file."

            Dim msgArray = New JArray()
            For Each msg In messages
                msgArray.Add(New JObject From {
                    {"role", msg.Role},
                    {"content", msg.Content}
                })
            Next

            Dim requestBody = New JObject From {
                {"model", "claude-sonnet-4-5-20250929"},
                {"max_tokens", 4096},
                {"system", systemPrompt},
                {"messages", msgArray}
            }

            Dim request = New HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages")
            request.Content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
            request.Headers.Add("x-api-key", apiKey)
            request.Headers.Add("anthropic-version", "2023-06-01")

            Dim response = Await _httpClient.SendAsync(request)
            Dim responseText = Await response.Content.ReadAsStringAsync()

            If Not response.IsSuccessStatusCode Then
                Dim errObj = JObject.Parse(responseText)
                Dim errMsg = errObj?.SelectToken("error.message")?.ToString()
                Return $"API Error ({response.StatusCode}): {If(errMsg, responseText)}"
            End If

            ' Parse response
            Dim json = JObject.Parse(responseText)
            Dim contentArray = json.SelectToken("content")
            If contentArray IsNot Nothing Then
                Dim sb As New StringBuilder()
                For Each item In contentArray
                    If item.Value(Of String)("type") = "text" Then
                        sb.Append(item.Value(Of String)("text"))
                    End If
                Next
                Return sb.ToString()
            End If

            Return "Could not parse response."
        End Function

        Public Sub ClearHistory()
            _chatHistory.Clear()
            _lastResponse = ""
        End Sub

        ''' <summary>
        ''' Extract code from markdown code blocks in AI response
        ''' </summary>
        Public Shared Function ExtractCodeFromResponse(response As String) As String
            If String.IsNullOrEmpty(response) Then Return ""

            ' Try various code block markers
            Dim markers = {"```freebasic" & vbLf, "```freebasic" & vbCrLf,
                          "```basic" & vbLf, "```basic" & vbCrLf,
                          "```fb" & vbLf, "```fb" & vbCrLf,
                          "```" & vbLf, "```" & vbCrLf}

            For Each marker In markers
                Dim startIdx = response.IndexOf(marker, StringComparison.OrdinalIgnoreCase)
                If startIdx >= 0 Then
                    startIdx += marker.Length
                    Dim endIdx = response.IndexOf("```", startIdx)
                    If endIdx > startIdx Then
                        Dim code = response.Substring(startIdx, endIdx - startIdx).TrimEnd()
                        Return FixFreeBASICCode(code)
                    End If
                End If
            Next

            ' Check if it looks like pure code
            Dim trimmed = response.TrimStart().ToUpper()
            If trimmed.StartsWith("'") OrElse trimmed.StartsWith("DIM ") OrElse
               trimmed.StartsWith("PRINT") OrElse trimmed.StartsWith("SUB ") OrElse
               trimmed.StartsWith("FUNCTION ") OrElse trimmed.StartsWith("#INC") OrElse
               trimmed.StartsWith("DECLARE ") Then
                Return FixFreeBASICCode(response.Trim())
            End If

            Return ""
        End Function

        ''' <summary>
        ''' Fix common AI-generated code issues for FreeBASIC
        ''' </summary>
        Private Shared Function FixFreeBASICCode(code As String) As String
            ' Fix smart quotes
            code = code.Replace(ChrW(&H2018), "'").Replace(ChrW(&H2019), "'")
            code = code.Replace(ChrW(&H201C), """").Replace(ChrW(&H201D), """")

            Dim lines = code.Split({vbLf}, StringSplitOptions.None)
            Dim result As New StringBuilder()
            Dim inBlockComment = False

            For i = 0 To lines.Length - 1
                Dim line = lines(i).TrimEnd(vbCr(0))

                If inBlockComment Then
                    Dim endBlock = line.IndexOf("*/")
                    If endBlock >= 0 Then
                        line = "' " & line.Substring(0, endBlock)
                        inBlockComment = False
                    Else
                        line = "' " & line
                    End If
                Else
                    ' Fix // comments (but not URLs like http://)
                    Dim inStr = False
                    For j = 0 To line.Length - 2
                        Dim c = line(j)
                        If c = """"c Then
                            inStr = Not inStr
                        ElseIf Not inStr AndAlso line.Substring(j, 2) = "//" Then
                            If j = 0 OrElse line(j - 1) <> ":"c Then
                                line = line.Substring(0, j).TrimEnd() & " ' " & line.Substring(j + 2).TrimStart()
                                Exit For
                            End If
                        ElseIf Not inStr AndAlso c = "'"c Then
                            Exit For
                        End If
                    Next

                    ' Fix /* block comments
                    If Not inStr Then
                        Dim blockStart = line.IndexOf("/*")
                        If blockStart >= 0 Then
                            Dim blockEnd = line.IndexOf("*/", blockStart + 2)
                            If blockEnd >= 0 Then
                                Dim before = line.Substring(0, blockStart).TrimEnd()
                                Dim comment = line.Substring(blockStart + 2, blockEnd - blockStart - 2).Trim()
                                Dim after = line.Substring(blockEnd + 2).TrimStart()
                                line = (before & " " & after).TrimEnd() & " ' " & comment
                            Else
                                line = line.Substring(0, blockStart) & " ' " & line.Substring(blockStart + 2)
                                inBlockComment = True
                            End If
                        End If
                    End If
                End If

                If i > 0 Then result.AppendLine()
                result.Append(line)
            Next

            Return result.ToString()
        End Function
    End Class

