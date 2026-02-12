Imports ScintillaNET
Imports System.Drawing


    ''' <summary>
    ''' Manual code folding engine for FreeBASIC.
    ''' Scans source for block keywords and sets fold levels on each Scintilla line.
    ''' Works reliably even with unformatted or inconsistently indented code.
    ''' Handles ENDIF (no space), single-line IF, string-aware comment stripping,
    ''' CONSTRUCTOR/DESTRUCTOR/OPERATOR, TYPE vs type-alias, #IF/#ENDIF, and /' '/ comments.
    ''' </summary>
    Public Module FoldingManager

        Private Const FOLD_BASE As Integer = &H400    ' SC_FOLDLEVELBASE (1024)
        Private Const FOLD_HEADER As Integer = &H2000  ' SC_FOLDLEVELHEADERFLAG

        ''' <summary>
        ''' Recompute fold levels for every line in the given Scintilla control.
        ''' Call this after text changes (debounced).
        ''' </summary>
        Public Sub UpdateFoldLevels(sci As Scintilla)
            If sci Is Nothing OrElse sci.Lines.Count = 0 Then Return

            Dim lineCount = sci.Lines.Count
            Dim currentLevel As Integer = 0
            Dim inMultiLineComment As Boolean = False

            For i = 0 To lineCount - 1
                Dim rawLine = sci.Lines(i).Text
                Dim trimmed = rawLine.TrimStart()
                Dim upper = trimmed.ToUpper()
                Dim lineLevel = currentLevel

                ' ---- Multi-line comment: /' ... '/ ----
                If inMultiLineComment Then
                    If trimmed.Contains("'/") Then
                        inMultiLineComment = False
                        currentLevel -= 1
                        If currentLevel < 0 Then currentLevel = 0
                    End If
                    sci.Lines(i).FoldLevel = FOLD_BASE + lineLevel
                    Continue For
                End If

                ' Check for multi-line comment open
                If upper.TrimStart().StartsWith("/'") Then
                    ' Only fold if close is NOT on same line
                    If Not trimmed.Substring(2).Contains("'/") Then
                        sci.Lines(i).FoldLevel = FOLD_BASE + lineLevel Or FOLD_HEADER
                        inMultiLineComment = True
                        currentLevel += 1
                        Continue For
                    Else
                        ' Single-line block comment /' ... '/ — no fold
                        sci.Lines(i).FoldLevel = FOLD_BASE + lineLevel
                        Continue For
                    End If
                End If

                ' Strip trailing comment (string-aware) and get code portion
                Dim codePart = StripTrailingComment(trimmed)
                Dim codeUpper = codePart.Trim().ToUpper()

                ' Blank or comment-only lines
                If codeUpper.Length = 0 OrElse codeUpper.StartsWith("'") OrElse codeUpper.StartsWith("REM ") Then
                    sci.Lines(i).FoldLevel = FOLD_BASE + lineLevel
                    Continue For
                End If

                ' ---- Calculate fold delta ----
                Dim delta = GetBlockDelta(codeUpper)

                If delta > 0 Then
                    ' Opens a block — mark as fold header
                    sci.Lines(i).FoldLevel = FOLD_BASE + lineLevel Or FOLD_HEADER
                    currentLevel += delta
                ElseIf delta < 0 Then
                    ' Closes a block — show at decreased level
                    currentLevel += delta  ' delta is negative
                    If currentLevel < 0 Then currentLevel = 0
                    sci.Lines(i).FoldLevel = FOLD_BASE + currentLevel
                Else
                    ' Check for mid-block keywords (ELSE, ELSEIF, CASE)
                    If IsMidBlockKeyword(codeUpper) Then
                        Dim parentLevel = If(currentLevel > 0, currentLevel - 1, 0)
                        sci.Lines(i).FoldLevel = FOLD_BASE + parentLevel Or FOLD_HEADER
                    Else
                        sci.Lines(i).FoldLevel = FOLD_BASE + lineLevel
                    End If
                End If
            Next
        End Sub

        ''' <summary>
        ''' Returns +N for block openers, -N for block closers, 0 for regular/mid-block lines.
        ''' </summary>
        Private Function GetBlockDelta(codeUpper As String) As Integer
            Dim delta = 0

            ' ============ CLOSERS (check first) ============

            ' Two-word END closers: END SUB, END FUNCTION, END TYPE, etc.
            If StartsWithKeyword(codeUpper, "END SUB") OrElse
               StartsWithKeyword(codeUpper, "END FUNCTION") OrElse
               StartsWithKeyword(codeUpper, "END TYPE") OrElse
               StartsWithKeyword(codeUpper, "END ENUM") OrElse
               StartsWithKeyword(codeUpper, "END IF") OrElse
               StartsWithKeyword(codeUpper, "END SELECT") OrElse
               StartsWithKeyword(codeUpper, "END WITH") OrElse
               StartsWithKeyword(codeUpper, "END SCOPE") OrElse
               StartsWithKeyword(codeUpper, "END CONSTRUCTOR") OrElse
               StartsWithKeyword(codeUpper, "END DESTRUCTOR") OrElse
               StartsWithKeyword(codeUpper, "END PROPERTY") OrElse
               StartsWithKeyword(codeUpper, "END OPERATOR") OrElse
               StartsWithKeyword(codeUpper, "END UNION") OrElse
               StartsWithKeyword(codeUpper, "END NAMESPACE") OrElse
               StartsWithKeyword(codeUpper, "END EXTERN") Then
                delta -= 1
            End If

            ' Single-word closers
            If StartsWithKeyword(codeUpper, "NEXT") Then delta -= 1
            If StartsWithKeyword(codeUpper, "LOOP") Then delta -= 1
            If StartsWithKeyword(codeUpper, "WEND") Then delta -= 1
            If StartsWithKeyword(codeUpper, "ENDIF") Then delta -= 1  ' FreeBASIC allows ENDIF (no space)

            ' Preprocessor closers
            If codeUpper.StartsWith("#ENDIF") OrElse codeUpper.StartsWith("# ENDIF") Then delta -= 1
            If codeUpper.StartsWith("#ENDMACRO") Then delta -= 1

            ' ============ OPENERS ============

            ' SUB/FUNCTION/CONSTRUCTOR/DESTRUCTOR/OPERATOR/PROPERTY (not DECLARE, not END)
            If Not StartsWithKeyword(codeUpper, "DECLARE") AndAlso Not StartsWithKeyword(codeUpper, "END") Then
                If StartsWithKeyword(codeUpper, "SUB") OrElse
                   StartsWithModifier(codeUpper, "SUB") Then
                    delta += 1
                ElseIf StartsWithKeyword(codeUpper, "FUNCTION") OrElse
                       StartsWithModifier(codeUpper, "FUNCTION") Then
                    delta += 1
                ElseIf StartsWithKeyword(codeUpper, "CONSTRUCTOR") Then
                    delta += 1
                ElseIf StartsWithKeyword(codeUpper, "DESTRUCTOR") Then
                    delta += 1
                ElseIf StartsWithKeyword(codeUpper, "OPERATOR") Then
                    delta += 1
                ElseIf StartsWithKeyword(codeUpper, "PROPERTY") OrElse
                       StartsWithModifier(codeUpper, "PROPERTY") Then
                    delta += 1
                End If
            End If

            ' TYPE name (block definition, NOT "TYPE AS ..." alias)
            If StartsWithKeyword(codeUpper, "TYPE") Then
                Dim afterType = codeUpper.Substring(4).TrimStart()
                If afterType.Length > 0 AndAlso Not afterType.StartsWith("AS ") Then
                    ' Check it's not "TYPE name AS othertype" (single-line alias)
                    Dim parts = afterType.Split(New Char() {" "c, CChar(vbTab)}, StringSplitOptions.RemoveEmptyEntries)
                    If parts.Length < 2 OrElse parts(1) <> "AS" Then
                        delta += 1
                    End If
                End If
            End If

            ' ENUM
            If StartsWithKeyword(codeUpper, "ENUM") Then delta += 1

            ' UNION
            If StartsWithKeyword(codeUpper, "UNION") Then delta += 1

            ' NAMESPACE
            If StartsWithKeyword(codeUpper, "NAMESPACE") Then delta += 1

            ' EXTERN "C" / EXTERN "Windows" etc.
            If StartsWithKeyword(codeUpper, "EXTERN") AndAlso codeUpper.Contains("""") Then delta += 1

            ' SELECT CASE
            If StartsWithKeyword(codeUpper, "SELECT") Then delta += 1

            ' IF (multi-line only: "IF ... THEN" with nothing after THEN)
            If StartsWithKeyword(codeUpper, "IF") Then
                If IsMultiLineIf(codeUpper) Then
                    delta += 1
                End If
            End If

            ' FOR (but not single-line FOR...NEXT)
            If StartsWithKeyword(codeUpper, "FOR") Then
                If Not codeUpper.Contains(": NEXT") AndAlso Not codeUpper.Contains(":NEXT") Then
                    delta += 1
                End If
            End If

            ' DO
            If StartsWithKeyword(codeUpper, "DO") Then delta += 1

            ' WHILE
            If StartsWithKeyword(codeUpper, "WHILE") Then delta += 1

            ' WITH (not END WITH — already checked above)
            If StartsWithKeyword(codeUpper, "WITH") AndAlso Not StartsWithKeyword(codeUpper, "END WITH") Then
                delta += 1
            End If

            ' SCOPE (not END SCOPE)
            If StartsWithKeyword(codeUpper, "SCOPE") AndAlso Not StartsWithKeyword(codeUpper, "END SCOPE") Then
                delta += 1
            End If

            ' Preprocessor openers
            If codeUpper.StartsWith("#IF") OrElse codeUpper.StartsWith("# IF") OrElse
               codeUpper.StartsWith("#IFDEF") OrElse codeUpper.StartsWith("#IFNDEF") OrElse
               codeUpper.StartsWith("# IFDEF") OrElse codeUpper.StartsWith("# IFNDEF") Then
                delta += 1
            End If
            If codeUpper.StartsWith("#MACRO") Then delta += 1

            Return delta
        End Function

        ' ---- Helpers ----

        ''' <summary>
        ''' Determine if an IF statement is multi-line.
        ''' Multi-line: "IF condition THEN" with nothing meaningful after THEN.
        ''' Single-line: "IF condition THEN statement" or "IF cond GOTO label".
        ''' </summary>
        Private Function IsMultiLineIf(codeUpper As String) As Boolean
            ' Find THEN keyword (string-aware)
            Dim thenPos = FindKeywordThen(codeUpper)
            If thenPos < 0 Then Return False  ' No THEN = single-line (GOTO, etc.)

            ' Check what comes after THEN
            Dim afterThen = codeUpper.Substring(thenPos + 4).Trim()
            If afterThen.Length = 0 Then Return True
            If afterThen.StartsWith("'") Then Return True
            If afterThen.StartsWith("REM ") Then Return True

            ' Code after THEN → single-line IF
            Return False
        End Function

        ''' <summary>
        ''' Find THEN keyword position, skipping occurrences inside string literals.
        ''' </summary>
        Private Function FindKeywordThen(upper As String) As Integer
            Dim inStr = False
            For i = 0 To upper.Length - 4
                If upper(i) = """"c Then
                    inStr = Not inStr
                ElseIf Not inStr AndAlso i + 4 <= upper.Length Then
                    If upper.Substring(i, 4) = "THEN" Then
                        Dim beforeOk = (i = 0) OrElse (Not Char.IsLetterOrDigit(upper(i - 1)) AndAlso upper(i - 1) <> "_"c)
                        Dim afterOk = (i + 4 >= upper.Length) OrElse (Not Char.IsLetterOrDigit(upper(i + 4)) AndAlso upper(i + 4) <> "_"c)
                        If beforeOk AndAlso afterOk Then Return i
                    End If
                End If
            Next
            Return -1
        End Function

        ''' <summary>
        ''' Check if line is a mid-block keyword (ELSE, ELSEIF, CASE, #ELSE, #ELSEIF).
        ''' These create sub-sections within a block without changing net fold level.
        ''' </summary>
        Private Function IsMidBlockKeyword(codeUpper As String) As Boolean
            If StartsWithKeyword(codeUpper, "ELSE") AndAlso Not StartsWithKeyword(codeUpper, "ELSEIF") Then Return True
            If StartsWithKeyword(codeUpper, "ELSEIF") Then Return True
            If StartsWithKeyword(codeUpper, "CASE") Then Return True
            If codeUpper.StartsWith("#ELSE") AndAlso Not codeUpper.StartsWith("#ELSEIF") Then Return True
            If codeUpper.StartsWith("#ELSEIF") Then Return True
            Return False
        End Function

        ''' <summary>
        ''' Check if string starts with keyword as a whole word (not a prefix of another word).
        ''' "ENDIF" matches "ENDIF" but not "ENDIFX". Also accepts ( : ' as delimiters.
        ''' </summary>
        Private Function StartsWithKeyword(codeUpper As String, keyword As String) As Boolean
            If Not codeUpper.StartsWith(keyword) Then Return False
            If codeUpper.Length = keyword.Length Then Return True
            Dim c = codeUpper(keyword.Length)
            Return c = " "c OrElse c = CChar(vbTab) OrElse c = "("c OrElse c = ":"c OrElse c = "'"c
        End Function

        ''' <summary>
        ''' Check if line starts with PUBLIC/PRIVATE/STATIC/ABSTRACT/VIRTUAL followed by keyword.
        ''' Example: "PUBLIC SUB ...", "PRIVATE FUNCTION ...", "STATIC SUB ..."
        ''' </summary>
        Private Function StartsWithModifier(codeUpper As String, keyword As String) As Boolean
            For Each modifier In {"PUBLIC", "PRIVATE", "STATIC", "ABSTRACT", "VIRTUAL"}
                If StartsWithKeyword(codeUpper, modifier) Then
                    Dim rest = codeUpper.Substring(modifier.Length).TrimStart()
                    If StartsWithKeyword(rest, keyword) Then Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Strip trailing comment from a line, aware of string literals.
        ''' "PRINT "it's" ' comment" → "PRINT "it's" "
        ''' </summary>
        Private Function StripTrailingComment(line As String) As String
            Dim inString = False
            For j = 0 To line.Length - 1
                Dim c = line(j)
                If c = """"c Then
                    inString = Not inString
                ElseIf Not inString AndAlso c = "'"c Then
                    Return line.Substring(0, j)
                End If
            Next
            Return line
        End Function

        ' ---- Public commands for menu/toolbar actions ----

        ''' <summary>Toggle fold at the current line.</summary>
        Public Sub ToggleFold(sci As Scintilla)
            Dim line = sci.CurrentLine
            ' If current line is not a header, find the parent fold header
            If (sci.Lines(line).FoldLevel And FOLD_HEADER) = 0 Then
                Dim lvl = sci.Lines(line).FoldLevel And &HFFF
                For j = line - 1 To 0 Step -1
                    If (sci.Lines(j).FoldLevel And FOLD_HEADER) <> 0 Then
                        Dim headerLvl = sci.Lines(j).FoldLevel And &HFFF
                        If headerLvl < lvl Then
                            line = j
                            Exit For
                        End If
                    End If
                Next
            End If
            If (sci.Lines(line).FoldLevel And FOLD_HEADER) <> 0 Then
                sci.Lines(line).ToggleFold()
            End If
        End Sub

        ''' <summary>Fold all blocks.</summary>
        Public Sub FoldAll(sci As Scintilla)
            sci.FoldAll(FoldAction.Contract)
        End Sub

        ''' <summary>Unfold all blocks.</summary>
        Public Sub UnfoldAll(sci As Scintilla)
            sci.FoldAll(FoldAction.Expand)
        End Sub

        ''' <summary>Fold all blocks at or beyond a given depth.</summary>
        Public Sub FoldToLevel(sci As Scintilla, maxLevel As Integer)
            UnfoldAll(sci)
            For i = 0 To sci.Lines.Count - 1
                Dim lvl = (sci.Lines(i).FoldLevel And &HFFF) - FOLD_BASE
                If (sci.Lines(i).FoldLevel And FOLD_HEADER) <> 0 AndAlso lvl >= maxLevel Then
                    If sci.Lines(i).Expanded Then
                        sci.Lines(i).ToggleFold()
                    End If
                End If
            Next
        End Sub

        ''' <summary>Set up the fold margin (margin index 3) on a Scintilla control.</summary>
        Public Sub SetupFoldMargin(sci As Scintilla, isDark As Boolean)
            ' Margin 3 = fold margin
            Dim foldMargin = sci.Margins(3)
            foldMargin.Type = MarginType.Symbol
            foldMargin.Mask = Marker.MaskFolders
            foldMargin.Width = 20
            foldMargin.Sensitive = True

            Dim foldBack = If(isDark, Color.FromArgb(30, 30, 30), Color.White)
            Dim foldFore = If(isDark, Color.FromArgb(120, 120, 120), Color.FromArgb(140, 140, 140))

            ' Box-style fold markers (classic +/- look)
            sci.Markers(Marker.Folder).Symbol = MarkerSymbol.BoxPlus
            sci.Markers(Marker.FolderOpen).Symbol = MarkerSymbol.BoxMinus
            sci.Markers(Marker.FolderEnd).Symbol = MarkerSymbol.BoxPlusConnected
            sci.Markers(Marker.FolderOpenMid).Symbol = MarkerSymbol.BoxMinusConnected
            sci.Markers(Marker.FolderMidTail).Symbol = MarkerSymbol.TCorner
            sci.Markers(Marker.FolderSub).Symbol = MarkerSymbol.VLine
            sci.Markers(Marker.FolderTail).Symbol = MarkerSymbol.LCorner

            ' Set colors for all fold markers
            For Each mk In {Marker.Folder, Marker.FolderOpen, Marker.FolderEnd,
                            Marker.FolderOpenMid, Marker.FolderMidTail,
                            Marker.FolderSub, Marker.FolderTail}
                sci.Markers(mk).SetBackColor(foldFore)
                sci.Markers(mk).SetForeColor(foldBack)
            Next

            ' Fold margin background
            sci.SetFoldMarginColor(True, foldBack)
            sci.SetFoldMarginHighlightColor(True, foldBack)

            ' Enable automatic fold actions on click and show
            sci.AutomaticFold = AutomaticFold.Show Or AutomaticFold.Click Or AutomaticFold.Change

            ' CRITICAL: Disable built-in lexer fold — we set levels manually
            sci.SetProperty("fold", "0")

            ' Draw line below collapsed block (FoldFlags = 16 = SC_FOLDFLAG_LINEAFTER_CONTRACTED)
            sci.DirectMessage(2233, New IntPtr(16), IntPtr.Zero)  ' SCI_SETFOLDFLAGS
        End Sub

        ''' <summary>Hide the fold margin.</summary>
        Public Sub HideFoldMargin(sci As Scintilla)
            sci.Margins(3).Width = 0
        End Sub

    End Module
