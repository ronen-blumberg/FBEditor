Imports System.Text.RegularExpressions


    Public Enum OutlineItemType
        TypeDef = 1
        EnumDef = 2
        ConstDef = 3
        Variable = 4
        ArrayDef = 5
        SubDef = 6
        FuncDef = 7
        PropertyDef = 8
        DeclareDef = 9
        GlobalVar = 10
        GlobalArray = 11
        DynArray = 12
    End Enum

    Public Class OutlineItem
        Public ItemType As OutlineItemType
        Public Name As String = ""
        Public LineNumber As Integer = 0
        Public Signature As String = ""
        Public DataType As String = ""
        Public ReadOnly Property Category As String
            Get
                Select Case ItemType
                    Case OutlineItemType.SubDef, OutlineItemType.FuncDef : Return "Procedures"
                    Case OutlineItemType.TypeDef : Return "Types"
                    Case OutlineItemType.EnumDef : Return "Enums"
                    Case OutlineItemType.ConstDef : Return "Constants"
                    Case OutlineItemType.GlobalVar : Return "Global Variables"
                    Case OutlineItemType.GlobalArray, OutlineItemType.DynArray : Return "Global Arrays"
                    Case OutlineItemType.Variable : Return "Variables"
                    Case OutlineItemType.ArrayDef : Return "Arrays"
                    Case OutlineItemType.PropertyDef : Return "Properties"
                    Case OutlineItemType.DeclareDef : Return "Declares"
                    Case Else : Return "Other"
                End Select
            End Get
        End Property
        Public ReadOnly Property Icon As String
            Get
                Select Case ItemType
                    Case OutlineItemType.SubDef : Return "S"
                    Case OutlineItemType.FuncDef : Return "F"
                    Case OutlineItemType.TypeDef : Return "T"
                    Case OutlineItemType.EnumDef : Return "E"
                    Case OutlineItemType.ConstDef : Return "C"
                    Case OutlineItemType.Variable : Return "V"
                    Case OutlineItemType.ArrayDef : Return "A"
                    Case OutlineItemType.GlobalVar : Return "G"
                    Case OutlineItemType.GlobalArray : Return "GA"
                    Case OutlineItemType.DynArray : Return "DA"
                    Case OutlineItemType.PropertyDef : Return "P"
                    Case OutlineItemType.DeclareDef : Return "D"
                    Case Else : Return "?"
                End Select
            End Get
        End Property
    End Class

    Public Module CodeOutline

        Public Function ParseOutline(code As String) As List(Of OutlineItem)
            Dim items As New List(Of OutlineItem)()
            If String.IsNullOrEmpty(code) Then Return items

            Dim lines = code.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(vbLf(0))
            Dim inType = False, inEnum = False, inMultiLineComment = False
            Dim inProc = False ' Track if inside Sub/Function/Property/Constructor/Destructor

            For i = 0 To lines.Length - 1
                Dim line = lines(i).Trim()
                Dim upper = line.ToUpper()
                If line.Length = 0 Then Continue For
                If line.StartsWith("'") OrElse upper.StartsWith("REM ") Then Continue For

                ' Multi-line comments
                If upper.StartsWith("/'") Then inMultiLineComment = True : Continue For
                If inMultiLineComment Then
                    If line.Contains("'/") Then inMultiLineComment = False
                    Continue For
                End If

                ' Track Type/Enum blocks (don't parse inner members at top level)
                If Regex.IsMatch(upper, "^\s*TYPE\s+\w+") AndAlso Not upper.Contains(" AS ") Then
                    Dim m = Regex.Match(line, "(?i)type\s+(\w+)")
                    If m.Success Then
                        items.Add(New OutlineItem() With {
                            .ItemType = OutlineItemType.TypeDef,
                            .Name = m.Groups(1).Value,
                            .LineNumber = i + 1
                        })
                    End If
                    inType = True
                    Continue For
                End If
                If upper.TrimStart().StartsWith("END TYPE") Then inType = False : Continue For
                If inType Then Continue For

                If Regex.IsMatch(upper, "^\s*ENUM\s+\w+") Then
                    Dim m = Regex.Match(line, "(?i)enum\s+(\w+)")
                    If m.Success Then
                        items.Add(New OutlineItem() With {
                            .ItemType = OutlineItemType.EnumDef,
                            .Name = m.Groups(1).Value,
                            .LineNumber = i + 1
                        })
                    End If
                    inEnum = True
                    Continue For
                End If
                If upper.TrimStart().StartsWith("END ENUM") Then inEnum = False : Continue For
                If inEnum Then Continue For

                ' Track procedure scope (END SUB/FUNCTION/PROPERTY/CONSTRUCTOR/DESTRUCTOR)
                If upper.StartsWith("END SUB") OrElse upper.StartsWith("END FUNCTION") OrElse
                   upper.StartsWith("END PROPERTY") OrElse upper.StartsWith("END CONSTRUCTOR") OrElse
                   upper.StartsWith("END DESTRUCTOR") OrElse upper.StartsWith("END OPERATOR") Then
                    inProc = False
                    Continue For
                End If

                ' SUB
                Dim subMatch = Regex.Match(line, "(?i)^(public\s+|private\s+|static\s+)*sub\s+(\w+)\s*(\(.*)?$")
                If subMatch.Success Then
                    items.Add(New OutlineItem() With {
                        .ItemType = OutlineItemType.SubDef,
                        .Name = subMatch.Groups(2).Value,
                        .LineNumber = i + 1,
                        .Signature = line
                    })
                    inProc = True
                    Continue For
                End If

                ' FUNCTION
                Dim funcMatch = Regex.Match(line, "(?i)^(public\s+|private\s+|static\s+)*function\s+(\w+)\s*(\(.*)?")
                If funcMatch.Success Then
                    Dim dt As String = ""
                    Dim lastParen As Integer = line.LastIndexOf(")"c)
                    If lastParen >= 0 AndAlso lastParen < line.Length - 1 Then
                        Dim afterParen As String = line.Substring(lastParen + 1)
                        Dim asMatch = Regex.Match(afterParen, "(?i)^\s*as\s+(\w+)")
                        If asMatch.Success Then dt = asMatch.Groups(1).Value
                    End If
                    If dt.Length = 0 AndAlso lastParen < 0 Then
                        Dim noParenMatch = Regex.Match(line, "(?i)function\s+\w+\s+as\s+(\w+)")
                        If noParenMatch.Success Then dt = noParenMatch.Groups(1).Value
                    End If
                    items.Add(New OutlineItem() With {
                        .ItemType = OutlineItemType.FuncDef,
                        .Name = funcMatch.Groups(2).Value,
                        .LineNumber = i + 1,
                        .Signature = line,
                        .DataType = dt
                    })
                    inProc = True
                    Continue For
                End If

                ' PROPERTY
                Dim propMatch = Regex.Match(line, "(?i)^(public\s+|private\s+)*property\s+(\w+)")
                If propMatch.Success Then
                    items.Add(New OutlineItem() With {
                        .ItemType = OutlineItemType.PropertyDef,
                        .Name = propMatch.Groups(2).Value,
                        .LineNumber = i + 1
                    })
                    inProc = True
                    Continue For
                End If

                ' CONST
                Dim constMatch = Regex.Match(line, "(?i)^const\s+(\w+)")
                If constMatch.Success Then
                    items.Add(New OutlineItem() With {
                        .ItemType = OutlineItemType.ConstDef,
                        .Name = constMatch.Groups(1).Value,
                        .LineNumber = i + 1
                    })
                    Continue For
                End If

                ' #DEFINE
                Dim defineMatch = Regex.Match(line, "(?i)^#define\s+(\w+)")
                If defineMatch.Success Then
                    items.Add(New OutlineItem() With {
                        .ItemType = OutlineItemType.ConstDef,
                        .Name = defineMatch.Groups(1).Value,
                        .LineNumber = i + 1
                    })
                    Continue For
                End If

                ' DECLARE
                Dim declMatch = Regex.Match(line, "(?i)^declare\s+(sub|function)\s+(\w+)")
                If declMatch.Success Then
                    items.Add(New OutlineItem() With {
                        .ItemType = OutlineItemType.DeclareDef,
                        .Name = declMatch.Groups(2).Value,
                        .LineNumber = i + 1,
                        .Signature = line
                    })
                    Continue For
                End If

                ' DIM / REDIM / COMMON / STATIC variable declarations
                ' Use robust keyword parsing instead of complex regex alternation

                ' Check if line starts with a variable declaration keyword
                Dim dimLine As String = Nothing
                Dim isShared = False
                Dim isRedim = False

                If upper.StartsWith("DIM SHARED ") OrElse upper.StartsWith("DIM" & vbTab & "SHARED") Then
                    dimLine = line.Substring(line.ToUpper().IndexOf("SHARED") + 6).Trim()
                    isShared = True
                ElseIf upper.StartsWith("COMMON SHARED ") Then
                    dimLine = line.Substring(line.ToUpper().IndexOf("SHARED") + 6).Trim()
                    isShared = True
                ElseIf upper.StartsWith("REDIM SHARED ") Then
                    dimLine = line.Substring(line.ToUpper().IndexOf("SHARED") + 6).Trim()
                    isShared = True : isRedim = True
                ElseIf upper.StartsWith("REDIM PRESERVE ") Then
                    dimLine = line.Substring(15).Trim()
                    isRedim = True
                ElseIf upper.StartsWith("REDIM ") Then
                    dimLine = line.Substring(6).Trim()
                    isRedim = True
                ElseIf upper.StartsWith("DIM ") Then
                    dimLine = line.Substring(4).Trim()
                ElseIf upper.StartsWith("COMMON ") Then
                    dimLine = line.Substring(7).Trim()
                ElseIf upper.StartsWith("STATIC ") AndAlso Not upper.StartsWith("STATIC SUB") AndAlso Not upper.StartsWith("STATIC FUNCTION") Then
                    dimLine = line.Substring(7).Trim()
                End If

                If dimLine IsNot Nothing AndAlso dimLine.Length > 0 AndAlso Not inType Then
                    Dim isGlobal = isShared OrElse Not inProc
                    Dim dimUpper = dimLine.ToUpper()

                    ' Pattern A: AS <type> var1, var2, var3 (type-first syntax)
                    If dimUpper.StartsWith("AS ") Then
                        Dim afterAs = dimLine.Substring(3).Trim()
                        Dim spacePos = afterAs.IndexOf(" "c)
                        If spacePos > 0 Then
                            Dim dt = afterAs.Substring(0, spacePos)
                            Dim varList = afterAs.Substring(spacePos + 1).Trim()
                            For Each varPart In varList.Split(","c)
                                Dim vName = varPart.Trim()
                                ' Strip initializer: "x = 5" -> "x"
                                Dim eqPos = vName.IndexOf("="c)
                                If eqPos > 0 Then vName = vName.Substring(0, eqPos).Trim()
                                ' Check for array parens in the name itself: "arr(10)" -> "arr"
                                Dim hasArrayParens = vName.Contains("(")
                                Dim parenPos = vName.IndexOf("("c)
                                If parenPos > 0 Then vName = vName.Substring(0, parenPos).Trim()
                                If vName.Length > 0 AndAlso Regex.IsMatch(vName, "^\w+$") Then
                                    Dim itemType As OutlineItemType
                                    If isRedim Then
                                        itemType = OutlineItemType.DynArray
                                    ElseIf hasArrayParens AndAlso isGlobal Then
                                        itemType = OutlineItemType.GlobalArray
                                    ElseIf hasArrayParens Then
                                        itemType = OutlineItemType.ArrayDef
                                    ElseIf isGlobal Then
                                        itemType = OutlineItemType.GlobalVar
                                    Else
                                        itemType = OutlineItemType.Variable
                                    End If
                                    items.Add(New OutlineItem() With {
                                        .ItemType = itemType,
                                        .Name = vName,
                                        .LineNumber = i + 1,
                                        .DataType = dt
                                    })
                                End If
                            Next
                        End If

                    Else
                        ' Pattern B: var1 AS type [= init], var2 AS type (standard syntax)
                        For Each varPart In dimLine.Split(","c)
                            Dim part = varPart.Trim()
                            If part.Length = 0 Then Continue For
                            ' Strip initializer first: "excuse as integer = int(rnd * 4) + 1" -> "excuse as integer"
                            Dim eqPos = part.IndexOf("="c)
                            If eqPos > 0 Then part = part.Substring(0, eqPos).Trim()
                            ' Extract variable name (first word)
                            Dim nameMatch = Regex.Match(part, "^(\w+)")
                            If nameMatch.Success Then
                                Dim vName = nameMatch.Groups(1).Value
                                If vName.ToUpper() = "AS" OrElse vName.ToUpper() = "SHARED" Then Continue For
                                ' Extract type if present
                                Dim dt = ""
                                Dim asMatch2 = Regex.Match(part, "(?i)\bas\s+(\w+)")
                                If asMatch2.Success Then dt = asMatch2.Groups(1).Value
                                ' Check for array parens in the name part (before AS or at end)
                                Dim hasArrayParens = part.Contains("(") AndAlso part.Contains(")")
                                Dim itemType As OutlineItemType
                                If isRedim Then
                                    itemType = OutlineItemType.DynArray
                                ElseIf hasArrayParens AndAlso isGlobal Then
                                    itemType = OutlineItemType.GlobalArray
                                ElseIf hasArrayParens Then
                                    itemType = OutlineItemType.ArrayDef
                                ElseIf isGlobal Then
                                    itemType = OutlineItemType.GlobalVar
                                Else
                                    itemType = OutlineItemType.Variable
                                End If
                                items.Add(New OutlineItem() With {
                                    .ItemType = itemType,
                                    .Name = vName,
                                    .LineNumber = i + 1,
                                    .DataType = dt
                                })
                            End If
                        Next
                    End If
                End If
            Next

            Return items
        End Function
    End Module

