Imports System.Configuration

''' <summary>
''' .NET built-in user settings - automatically saved to user.config in AppData\Local
''' This is guaranteed to work because .NET handles all the file I/O internally.
''' </summary>
Public NotInheritable Class UserSettings
    Inherits ApplicationSettingsBase

    Private Shared ReadOnly _default As New UserSettings()

    Public Shared ReadOnly Property [Default] As UserSettings
        Get
            Return _default
        End Get
    End Property

    ' ===== Compiler Paths =====

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property FBCPath As String
        Get
            Return CStr(Me("FBCPath"))
        End Get
        Set(value As String)
            Me("FBCPath") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property FBC32Path As String
        Get
            Return CStr(Me("FBC32Path"))
        End Get
        Set(value As String)
            Me("FBC32Path") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property FBC64Path As String
        Get
            Return CStr(Me("FBC64Path"))
        End Get
        Set(value As String)
            Me("FBC64Path") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property FBDocPath As String
        Get
            Return CStr(Me("FBDocPath"))
        End Get
        Set(value As String)
            Me("FBDocPath") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property W9DocPath As String
        Get
            Return CStr(Me("W9DocPath"))
        End Get
        Set(value As String)
            Me("W9DocPath") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property APIKeyFilePath As String
        Get
            Return CStr(Me("APIKeyFilePath"))
        End Get
        Set(value As String)
            Me("APIKeyFilePath") = value
        End Set
    End Property

    ' ===== Build Options =====

    <UserScopedSetting()>
    <DefaultSettingValue("0")>
    Public Property TargetType As Integer
        Get
            Return CInt(Me("TargetType"))
        End Get
        Set(value As Integer)
            Me("TargetType") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("0")>
    Public Property Optimization As Integer
        Get
            Return CInt(Me("Optimization"))
        End Get
        Set(value As Integer)
            Me("Optimization") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("0")>
    Public Property ErrorChecking As Integer
        Get
            Return CInt(Me("ErrorChecking"))
        End Get
        Set(value As Integer)
            Me("ErrorChecking") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("0")>
    Public Property LangDialect As Integer
        Get
            Return CInt(Me("LangDialect"))
        End Get
        Set(value As Integer)
            Me("LangDialect") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("0")>
    Public Property CodeGen As Integer
        Get
            Return CInt(Me("CodeGen"))
        End Get
        Set(value As Integer)
            Me("CodeGen") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("0")>
    Public Property Warnings As Integer
        Get
            Return CInt(Me("Warnings"))
        End Get
        Set(value As Integer)
            Me("Warnings") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("False")>
    Public Property DebugInfo As Boolean
        Get
            Return CBool(Me("DebugInfo"))
        End Get
        Set(value As Boolean)
            Me("DebugInfo") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("False")>
    Public Property Verbose As Boolean
        Get
            Return CBool(Me("Verbose"))
        End Get
        Set(value As Boolean)
            Me("Verbose") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("False")>
    Public Property ShowCommands As Boolean
        Get
            Return CBool(Me("ShowCommands"))
        End Get
        Set(value As Boolean)
            Me("ShowCommands") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("False")>
    Public Property GenerateMap As Boolean
        Get
            Return CBool(Me("GenerateMap"))
        End Get
        Set(value As Boolean)
            Me("GenerateMap") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("False")>
    Public Property EmitASM As Boolean
        Get
            Return CBool(Me("EmitASM"))
        End Get
        Set(value As Boolean)
            Me("EmitASM") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("False")>
    Public Property KeepIntermediate As Boolean
        Get
            Return CBool(Me("KeepIntermediate"))
        End Get
        Set(value As Boolean)
            Me("KeepIntermediate") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("0")>
    Public Property TargetArch As Integer
        Get
            Return CInt(Me("TargetArch"))
        End Get
        Set(value As Integer)
            Me("TargetArch") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("0")>
    Public Property FPU As Integer
        Get
            Return CInt(Me("FPU"))
        End Get
        Set(value As Integer)
            Me("FPU") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("0")>
    Public Property StackSize As Integer
        Get
            Return CInt(Me("StackSize"))
        End Get
        Set(value As Integer)
            Me("StackSize") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property OutputFile As String
        Get
            Return CStr(Me("OutputFile"))
        End Get
        Set(value As String)
            Me("OutputFile") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property ExtraCompilerOpts As String
        Get
            Return CStr(Me("ExtraCompilerOpts"))
        End Get
        Set(value As String)
            Me("ExtraCompilerOpts") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property ExtraLinkerOpts As String
        Get
            Return CStr(Me("ExtraLinkerOpts"))
        End Get
        Set(value As String)
            Me("ExtraLinkerOpts") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property IncludePaths As String
        Get
            Return CStr(Me("IncludePaths"))
        End Get
        Set(value As String)
            Me("IncludePaths") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property LibraryPaths As String
        Get
            Return CStr(Me("LibraryPaths"))
        End Get
        Set(value As String)
            Me("LibraryPaths") = value
        End Set
    End Property

    ' ===== Editor Settings =====

    <UserScopedSetting()>
    <DefaultSettingValue("Consolas")>
    Public Property EditorFont As String
        Get
            Return CStr(Me("EditorFont"))
        End Get
        Set(value As String)
            Me("EditorFont") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("11")>
    Public Property EditorFontSize As Integer
        Get
            Return CInt(Me("EditorFontSize"))
        End Get
        Set(value As Integer)
            Me("EditorFontSize") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("4")>
    Public Property TabWidth As Integer
        Get
            Return CInt(Me("TabWidth"))
        End Get
        Set(value As Integer)
            Me("TabWidth") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("True")>
    Public Property UseTabs As Boolean
        Get
            Return CBool(Me("UseTabs"))
        End Get
        Set(value As Boolean)
            Me("UseTabs") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("True")>
    Public Property ShowLineNumbers As Boolean
        Get
            Return CBool(Me("ShowLineNumbers"))
        End Get
        Set(value As Boolean)
            Me("ShowLineNumbers") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("True")>
    Public Property ShowIndentGuides As Boolean
        Get
            Return CBool(Me("ShowIndentGuides"))
        End Get
        Set(value As Boolean)
            Me("ShowIndentGuides") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("False")>
    Public Property WordWrap As Boolean
        Get
            Return CBool(Me("WordWrap"))
        End Get
        Set(value As Boolean)
            Me("WordWrap") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("False")>
    Public Property ShowWhitespace As Boolean
        Get
            Return CBool(Me("ShowWhitespace"))
        End Get
        Set(value As Boolean)
            Me("ShowWhitespace") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("True")>
    Public Property AutoIndent As Boolean
        Get
            Return CBool(Me("AutoIndent"))
        End Get
        Set(value As Boolean)
            Me("AutoIndent") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("True")>
    Public Property AutoComplete As Boolean
        Get
            Return CBool(Me("AutoComplete"))
        End Get
        Set(value As Boolean)
            Me("AutoComplete") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("True")>
    Public Property HighlightCurrentLine As Boolean
        Get
            Return CBool(Me("HighlightCurrentLine"))
        End Get
        Set(value As Boolean)
            Me("HighlightCurrentLine") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("True")>
    Public Property ShowFolding As Boolean
        Get
            Return CBool(Me("ShowFolding"))
        End Get
        Set(value As Boolean)
            Me("ShowFolding") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("1")>
    Public Property DefaultEncoding As Integer
        Get
            Return CInt(Me("DefaultEncoding"))
        End Get
        Set(value As Integer)
            Me("DefaultEncoding") = value
        End Set
    End Property

    <UserScopedSetting()>
    <DefaultSettingValue("False")>
    Public Property DarkTheme As Boolean
        Get
            Return CBool(Me("DarkTheme"))
        End Get
        Set(value As Boolean)
            Me("DarkTheme") = value
        End Set
    End Property

    ' ===== Recent Files (stored as semicolon-delimited string) =====

    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property RecentFilesList As String
        Get
            Return CStr(Me("RecentFilesList"))
        End Get
        Set(value As String)
            Me("RecentFilesList") = value
        End Set
    End Property


    <UserScopedSetting()>
    <DefaultSettingValue("")>
    Public Property GDBPath As String
        Get
            Return CStr(Me("GDBPath"))
        End Get
        Set(value As String)
            Me("GDBPath") = value
        End Set
    End Property
End Class
