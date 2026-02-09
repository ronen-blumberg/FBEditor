Imports System.Drawing
Imports System.Windows.Forms


    Public Module ThemeManager

        ' ---- Dark Theme Colors ----
        Public ReadOnly DarkBackground As Color = Color.FromArgb(30, 30, 30)
        Public ReadOnly DarkForeground As Color = Color.FromArgb(212, 212, 212)
        Public ReadOnly DarkPanel As Color = Color.FromArgb(37, 37, 38)
        Public ReadOnly DarkBorder As Color = Color.FromArgb(60, 60, 60)
        Public ReadOnly DarkSelection As Color = Color.FromArgb(38, 79, 120)
        Public ReadOnly DarkCurrentLine As Color = Color.FromArgb(42, 45, 46)
        Public ReadOnly DarkMenuBar As Color = Color.FromArgb(45, 45, 48)
        Public ReadOnly DarkToolbar As Color = Color.FromArgb(45, 45, 48)
        Public ReadOnly DarkStatusBar As Color = Color.FromArgb(0, 122, 204)
        Public ReadOnly DarkTabActive As Color = Color.FromArgb(30, 30, 30)
        Public ReadOnly DarkTabInactive As Color = Color.FromArgb(45, 45, 48)
        Public ReadOnly DarkTreeBack As Color = Color.FromArgb(37, 37, 38)
        Public ReadOnly DarkTreeFore As Color = Color.FromArgb(204, 204, 204)
        Public ReadOnly DarkOutputBack As Color = Color.FromArgb(30, 30, 30)
        Public ReadOnly DarkOutputFore As Color = Color.FromArgb(0, 255, 0)
        Public ReadOnly DarkButtonBack As Color = Color.FromArgb(60, 60, 62)
        Public ReadOnly DarkButtonFore As Color = Color.FromArgb(212, 212, 212)
        Public ReadOnly DarkTextBoxBack As Color = Color.FromArgb(48, 48, 48)
        Public ReadOnly DarkTextBoxFore As Color = Color.FromArgb(212, 212, 212)

        ' ---- Light Theme Colors ----
        Public ReadOnly LightBackground As Color = Color.White
        Public ReadOnly LightForeground As Color = Color.FromArgb(30, 30, 30)
        Public ReadOnly LightPanel As Color = Color.FromArgb(240, 240, 240)
        Public ReadOnly LightBorder As Color = Color.FromArgb(200, 200, 200)
        Public ReadOnly LightSelection As Color = Color.FromArgb(173, 214, 255)
        Public ReadOnly LightCurrentLine As Color = Color.FromArgb(255, 255, 224)
        Public ReadOnly LightMenuBar As Color = Color.FromArgb(240, 240, 240)
        Public ReadOnly LightToolbar As Color = Color.FromArgb(240, 240, 240)
        Public ReadOnly LightStatusBar As Color = Color.FromArgb(0, 122, 204)
        Public ReadOnly LightTabActive As Color = Color.White
        Public ReadOnly LightTabInactive As Color = Color.FromArgb(236, 236, 236)
        Public ReadOnly LightTreeBack As Color = Color.FromArgb(246, 246, 246)
        Public ReadOnly LightTreeFore As Color = Color.FromArgb(30, 30, 30)
        Public ReadOnly LightOutputBack As Color = Color.White
        Public ReadOnly LightOutputFore As Color = Color.FromArgb(30, 30, 30)
        Public ReadOnly LightButtonBack As Color = Color.FromArgb(225, 225, 225)
        Public ReadOnly LightButtonFore As Color = Color.FromArgb(30, 30, 30)
        Public ReadOnly LightTextBoxBack As Color = Color.White
        Public ReadOnly LightTextBoxFore As Color = Color.FromArgb(30, 30, 30)

        ' ---- Scintilla Dark Theme Syntax Colors (BGR format) ----
        Public ReadOnly DarkSciBack As Integer = &H1E1E1E
        Public ReadOnly DarkSciFore As Integer = &HD4D4D4
        Public ReadOnly DarkSciComment As Integer = &H57A64A
        Public ReadOnly DarkSciKeyword As Integer = &HDE7B56
        Public ReadOnly DarkSciKeyword2 As Integer = &HC586C0
        Public ReadOnly DarkSciString As Integer = &H6B99CE
        Public ReadOnly DarkSciNumber As Integer = &H5CB3B0
        Public ReadOnly DarkSciPreproc As Integer = &H7A86C5
        Public ReadOnly DarkSciOperator As Integer = &HD4D4D4
        Public ReadOnly DarkSciCaret As Integer = &HAEAFAD
        Public ReadOnly DarkSciSelBack As Integer = &H643F26
        Public ReadOnly DarkSciCurLine As Integer = &H2A2D2E
        Public ReadOnly DarkSciLineNumFore As Integer = &H858585
        Public ReadOnly DarkSciLineNumBack As Integer = &H1E1E1E
        Public ReadOnly DarkSciFoldMargin As Integer = &H252526
        Public ReadOnly DarkSciIndentGuide As Integer = &H404040

        ' ---- Scintilla Light Theme Syntax Colors (BGR format) ----
        Public ReadOnly LightSciBack As Integer = &HFFFFFF
        Public ReadOnly LightSciFore As Integer = &H1E1E1E
        Public ReadOnly LightSciComment As Integer = &H008000
        Public ReadOnly LightSciKeyword As Integer = &HFF0000
        Public ReadOnly LightSciKeyword2 As Integer = &H800080
        Public ReadOnly LightSciString As Integer = &H0000CE
        Public ReadOnly LightSciNumber As Integer = &H808000
        Public ReadOnly LightSciPreproc As Integer = &H808080
        Public ReadOnly LightSciOperator As Integer = &H1E1E1E
        Public ReadOnly LightSciCaret As Integer = &H000000
        Public ReadOnly LightSciSelBack As Integer = &HFFCC99
        Public ReadOnly LightSciCurLine As Integer = &HE0FFFF
        Public ReadOnly LightSciLineNumFore As Integer = &H808080
        Public ReadOnly LightSciLineNumBack As Integer = &HFFFFFF
        Public ReadOnly LightSciFoldMargin As Integer = &HF0F0F0
        Public ReadOnly LightSciIndentGuide As Integer = &HC0C0C0

        Public Sub ApplyTheme(form As Form, isDark As Boolean)
            If isDark Then
                ApplyDarkTheme(form)
            Else
                ApplyLightTheme(form)
            End If
        End Sub

        Private Sub ApplyDarkTheme(ctrl As Control)
            If TypeOf ctrl Is Form Then
                ctrl.BackColor = DarkPanel
                ctrl.ForeColor = DarkForeground
            ElseIf TypeOf ctrl Is MenuStrip Then
                DirectCast(ctrl, MenuStrip).BackColor = DarkMenuBar
                DirectCast(ctrl, MenuStrip).ForeColor = DarkForeground
                DirectCast(ctrl, MenuStrip).Renderer = New DarkMenuRenderer()
            ElseIf TypeOf ctrl Is ToolStrip Then
                DirectCast(ctrl, ToolStrip).BackColor = DarkToolbar
                DirectCast(ctrl, ToolStrip).ForeColor = DarkForeground
                DirectCast(ctrl, ToolStrip).Renderer = New DarkToolStripRenderer()
            ElseIf TypeOf ctrl Is StatusStrip Then
                ctrl.BackColor = DarkStatusBar
                ctrl.ForeColor = Color.White
            ElseIf TypeOf ctrl Is TreeView Then
                ctrl.BackColor = DarkTreeBack
                ctrl.ForeColor = DarkTreeFore
            ElseIf TypeOf ctrl Is TextBox OrElse TypeOf ctrl Is RichTextBox Then
                ctrl.BackColor = DarkTextBoxBack
                ctrl.ForeColor = DarkTextBoxFore
            ElseIf TypeOf ctrl Is Button Then
                Dim btn = DirectCast(ctrl, Button)
                btn.BackColor = DarkButtonBack
                btn.ForeColor = DarkButtonFore
                btn.FlatStyle = FlatStyle.Flat
                btn.FlatAppearance.BorderColor = DarkBorder
            ElseIf TypeOf ctrl Is ComboBox Then
                ctrl.BackColor = DarkTextBoxBack
                ctrl.ForeColor = DarkTextBoxFore
            ElseIf TypeOf ctrl Is SplitContainer Then
                ctrl.BackColor = DarkBorder
            ElseIf TypeOf ctrl Is TabControl Then
                ' TabControl styling is limited in WinForms
            ElseIf TypeOf ctrl Is Panel Then
                ctrl.BackColor = DarkPanel
                ctrl.ForeColor = DarkForeground
            ElseIf TypeOf ctrl Is Label Then
                ctrl.ForeColor = DarkForeground
            End If

            For Each child As Control In ctrl.Controls
                ApplyDarkTheme(child)
            Next
        End Sub

        Private Sub ApplyLightTheme(ctrl As Control)
            If TypeOf ctrl Is Form Then
                ctrl.BackColor = LightPanel
                ctrl.ForeColor = LightForeground
            ElseIf TypeOf ctrl Is MenuStrip Then
                DirectCast(ctrl, MenuStrip).BackColor = LightMenuBar
                DirectCast(ctrl, MenuStrip).ForeColor = LightForeground
                DirectCast(ctrl, MenuStrip).Renderer = New ToolStripProfessionalRenderer()
            ElseIf TypeOf ctrl Is ToolStrip Then
                DirectCast(ctrl, ToolStrip).BackColor = LightToolbar
                DirectCast(ctrl, ToolStrip).ForeColor = LightForeground
                DirectCast(ctrl, ToolStrip).Renderer = New ToolStripProfessionalRenderer()
            ElseIf TypeOf ctrl Is StatusStrip Then
                ctrl.BackColor = LightStatusBar
                ctrl.ForeColor = Color.White
            ElseIf TypeOf ctrl Is TreeView Then
                ctrl.BackColor = LightTreeBack
                ctrl.ForeColor = LightTreeFore
            ElseIf TypeOf ctrl Is TextBox OrElse TypeOf ctrl Is RichTextBox Then
                ctrl.BackColor = LightTextBoxBack
                ctrl.ForeColor = LightTextBoxFore
            ElseIf TypeOf ctrl Is Button Then
                Dim btn = DirectCast(ctrl, Button)
                btn.BackColor = LightButtonBack
                btn.ForeColor = LightButtonFore
                btn.FlatStyle = FlatStyle.Standard
            ElseIf TypeOf ctrl Is ComboBox Then
                ctrl.BackColor = LightTextBoxBack
                ctrl.ForeColor = LightTextBoxFore
            ElseIf TypeOf ctrl Is SplitContainer Then
                ctrl.BackColor = LightBorder
            ElseIf TypeOf ctrl Is Panel Then
                ctrl.BackColor = LightPanel
                ctrl.ForeColor = LightForeground
            ElseIf TypeOf ctrl Is Label Then
                ctrl.ForeColor = LightForeground
            End If

            For Each child As Control In ctrl.Controls
                ApplyLightTheme(child)
            Next
        End Sub
    End Module

    ' Custom dark menu renderer
    Public Class DarkMenuRenderer
        Inherits ToolStripProfessionalRenderer

        Public Sub New()
            MyBase.New(New DarkColorTable())
        End Sub

        Protected Overrides Sub OnRenderItemText(e As ToolStripItemTextRenderEventArgs)
            e.TextColor = Color.FromArgb(212, 212, 212)
            MyBase.OnRenderItemText(e)
        End Sub
    End Class

    Public Class DarkToolStripRenderer
        Inherits ToolStripProfessionalRenderer

        Public Sub New()
            MyBase.New(New DarkColorTable())
        End Sub
    End Class

    Public Class DarkColorTable
        Inherits ProfessionalColorTable

        Public Overrides ReadOnly Property MenuItemSelected As Color
            Get
                Return Color.FromArgb(62, 62, 64)
            End Get
        End Property
        Public Overrides ReadOnly Property MenuItemBorder As Color
            Get
                Return Color.FromArgb(62, 62, 64)
            End Get
        End Property
        Public Overrides ReadOnly Property MenuStripGradientBegin As Color
            Get
                Return Color.FromArgb(45, 45, 48)
            End Get
        End Property
        Public Overrides ReadOnly Property MenuStripGradientEnd As Color
            Get
                Return Color.FromArgb(45, 45, 48)
            End Get
        End Property
        Public Overrides ReadOnly Property ToolStripDropDownBackground As Color
            Get
                Return Color.FromArgb(45, 45, 48)
            End Get
        End Property
        Public Overrides ReadOnly Property ImageMarginGradientBegin As Color
            Get
                Return Color.FromArgb(45, 45, 48)
            End Get
        End Property
        Public Overrides ReadOnly Property ImageMarginGradientMiddle As Color
            Get
                Return Color.FromArgb(45, 45, 48)
            End Get
        End Property
        Public Overrides ReadOnly Property ImageMarginGradientEnd As Color
            Get
                Return Color.FromArgb(45, 45, 48)
            End Get
        End Property
        Public Overrides ReadOnly Property SeparatorDark As Color
            Get
                Return Color.FromArgb(62, 62, 64)
            End Get
        End Property
        Public Overrides ReadOnly Property SeparatorLight As Color
            Get
                Return Color.FromArgb(62, 62, 64)
            End Get
        End Property
    End Class

