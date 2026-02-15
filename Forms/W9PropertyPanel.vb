Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Property panel — displays and edits properties of the currently selected gadget.
''' Styled similar to Visual Studio's Properties window with categorized fields.
''' Shows form properties when no gadget is selected.
''' </summary>
Public Class W9PropertyPanel
    Inherits Panel

    Private _propGrid As PropertyGrid
    Private _gadgetCombo As ComboBox
    Private _headerLabel As Label
    Private _formDesign As W9FormDesign
    Private _selectedGadget As W9GadgetInstance = Nothing
    Private _isDarkTheme As Boolean = False

    ''' <summary>Fired when a property value changes.</summary>
    Public Event PropertyChanged(gadget As W9GadgetInstance, propName As String)
    Public Event FormPropertyChanged(propName As String)

    ' =========================================================================
    ' Constructor
    ' =========================================================================
    Public Sub New()
        Me.Dock = DockStyle.Fill
        Me.Padding = New Padding(0)

        ' Header
        _headerLabel = New Label() With {
            .Text = "Properties",
            .Dock = DockStyle.Top,
            .Height = 24,
            .TextAlign = ContentAlignment.MiddleLeft,
            .Padding = New Padding(6, 0, 0, 0),
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .BackColor = Color.FromArgb(0, 120, 215),
            .ForeColor = Color.White
        }
        Me.Controls.Add(_headerLabel)

        ' Gadget selector combo
        _gadgetCombo = New ComboBox() With {
            .Dock = DockStyle.Top,
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Font = New Font("Segoe UI", 9)
        }
        AddHandler _gadgetCombo.SelectedIndexChanged, AddressOf OnGadgetComboChanged
        Me.Controls.Add(_gadgetCombo)

        ' PropertyGrid
        _propGrid = New PropertyGrid() With {
            .Dock = DockStyle.Fill,
            .PropertySort = PropertySort.Categorized,
            .ToolbarVisible = False,
            .HelpVisible = True,
            .Font = New Font("Segoe UI", 9)
        }
        AddHandler _propGrid.PropertyValueChanged, AddressOf OnPropertyValueChanged
        Me.Controls.Add(_propGrid)

        ' Ordering (bottom to top for Dock)
        _propGrid.BringToFront()
        _gadgetCombo.BringToFront()
        _headerLabel.SendToBack()
    End Sub

    ' =========================================================================
    ' Public interface
    ' =========================================================================

    Public Sub SetFormDesign(design As W9FormDesign)
        _formDesign = design
        RefreshGadgetCombo()
    End Sub

    Public Sub ShowGadgetProperties(gadget As W9GadgetInstance)
        _selectedGadget = gadget
        If gadget IsNot Nothing Then
            Dim wrapper As New GadgetPropertyWrapper(gadget)
            _propGrid.SelectedObject = wrapper
            For i = 0 To _gadgetCombo.Items.Count - 1
                Dim item = TryCast(_gadgetCombo.Items(i), GadgetComboItem)
                If item IsNot Nothing AndAlso item.Gadget Is gadget Then
                    _gadgetCombo.SelectedIndex = i
                    Exit For
                End If
            Next
        Else
            ShowFormProperties()
        End If
    End Sub

    Public Sub ShowFormProperties()
        _selectedGadget = Nothing
        If _formDesign IsNot Nothing Then
            Dim wrapper As New FormPropertyWrapper(_formDesign)
            _propGrid.SelectedObject = wrapper
            If _gadgetCombo.Items.Count > 0 Then
                _gadgetCombo.SelectedIndex = 0
            End If
        End If
    End Sub

    Public Sub RefreshGadgetCombo()
        _gadgetCombo.Items.Clear()
        _gadgetCombo.Items.Add(New GadgetComboItem() With {
            .DisplayText = "(Form) - " & If(_formDesign IsNot Nothing, _formDesign.FormTitle, ""),
            .Gadget = Nothing
        })
        If _formDesign IsNot Nothing Then
            For Each g In _formDesign.Gadgets
                Dim tdef = W9GadgetRegistry.GetTypeDef(g.GadgetType)
                _gadgetCombo.Items.Add(New GadgetComboItem() With {
                    .DisplayText = g.EnumName & " - " & If(tdef IsNot Nothing, tdef.DisplayName, "Gadget"),
                    .Gadget = g
                })
            Next
        End If
        If _gadgetCombo.Items.Count > 0 Then _gadgetCombo.SelectedIndex = 0
    End Sub

    Public Sub RefreshProperties()
        _propGrid.Refresh()
    End Sub

    ' =========================================================================
    ' Events
    ' =========================================================================
    Private Sub OnGadgetComboChanged(sender As Object, e As EventArgs)
        If _gadgetCombo.SelectedItem Is Nothing Then Return
        Dim item = TryCast(_gadgetCombo.SelectedItem, GadgetComboItem)
        If item IsNot Nothing Then
            If item.Gadget Is Nothing Then
                ShowFormProperties()
            Else
                ShowGadgetProperties(item.Gadget)
            End If
        End If
    End Sub

    Private Sub OnPropertyValueChanged(sender As Object, e As PropertyValueChangedEventArgs)
        If _selectedGadget IsNot Nothing Then
            RaiseEvent PropertyChanged(_selectedGadget, e.ChangedItem.Label)
        Else
            RaiseEvent FormPropertyChanged(e.ChangedItem.Label)
        End If
    End Sub

    ' =========================================================================
    ' Theming
    ' =========================================================================
    Public Sub ApplyTheme(isDark As Boolean)
        _isDarkTheme = isDark
        If isDark Then
            Me.BackColor = Color.FromArgb(37, 37, 38)
            _propGrid.BackColor = Color.FromArgb(37, 37, 38)
            _propGrid.LineColor = Color.FromArgb(60, 60, 60)
            _propGrid.CategoryForeColor = Color.FromArgb(200, 200, 200)
            _propGrid.ViewBackColor = Color.FromArgb(37, 37, 38)
            _propGrid.ViewForeColor = Color.FromArgb(220, 220, 220)
            _propGrid.HelpBackColor = Color.FromArgb(45, 45, 48)
            _propGrid.HelpForeColor = Color.FromArgb(200, 200, 200)
            _gadgetCombo.BackColor = Color.FromArgb(51, 51, 55)
            _gadgetCombo.ForeColor = Color.FromArgb(220, 220, 220)
        Else
            Me.BackColor = Color.FromArgb(246, 246, 246)
            _propGrid.BackColor = Color.FromArgb(246, 246, 246)
            _propGrid.LineColor = Color.FromArgb(220, 220, 220)
            _propGrid.CategoryForeColor = Color.FromArgb(60, 60, 60)
            _propGrid.ViewBackColor = Color.White
            _propGrid.ViewForeColor = Color.Black
            _propGrid.HelpBackColor = Color.FromArgb(240, 240, 240)
            _propGrid.HelpForeColor = Color.Black
            _gadgetCombo.BackColor = Color.White
            _gadgetCombo.ForeColor = Color.Black
        End If
    End Sub

    Private Class GadgetComboItem
        Public DisplayText As String = ""
        Public Gadget As W9GadgetInstance = Nothing
        Public Overrides Function ToString() As String
            Return DisplayText
        End Function
    End Class
End Class

' =============================================================================
' Property wrapper for gadget
' =============================================================================
<System.ComponentModel.TypeConverter(GetType(System.ComponentModel.ExpandableObjectConverter))>
Public Class GadgetPropertyWrapper
    Private _g As W9GadgetInstance

    Public Sub New(g As W9GadgetInstance)
        _g = g
    End Sub

    <System.ComponentModel.Category("Identity"),
     System.ComponentModel.Description("The enum name used in generated code.")>
    Public Property EnumName As String
        Get
            Return _g.EnumName
        End Get
        Set(value As String)
            _g.EnumName = value
        End Set
    End Property

    <System.ComponentModel.Category("Identity"),
     System.ComponentModel.Description("The gadget type."),
     System.ComponentModel.ReadOnly(True)>
    Public ReadOnly Property GadgetType As String
        Get
            Dim tdef = W9GadgetRegistry.GetTypeDef(_g.GadgetType)
            Return If(tdef IsNot Nothing, tdef.DisplayName, _g.GadgetType.ToString())
        End Get
    End Property

    <System.ComponentModel.Category("Identity"),
     System.ComponentModel.Description("The Window9 function name."),
     System.ComponentModel.ReadOnly(True)>
    Public ReadOnly Property W9Function As String
        Get
            Dim tdef = W9GadgetRegistry.GetTypeDef(_g.GadgetType)
            Return If(tdef IsNot Nothing, tdef.W9FunctionName, "")
        End Get
    End Property

    <System.ComponentModel.Category("Layout"),
     System.ComponentModel.Description("X position on the form.")>
    Public Property X As Integer
        Get
            Return _g.X
        End Get
        Set(value As Integer)
            _g.X = value
        End Set
    End Property

    <System.ComponentModel.Category("Layout"),
     System.ComponentModel.Description("Y position on the form.")>
    Public Property Y As Integer
        Get
            Return _g.Y
        End Get
        Set(value As Integer)
            _g.Y = value
        End Set
    End Property

    <System.ComponentModel.Category("Layout"),
     System.ComponentModel.Description("Width of the gadget.")>
    Public Property W As Integer
        Get
            Return _g.W
        End Get
        Set(value As Integer)
            _g.W = Math.Max(10, value)
        End Set
    End Property

    <System.ComponentModel.Category("Layout"),
     System.ComponentModel.Description("Height of the gadget.")>
    Public Property H As Integer
        Get
            Return _g.H
        End Get
        Set(value As Integer)
            _g.H = Math.Max(10, value)
        End Set
    End Property

    <System.ComponentModel.Category("Appearance"),
     System.ComponentModel.Description("Display text of the gadget.")>
    Public Property Text As String
        Get
            Return _g.Text
        End Get
        Set(value As String)
            _g.Text = value
        End Set
    End Property

    <System.ComponentModel.Category("Appearance"),
     System.ComponentModel.Description("Font name for the gadget."),
     System.ComponentModel.TypeConverter(GetType(FontNameConverter))>
    Public Property FontName As String
        Get
            Return _g.FontName
        End Get
        Set(value As String)
            _g.FontName = value
        End Set
    End Property

    <System.ComponentModel.Category("Appearance"),
     System.ComponentModel.Description("Font size for the gadget (0 = default 11).")>
    Public Property FontSize As Integer
        Get
            Return _g.FontSize
        End Get
        Set(value As Integer)
            _g.FontSize = Math.Max(0, Math.Min(72, value))
        End Set
    End Property

    <System.ComponentModel.Category("Appearance"),
     System.ComponentModel.Description("Background color (empty = default).")>
    Public Property BackColor As Color
        Get
            Return _g.BackColor
        End Get
        Set(value As Color)
            _g.BackColor = value
        End Set
    End Property

    <System.ComponentModel.Category("Appearance"),
     System.ComponentModel.Description("Text/foreground color (empty = default).")>
    Public Property ForeColor As Color
        Get
            Return _g.ForeColor
        End Get
        Set(value As Color)
            _g.ForeColor = value
        End Set
    End Property

    <System.ComponentModel.Category("Behavior"),
     System.ComponentModel.Description("Window style flags (e.g. BS_DEFPUSHBUTTON).")>
    Public Property Style As String
        Get
            Return _g.Style
        End Get
        Set(value As String)
            _g.Style = value
        End Set
    End Property

    <System.ComponentModel.Category("Behavior"),
     System.ComponentModel.Description("Extended style flags.")>
    Public Property ExStyle As String
        Get
            Return _g.ExStyle
        End Get
        Set(value As String)
            _g.ExStyle = value
        End Set
    End Property

    <System.ComponentModel.Category("Editor"),
     System.ComponentModel.Description("Make editor read-only.")>
    Public Property IsReadOnly As Boolean
        Get
            Return _g.IsReadOnly
        End Get
        Set(value As Boolean)
            _g.IsReadOnly = value
        End Set
    End Property

    <System.ComponentModel.Category("Editor"),
     System.ComponentModel.Description("Enable word wrap for editor.")>
    Public Property WordWrap As Boolean
        Get
            Return _g.WordWrap
        End Get
        Set(value As Boolean)
            _g.WordWrap = value
        End Set
    End Property

    <System.ComponentModel.Category("Range"),
     System.ComponentModel.Description("Minimum value (for ScrollBar/TrackBar/Spin/ProgressBar).")>
    Public Property MinValue As Integer
        Get
            Return _g.MinValue
        End Get
        Set(value As Integer)
            _g.MinValue = value
        End Set
    End Property

    <System.ComponentModel.Category("Range"),
     System.ComponentModel.Description("Maximum value.")>
    Public Property MaxValue As Integer
        Get
            Return _g.MaxValue
        End Get
        Set(value As Integer)
            _g.MaxValue = value
        End Set
    End Property

    <System.ComponentModel.Category("Range"),
     System.ComponentModel.Description("Current/default value.")>
    Public Property CurrentValue As Integer
        Get
            Return _g.CurrentValue
        End Get
        Set(value As Integer)
            _g.CurrentValue = value
        End Set
    End Property

    <System.ComponentModel.Category("Range"),
     System.ComponentModel.Description("Orientation: 0=Horizontal, 1=Vertical.")>
    Public Property Orientation As Integer
        Get
            Return _g.Orientation
        End Get
        Set(value As Integer)
            _g.Orientation = value
        End Set
    End Property

    <System.ComponentModel.Category("Misc"),
     System.ComponentModel.Description("Lock gadget position on canvas.")>
    Public Property IsLocked As Boolean
        Get
            Return _g.IsLocked
        End Get
        Set(value As Boolean)
            _g.IsLocked = value
        End Set
    End Property

    <System.ComponentModel.Category("Misc"),
     System.ComponentModel.Description("Custom tag/notes.")>
    Public Property Tag As String
        Get
            Return _g.Tag
        End Get
        Set(value As String)
            _g.Tag = value
        End Set
    End Property

    ' ---- NEW PROPERTIES ----

    <System.ComponentModel.Category("Items"),
     System.ComponentModel.Description("Initial items for ComboBox/ListBox (one per line).")>
    Public Property Items As String
        Get
            Return _g.Items
        End Get
        Set(value As String)
            _g.Items = value
        End Set
    End Property

    <System.ComponentModel.Category("Behavior"),
     System.ComponentModel.Description("Initial checked state for CheckBox/OptionButton.")>
    Public Property IsChecked As Boolean
        Get
            Return _g.IsChecked
        End Get
        Set(value As Boolean)
            _g.IsChecked = value
        End Set
    End Property

    <System.ComponentModel.Category("Behavior"),
     System.ComponentModel.Description("Password mode for StringInput (masks text).")>
    Public Property IsPassword As Boolean
        Get
            Return _g.IsPassword
        End Get
        Set(value As Boolean)
            _g.IsPassword = value
        End Set
    End Property

    <System.ComponentModel.Category("Appearance"),
     System.ComponentModel.Description("Tooltip text shown on hover.")>
    Public Property Tooltip As String
        Get
            Return _g.Tooltip
        End Get
        Set(value As String)
            _g.Tooltip = value
        End Set
    End Property

    <System.ComponentModel.Category("Behavior"),
     System.ComponentModel.Description("Enabled state (False = disabled/grayed out).")>
    Public Property IsEnabled As Boolean
        Get
            Return _g.IsEnabled
        End Get
        Set(value As Boolean)
            _g.IsEnabled = value
        End Set
    End Property

    <System.ComponentModel.Category("Behavior"),
     System.ComponentModel.Description("Visible state (False = hidden at startup).")>
    Public Property IsVisible As Boolean
        Get
            Return _g.IsVisible
        End Get
        Set(value As Boolean)
            _g.IsVisible = value
        End Set
    End Property

    <System.ComponentModel.Category("Appearance"),
     System.ComponentModel.Description("Image file path for ImageBox/ButtonImage.")>
    Public Property ImagePath As String
        Get
            Return _g.ImagePath
        End Get
        Set(value As String)
            _g.ImagePath = value
        End Set
    End Property

    <System.ComponentModel.Category("Items"),
     System.ComponentModel.Description("Tab names for PanelGadget (one per line).")>
    Public Property TabNames As String
        Get
            Return _g.TabNames
        End Get
        Set(value As String)
            _g.TabNames = value
        End Set
    End Property

    <System.ComponentModel.Category("Editor"),
     System.ComponentModel.Description("Show vertical scrollbar in editor.")>
    Public Property HasVScroll As Boolean
        Get
            Return _g.HasVScroll
        End Get
        Set(value As Boolean)
            _g.HasVScroll = value
        End Set
    End Property

    <System.ComponentModel.Category("Editor"),
     System.ComponentModel.Description("Show horizontal scrollbar in editor.")>
    Public Property HasHScroll As Boolean
        Get
            Return _g.HasHScroll
        End Get
        Set(value As Boolean)
            _g.HasHScroll = value
        End Set
    End Property

    <System.ComponentModel.Category("Events"),
     System.ComponentModel.Description("Generate OnClick event handler.")>
    Public Property OnClickEvent As Boolean
        Get
            Return _g.OnClickEvent
        End Get
        Set(value As Boolean)
            _g.OnClickEvent = value
        End Set
    End Property

    <System.ComponentModel.Category("Events"),
     System.ComponentModel.Description("Generate OnChange event handler (selection/value change).")>
    Public Property OnChangeEvent As Boolean
        Get
            Return _g.OnChangeEvent
        End Get
        Set(value As Boolean)
            _g.OnChangeEvent = value
        End Set
    End Property

    <System.ComponentModel.Category("Events"),
     System.ComponentModel.Description("Generate OnDoubleClick event handler.")>
    Public Property OnDoubleClickEvent As Boolean
        Get
            Return _g.OnDoubleClickEvent
        End Get
        Set(value As Boolean)
            _g.OnDoubleClickEvent = value
        End Set
    End Property

    <System.ComponentModel.Category("TreeView/ListView"),
     System.ComponentModel.Description("Show lines connecting tree nodes (TVS_HASLINES).")>
    Public Property HasLines As Boolean
        Get
            Return _g.HasLines
        End Get
        Set(value As Boolean)
            _g.HasLines = value
        End Set
    End Property

    <System.ComponentModel.Category("TreeView/ListView"),
     System.ComponentModel.Description("Show expand/collapse buttons (TVS_HASBUTTONS).")>
    Public Property HasButtons As Boolean
        Get
            Return _g.HasButtons
        End Get
        Set(value As Boolean)
            _g.HasButtons = value
        End Set
    End Property

    <System.ComponentModel.Category("TreeView/ListView"),
     System.ComponentModel.Description("Show checkboxes next to items.")>
    Public Property HasCheckBoxes As Boolean
        Get
            Return _g.HasCheckBoxes
        End Get
        Set(value As Boolean)
            _g.HasCheckBoxes = value
        End Set
    End Property

    <System.ComponentModel.Category("TreeView/ListView"),
     System.ComponentModel.Description("Full row selection in ListView.")>
    Public Property FullRowSelect As Boolean
        Get
            Return _g.FullRowSelect
        End Get
        Set(value As Boolean)
            _g.FullRowSelect = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _g.EnumName
    End Function
End Class

' =============================================================================
' Property wrapper for form-level properties
' =============================================================================
<System.ComponentModel.TypeConverter(GetType(System.ComponentModel.ExpandableObjectConverter))>
Public Class FormPropertyWrapper
    Private _f As W9FormDesign

    Public Sub New(f As W9FormDesign)
        _f = f
    End Sub

    <System.ComponentModel.Category("Window"),
     System.ComponentModel.Description("Window title text.")>
    Public Property FormTitle As String
        Get
            Return _f.FormTitle
        End Get
        Set(value As String)
            _f.FormTitle = value
        End Set
    End Property

    <System.ComponentModel.Category("Window"),
     System.ComponentModel.Description("Window X position.")>
    Public Property FormX As Integer
        Get
            Return _f.FormX
        End Get
        Set(value As Integer)
            _f.FormX = value
        End Set
    End Property

    <System.ComponentModel.Category("Window"),
     System.ComponentModel.Description("Window Y position.")>
    Public Property FormY As Integer
        Get
            Return _f.FormY
        End Get
        Set(value As Integer)
            _f.FormY = value
        End Set
    End Property

    <System.ComponentModel.Category("Window"),
     System.ComponentModel.Description("Window width (also used as BaseWidth).")>
    Public Property FormWidth As Integer
        Get
            Return _f.FormWidth
        End Get
        Set(value As Integer)
            _f.FormWidth = Math.Max(200, value)
            _f.BaseWidth = _f.FormWidth
        End Set
    End Property

    <System.ComponentModel.Category("Window"),
     System.ComponentModel.Description("Window height (also used as BaseHeight).")>
    Public Property FormHeight As Integer
        Get
            Return _f.FormHeight
        End Get
        Set(value As Integer)
            _f.FormHeight = Math.Max(150, value)
            _f.BaseHeight = _f.FormHeight
        End Set
    End Property

    <System.ComponentModel.Category("Window"),
     System.ComponentModel.Description("Center window on screen at startup.")>
    Public Property CenterOnScreen As Boolean
        Get
            Return _f.CenterOnScreen
        End Get
        Set(value As Boolean)
            _f.CenterOnScreen = value
        End Set
    End Property

    <System.ComponentModel.Category("Window"),
     System.ComponentModel.Description("Window background color (empty = default).")>
    Public Property FormColor As Color
        Get
            Return _f.FormColor
        End Get
        Set(value As Color)
            _f.FormColor = value
        End Set
    End Property

    <System.ComponentModel.Category("Behavior"),
     System.ComponentModel.Description("Enable proportional resize with ScaleX/ScaleY macros.")>
    Public Property ProportionalResize As Boolean
        Get
            Return _f.ProportionalResize
        End Get
        Set(value As Boolean)
            _f.ProportionalResize = value
        End Set
    End Property

    <System.ComponentModel.Category("Behavior"),
     System.ComponentModel.Description("Generate size_change handler for EventSize.")>
    Public Property HandleResize As Boolean
        Get
            Return _f.HandleResize
        End Get
        Set(value As Boolean)
            _f.HandleResize = value
        End Set
    End Property

    <System.ComponentModel.Category("Behavior"),
     System.ComponentModel.Description("Add a timer to the event loop.")>
    Public Property HasTimer As Boolean
        Get
            Return _f.HasTimer
        End Get
        Set(value As Boolean)
            _f.HasTimer = value
        End Set
    End Property

    <System.ComponentModel.Category("Behavior"),
     System.ComponentModel.Description("Timer interval in milliseconds.")>
    Public Property TimerInterval As Integer
        Get
            Return _f.TimerInterval
        End Get
        Set(value As Integer)
            _f.TimerInterval = Math.Max(100, value)
        End Set
    End Property

    <System.ComponentModel.Category("Behavior"),
     System.ComponentModel.Description("Register Enter key as default shortcut.")>
    Public Property HasKeyboardShortcut As Boolean
        Get
            Return _f.HasKeyboardShortcut
        End Get
        Set(value As Boolean)
            _f.HasKeyboardShortcut = value
        End Set
    End Property

    <System.ComponentModel.Category("Code Generation"),
     System.ComponentModel.Description("Starting value for GadgetID enum.")>
    Public Property GadgetEnumStart As Integer
        Get
            Return _f.GadgetEnumStart
        End Get
        Set(value As Integer)
            _f.GadgetEnumStart = value
        End Set
    End Property

    <System.ComponentModel.Category("Code Generation"),
     System.ComponentModel.Description("Starting value for MenuID enum.")>
    Public Property MenuEnumStart As Integer
        Get
            Return _f.MenuEnumStart
        End Get
        Set(value As Integer)
            _f.MenuEnumStart = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _f.FormTitle
    End Function
End Class

' =============================================================================
' Font name dropdown converter — lists common Windows fonts in property grid
' =============================================================================
Public Class FontNameConverter
    Inherits System.ComponentModel.StringConverter

    Public Overrides Function GetStandardValuesSupported(context As System.ComponentModel.ITypeDescriptorContext) As Boolean
        Return True
    End Function

    Public Overrides Function GetStandardValuesExclusive(context As System.ComponentModel.ITypeDescriptorContext) As Boolean
        Return False  ' Allow typing custom names too
    End Function

    Public Overrides Function GetStandardValues(context As System.ComponentModel.ITypeDescriptorContext) As System.ComponentModel.TypeConverter.StandardValuesCollection
        ' Common fonts available on Windows (and likely on Linux via font packages)
        Dim fonts As New List(Of String)()
        fonts.Add("Consolas")
        fonts.Add("Courier New")
        fonts.Add("MS Dialog")
        fonts.Add("Segoe UI")
        fonts.Add("Tahoma")
        fonts.Add("Arial")
        fonts.Add("Verdana")
        fonts.Add("Calibri")
        fonts.Add("Times New Roman")
        fonts.Add("Lucida Console")
        fonts.Add("Microsoft Sans Serif")
        fonts.Add("Trebuchet MS")
        fonts.Add("Georgia")
        fonts.Add("Comic Sans MS")
        fonts.Add("Impact")

        ' Also add installed system fonts
        Try
            For Each ff In System.Drawing.FontFamily.Families
                If Not fonts.Contains(ff.Name) Then
                    fonts.Add(ff.Name)
                End If
            Next
        Catch
            ' Ignore if font enumeration fails
        End Try

        fonts.Sort()
        Return New System.ComponentModel.TypeConverter.StandardValuesCollection(fonts)
    End Function
End Class
