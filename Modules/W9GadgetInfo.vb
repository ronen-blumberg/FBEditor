Imports System.Drawing

''' <summary>
''' Data model for Window9 visual form designer.
''' Contains definitions for all supported gadgets, their properties,
''' and the form/project model used by the designer and code generator.
''' </summary>

''' <summary>Project types supported by FBEditor.</summary>
Public Enum ProjectType
    ConsoleApp = 0
    GUIApp = 1           ' -s gui (no Window9)
    Window9FormsApp = 2  ' -s gui + Window9
End Enum

''' <summary>Available Window9 gadget types.</summary>
Public Enum W9GadgetType
    Button = 0
    TextLabel = 1
    Editor = 2
    StringInput = 3
    CheckBox = 4
    OptionButton = 5
    ComboBox = 6
    ListBox = 7
    GroupBox = 8
    ImageBox = 9
    ProgressBar = 10
    ScrollBar = 11
    TrackBar = 12
    SpinBox = 13
    TreeView = 14
    ListView = 15
    StatusBar = 16
    PanelTab = 17
    Container = 18
    Splitter = 19
    Calendar = 20
    HyperLink = 21
    WebBrowser = 22
End Enum

''' <summary>Metadata about each gadget type for the toolbox and designer.</summary>
Public Class W9GadgetTypeDef
    Public GadgetType As W9GadgetType
    Public DisplayName As String = ""
    Public W9FunctionName As String = ""
    Public DefaultWidth As Integer = 100
    Public DefaultHeight As Integer = 30
    Public DefaultText As String = ""
    Public ToolboxCategory As String = "Common Controls"
    Public ToolboxIcon As String = ""
    Public SupportsText As Boolean = True
    Public SupportsColor As Boolean = True
    Public SupportsFont As Boolean = True
    Public SupportsResize As Boolean = True
    Public IsContainer As Boolean = False
    Public HasDefaultStyle As String = ""
    Public ExtraParams As String = ""
End Class

''' <summary>Instance of a gadget placed on the designer canvas.</summary>
Public Class W9GadgetInstance
    Public ID As Integer = 0
    Public GadgetType As W9GadgetType = W9GadgetType.Button
    Public Name As String = ""             ' Variable name prefix (e.g. "giButton1")
    Public EnumName As String = ""         ' Enum member name
    Public Text As String = ""
    Public X As Integer = 10
    Public Y As Integer = 10
    Public W As Integer = 100
    Public H As Integer = 30
    Public Style As String = ""
    Public ExStyle As String = ""
    Public FontName As String = ""
    Public FontSize As Integer = 0
    Public IsReadOnly As Boolean = False
    Public WordWrap As Boolean = False
    Public BackColor As Color = Color.Empty
    Public ForeColor As Color = Color.Empty
    Public Tag As String = ""
    Public ZOrder As Integer = 0

    ' Designer-only state
    Public IsSelected As Boolean = False
    Public IsLocked As Boolean = False

    ' Scrollbar / Trackbar / Spin specifics
    Public MinValue As Integer = 0
    Public MaxValue As Integer = 100
    Public CurrentValue As Integer = 0
    Public Orientation As Integer = 0    ' 0=Horiz, 1=Vert

    ' StatusBar fields
    Public StatusBarFields As New List(Of StatusBarFieldInfo)()

    ' ListView columns
    Public ListViewColumns As New List(Of String)()

    ' ---- NEW: Additional gadget-specific properties ----

    ' ComboBox / ListBox initial items (one per line)
    Public Items As String = ""

    ' CheckBox / OptionButton initial state
    Public IsChecked As Boolean = False

    ' StringInput password mode
    Public IsPassword As Boolean = False

    ' Tooltip text
    Public Tooltip As String = ""

    ' Enabled state (False = DisableGadget at creation)
    Public IsEnabled As Boolean = True

    ' Visible state (False = HideGadget at creation)
    Public IsVisible As Boolean = True

    ' ImageBox / ButtonImage: image file path
    Public ImagePath As String = ""

    ' PanelTab: tab names (one per line)
    Public TabNames As String = ""

    ' Editor: multiline with scrollbar flags
    Public HasVScroll As Boolean = True
    Public HasHScroll As Boolean = False

    ' Events - which event handlers to generate
    Public OnClickEvent As Boolean = False
    Public OnChangeEvent As Boolean = False
    Public OnDoubleClickEvent As Boolean = False

    ' TreeView / ListView style options
    Public HasLines As Boolean = True       ' TVS_HASLINES
    Public HasButtons As Boolean = True     ' TVS_HASBUTTONS
    Public HasCheckBoxes As Boolean = False ' TVS_CHECKBOXES / LVS_EX_CHECKBOXES
    Public FullRowSelect As Boolean = False ' LVS_EX_FULLROWSELECT

    ''' <summary>Get the display bounds for painting on the canvas.</summary>
    Public ReadOnly Property Bounds As Rectangle
        Get
            Return New Rectangle(X, Y, W, H)
        End Get
    End Property

    Public Function Clone() As W9GadgetInstance
        Dim c As New W9GadgetInstance()
        c.ID = ID : c.GadgetType = GadgetType : c.Name = Name : c.EnumName = EnumName
        c.Text = Text : c.X = X : c.Y = Y : c.W = W : c.H = H
        c.Style = Style : c.ExStyle = ExStyle : c.FontName = FontName : c.FontSize = FontSize
        c.IsReadOnly = IsReadOnly : c.WordWrap = WordWrap
        c.BackColor = BackColor : c.ForeColor = ForeColor : c.Tag = Tag : c.ZOrder = ZOrder
        c.MinValue = MinValue : c.MaxValue = MaxValue : c.CurrentValue = CurrentValue
        c.Orientation = Orientation : c.IsLocked = IsLocked
        c.Items = Items : c.IsChecked = IsChecked : c.IsPassword = IsPassword
        c.Tooltip = Tooltip : c.IsEnabled = IsEnabled : c.IsVisible = IsVisible
        c.ImagePath = ImagePath : c.TabNames = TabNames
        c.HasVScroll = HasVScroll : c.HasHScroll = HasHScroll
        c.OnClickEvent = OnClickEvent : c.OnChangeEvent = OnChangeEvent
        c.OnDoubleClickEvent = OnDoubleClickEvent
        c.HasLines = HasLines : c.HasButtons = HasButtons
        c.HasCheckBoxes = HasCheckBoxes : c.FullRowSelect = FullRowSelect
        Return c
    End Function
End Class

''' <summary>Status bar field definition.</summary>
Public Class StatusBarFieldInfo
    Public Width As Integer = -1
    Public Text As String = ""
End Class

''' <summary>Menu item for the menu designer.</summary>
Public Class W9MenuItemInfo
    Public ID As Integer = 0
    Public EnumName As String = ""
    Public Text As String = ""
    Public IsSeparator As Boolean = False
    Public Children As New List(Of W9MenuItemInfo)()
    Public IsTopLevel As Boolean = False
End Class

''' <summary>
''' Complete form design model — represents one Window9 form with all its gadgets and menus.
''' This is what gets serialized/deserialized and used by the code generator.
''' </summary>
Public Class W9FormDesign
    Public FormTitle As String = "My Window9 Application"
    Public FormX As Integer = 100
    Public FormY As Integer = 50
    Public FormWidth As Integer = 800
    Public FormHeight As Integer = 600
    Public CenterOnScreen As Boolean = True
    Public FormColor As Color = Color.Empty
    Public BaseWidth As Integer = 800
    Public BaseHeight As Integer = 600
    Public ProportionalResize As Boolean = True

    Public Gadgets As New List(Of W9GadgetInstance)()
    Public MenuItems As New List(Of W9MenuItemInfo)()

    ' Event loop options
    Public HandleResize As Boolean = True
    Public HasTimer As Boolean = False
    Public TimerInterval As Integer = 1000
    Public HasKeyboardShortcut As Boolean = False
    Public DefaultButtonID As Integer = 0

    ' Enum start counters
    Public GadgetEnumStart As Integer = 100
    Public MenuEnumStart As Integer = 200
    Public ShortcutEnumStart As Integer = 1000

    Private _nextGadgetId As Integer = 101
    Private _nextMenuId As Integer = 201

    Public Function GetNextGadgetID() As Integer
        _nextGadgetId += 1
        Return _nextGadgetId - 1
    End Function

    Public Function GetNextMenuID() As Integer
        _nextMenuId += 1
        Return _nextMenuId - 1
    End Function

    ''' <summary>Remove a gadget by reference.</summary>
    Public Sub RemoveGadget(g As W9GadgetInstance)
        Gadgets.Remove(g)
    End Sub

    ''' <summary>Find the topmost gadget at a given point.</summary>
    Public Function HitTest(pt As Point) As W9GadgetInstance
        ' Search in reverse Z-order (topmost first)
        For i = Gadgets.Count - 1 To 0 Step -1
            If Gadgets(i).Bounds.Contains(pt) Then Return Gadgets(i)
        Next
        Return Nothing
    End Function

    ''' <summary>Clear selection on all gadgets.</summary>
    Public Sub ClearSelection()
        For Each g In Gadgets
            g.IsSelected = False
        Next
    End Sub
End Class

''' <summary>
''' Registry of all Window9 gadget type definitions.
''' Used by the toolbox and code generator.
''' </summary>
Public Module W9GadgetRegistry

    Private _types As List(Of W9GadgetTypeDef) = Nothing

    Public ReadOnly Property AllTypes As List(Of W9GadgetTypeDef)
        Get
            If _types Is Nothing Then BuildRegistry()
            Return _types
        End Get
    End Property

    Public Function GetTypeDef(gt As W9GadgetType) As W9GadgetTypeDef
        Return AllTypes.Find(Function(t) t.GadgetType = gt)
    End Function

    Private Sub BuildRegistry()
        _types = New List(Of W9GadgetTypeDef)()

        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.Button, .DisplayName = "Button",
            .W9FunctionName = "ButtonGadget", .DefaultWidth = 100, .DefaultHeight = 30,
            .DefaultText = "Button", .ToolboxIcon = "B", .ToolboxCategory = "Common Controls",
            .HasDefaultStyle = "BS_DEFPUSHBUTTON"
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.TextLabel, .DisplayName = "Text Label",
            .W9FunctionName = "TextGadget", .DefaultWidth = 120, .DefaultHeight = 22,
            .DefaultText = "Label", .ToolboxIcon = "T", .ToolboxCategory = "Common Controls",
            .HasDefaultStyle = "SS_NOTIFY"
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.Editor, .DisplayName = "Editor (Multiline)",
            .W9FunctionName = "EditorGadget", .DefaultWidth = 300, .DefaultHeight = 200,
            .DefaultText = "", .ToolboxIcon = "Ed", .ToolboxCategory = "Common Controls"
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.StringInput, .DisplayName = "String Input",
            .W9FunctionName = "StringGadget", .DefaultWidth = 200, .DefaultHeight = 24,
            .DefaultText = "", .ToolboxIcon = "S", .ToolboxCategory = "Common Controls"
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.CheckBox, .DisplayName = "CheckBox",
            .W9FunctionName = "CheckBoxGadget", .DefaultWidth = 130, .DefaultHeight = 24,
            .DefaultText = "CheckBox", .ToolboxIcon = "Ch", .ToolboxCategory = "Common Controls"
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.OptionButton, .DisplayName = "Option (Radio)",
            .W9FunctionName = "OptionGadget", .DefaultWidth = 130, .DefaultHeight = 24,
            .DefaultText = "Option", .ToolboxIcon = "O", .ToolboxCategory = "Common Controls"
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.ComboBox, .DisplayName = "ComboBox",
            .W9FunctionName = "ComboBoxGadget", .DefaultWidth = 180, .DefaultHeight = 24,
            .DefaultText = "", .ToolboxIcon = "Cb", .ToolboxCategory = "Common Controls",
            .SupportsText = False
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.ListBox, .DisplayName = "ListBox",
            .W9FunctionName = "ListBoxGadget", .DefaultWidth = 180, .DefaultHeight = 120,
            .DefaultText = "", .ToolboxIcon = "Lb", .ToolboxCategory = "Common Controls",
            .SupportsText = False
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.GroupBox, .DisplayName = "GroupBox",
            .W9FunctionName = "GroupGadget", .DefaultWidth = 200, .DefaultHeight = 150,
            .DefaultText = "Group", .ToolboxIcon = "G", .ToolboxCategory = "Containers",
            .IsContainer = True
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.ImageBox, .DisplayName = "Image",
            .W9FunctionName = "ImageGadget", .DefaultWidth = 100, .DefaultHeight = 100,
            .DefaultText = "", .ToolboxIcon = "Im", .ToolboxCategory = "Common Controls",
            .SupportsText = False
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.ProgressBar, .DisplayName = "ProgressBar",
            .W9FunctionName = "ProgressBarGadget", .DefaultWidth = 200, .DefaultHeight = 24,
            .DefaultText = "", .ToolboxIcon = "Pb", .ToolboxCategory = "Common Controls",
            .SupportsText = False
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.ScrollBar, .DisplayName = "ScrollBar",
            .W9FunctionName = "ScrollBarGadget", .DefaultWidth = 200, .DefaultHeight = 20,
            .DefaultText = "", .ToolboxIcon = "Sc", .ToolboxCategory = "Common Controls",
            .SupportsText = False
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.TrackBar, .DisplayName = "TrackBar (Slider)",
            .W9FunctionName = "TrackBarGadget", .DefaultWidth = 200, .DefaultHeight = 40,
            .DefaultText = "", .ToolboxIcon = "Tk", .ToolboxCategory = "Common Controls",
            .SupportsText = False
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.SpinBox, .DisplayName = "Spin",
            .W9FunctionName = "SpinGadget", .DefaultWidth = 80, .DefaultHeight = 24,
            .DefaultText = "", .ToolboxIcon = "Sp", .ToolboxCategory = "Common Controls",
            .SupportsText = False
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.TreeView, .DisplayName = "TreeView",
            .W9FunctionName = "TreeViewGadget", .DefaultWidth = 200, .DefaultHeight = 200,
            .DefaultText = "", .ToolboxIcon = "Tv", .ToolboxCategory = "Data Controls",
            .SupportsText = False
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.ListView, .DisplayName = "ListView",
            .W9FunctionName = "ListViewGadget", .DefaultWidth = 300, .DefaultHeight = 200,
            .DefaultText = "", .ToolboxIcon = "Lv", .ToolboxCategory = "Data Controls",
            .SupportsText = False
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.StatusBar, .DisplayName = "StatusBar",
            .W9FunctionName = "StatusBarGadget", .DefaultWidth = 0, .DefaultHeight = 24,
            .DefaultText = "", .ToolboxIcon = "Sb", .ToolboxCategory = "Common Controls",
            .SupportsResize = False
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.PanelTab, .DisplayName = "Panel (Tabs)",
            .W9FunctionName = "PanelGadget", .DefaultWidth = 300, .DefaultHeight = 200,
            .DefaultText = "", .ToolboxIcon = "Pn", .ToolboxCategory = "Containers",
            .SupportsText = False, .IsContainer = True
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.Container, .DisplayName = "Container",
            .W9FunctionName = "ContainerGadget", .DefaultWidth = 250, .DefaultHeight = 180,
            .DefaultText = "", .ToolboxIcon = "Cn", .ToolboxCategory = "Containers",
            .SupportsText = False, .IsContainer = True
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.Calendar, .DisplayName = "Calendar",
            .W9FunctionName = "CalendarGadget", .DefaultWidth = 220, .DefaultHeight = 180,
            .DefaultText = "", .ToolboxIcon = "Ca", .ToolboxCategory = "Common Controls",
            .SupportsText = False
        })
        _types.Add(New W9GadgetTypeDef() With {
            .GadgetType = W9GadgetType.HyperLink, .DisplayName = "HyperLink",
            .W9FunctionName = "HyperLinkGadget", .DefaultWidth = 150, .DefaultHeight = 22,
            .DefaultText = "Click here", .ToolboxIcon = "Hl", .ToolboxCategory = "Common Controls"
        })
    End Sub

    ''' <summary>Generate a default enum name for a gadget instance.</summary>
    Public Function GenerateEnumName(gt As W9GadgetType, index As Integer) As String
        Select Case gt
            Case W9GadgetType.Button : Return "giButton" & index
            Case W9GadgetType.TextLabel : Return "giLabel" & index
            Case W9GadgetType.Editor : Return "giEditor" & index
            Case W9GadgetType.StringInput : Return "giString" & index
            Case W9GadgetType.CheckBox : Return "giCheckBox" & index
            Case W9GadgetType.OptionButton : Return "giOption" & index
            Case W9GadgetType.ComboBox : Return "giComboBox" & index
            Case W9GadgetType.ListBox : Return "giListBox" & index
            Case W9GadgetType.GroupBox : Return "giGroup" & index
            Case W9GadgetType.ImageBox : Return "giImage" & index
            Case W9GadgetType.ProgressBar : Return "giProgressBar" & index
            Case W9GadgetType.ScrollBar : Return "giScrollBar" & index
            Case W9GadgetType.TrackBar : Return "giTrackBar" & index
            Case W9GadgetType.SpinBox : Return "giSpin" & index
            Case W9GadgetType.TreeView : Return "giTreeView" & index
            Case W9GadgetType.ListView : Return "giListView" & index
            Case W9GadgetType.StatusBar : Return "giStatusBar" & index
            Case W9GadgetType.PanelTab : Return "giPanel" & index
            Case W9GadgetType.Container : Return "giContainer" & index
            Case W9GadgetType.Calendar : Return "giCalendar" & index
            Case W9GadgetType.HyperLink : Return "giHyperLink" & index
            Case Else : Return "giGadget" & index
        End Select
    End Function
End Module
