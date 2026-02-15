Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO

''' <summary>
''' The main form designer panel that replaces the code editor area when in
''' Design view. Provides a Visual Studio-like layout with:
'''   - Left: Toolbox panel (gadget palette)
'''   - Center: Design canvas (the form surface)
'''   - Right: Properties panel
'''   - Bottom: Status bar with coordinates/gadget info
'''   - Top: Designer toolbar (grid, snap, generate code, menus, etc.)
'''
''' Integrates with MainForm as a swappable panel — user switches between
''' "Code" and "Design" views like in Visual Studio.
''' </summary>
Public Class FormDesignerPanel
    Inherits Panel

    ' ---- Sub-panels ----
    Private _toolbox As W9ToolboxPanel
    Private _canvas As W9DesignerCanvas
    Private _properties As W9PropertyPanel
    Private _toolbar As ToolStrip
    Private _statusLabel As Label
    Private _titleBarLabel As Label

    ' ---- Layout splitters ----
    Private _leftSplitter As SplitContainer
    Private _rightSplitter As SplitContainer

    ' ---- Data ----
    Private _formDesign As W9FormDesign
    Private _isDirty As Boolean = False

    ' ---- Events for MainForm integration ----
    Public Event DesignDirtyChanged(isDirty As Boolean)
    Public Event StatusTextChanged(text As String)

    ' =========================================================================
    ' Constructor
    ' =========================================================================
    Public Sub New()
        Me.Dock = DockStyle.Fill
        _formDesign = New W9FormDesign()
        InitializeLayout()
        WireEvents()
    End Sub

    ' =========================================================================
    ' Properties
    ' =========================================================================
    Public ReadOnly Property FormDesign As W9FormDesign
        Get
            Return _formDesign
        End Get
    End Property

    Public ReadOnly Property IsDirty As Boolean
        Get
            Return _isDirty
        End Get
    End Property

    Public ReadOnly Property Canvas As W9DesignerCanvas
        Get
            Return _canvas
        End Get
    End Property

    ' =========================================================================
    ' Layout initialization
    ' =========================================================================
    Private Sub InitializeLayout()
        Me.SuspendLayout()

        ' === Toolbar at top ===
        _toolbar = CreateToolbar()
        Me.Controls.Add(_toolbar)

        ' === Status bar at bottom ===
        _statusLabel = New Label() With {
            .Dock = DockStyle.Bottom,
            .Height = 22,
            .TextAlign = ContentAlignment.MiddleLeft,
            .Text = "  Ready — Click a gadget in the Toolbox, then draw on the form",
            .Font = New Font("Segoe UI", 8),
            .BackColor = Color.FromArgb(0, 120, 215),
            .ForeColor = Color.White,
            .Padding = New Padding(4, 0, 0, 0)
        }
        Me.Controls.Add(_statusLabel)

        ' === Main 3-panel layout using nested SplitContainers ===

        ' Right splitter: [Canvas | Properties]
        _rightSplitter = New SplitContainer() With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Vertical,
            .FixedPanel = FixedPanel.Panel2,
            .SplitterDistance = 500,
            .SplitterWidth = 4,
            .BorderStyle = BorderStyle.None
        }

        ' Left splitter: [Toolbox | rightSplitter]
        _leftSplitter = New SplitContainer() With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Vertical,
            .FixedPanel = FixedPanel.Panel1,
            .SplitterDistance = 180,
            .SplitterWidth = 4,
            .BorderStyle = BorderStyle.None
        }

        ' Create the sub-panels
        _toolbox = New W9ToolboxPanel()
        _canvas = New W9DesignerCanvas()
        _properties = New W9PropertyPanel()

        ' Canvas gets a scrollable wrapper
        Dim canvasWrapper As New Panel() With {
            .Dock = DockStyle.Fill,
            .AutoScroll = True,
            .BackColor = Color.FromArgb(60, 60, 60)  ' Dark surround
        }

        ' Decorative title bar above the canvas (not part of the design surface)
        _titleBarLabel = New Label() With {
            .Text = "  " & _formDesign.FormTitle,
            .Location = New Point(20, 5),
            .AutoSize = False,
            .Size = New Size(_formDesign.FormWidth, 26),
            .BackColor = Color.FromArgb(0, 120, 215),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 9, FontStyle.Regular),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        canvasWrapper.Controls.Add(_titleBarLabel)

        _canvas.Dock = DockStyle.None
        _canvas.Location = New Point(20, 31)
        UpdateCanvasSize()
        canvasWrapper.Controls.Add(_canvas)

        ' Assemble
        _leftSplitter.Panel1.Controls.Add(_toolbox)
        _rightSplitter.Panel1.Controls.Add(canvasWrapper)
        _rightSplitter.Panel2.Controls.Add(_properties)
        _leftSplitter.Panel2.Controls.Add(_rightSplitter)

        Me.Controls.Add(_leftSplitter)

        ' Ensure toolbar is on top
        _leftSplitter.BringToFront()
        _toolbar.SendToBack()

        ' Initialize property panel
        _properties.SetFormDesign(_formDesign)
        _properties.ShowFormProperties()
        _canvas.FormDesign = _formDesign

        Me.ResumeLayout(True)
    End Sub

    Private Function CreateToolbar() As ToolStrip
        Dim ts As New ToolStrip() With {
            .GripStyle = ToolStripGripStyle.Hidden,
            .RenderMode = ToolStripRenderMode.System,
            .Font = New Font("Segoe UI", 9)
        }

        ' Pointer mode
        Dim btnPointer = New ToolStripButton("Pointer") With {.ToolTipText = "Select/Move mode (Esc)"}
        AddHandler btnPointer.Click, Sub(s, e)
                                         _toolbox.ResetToPointer()
                                         _canvas.ClearPendingGadgetType()
                                     End Sub
        ts.Items.Add(btnPointer)
        ts.Items.Add(New ToolStripSeparator())

        ' Grid toggle
        Dim btnGrid = New ToolStripButton("Grid") With {.ToolTipText = "Show/Hide grid", .Checked = True, .CheckOnClick = True}
        AddHandler btnGrid.CheckedChanged, Sub(s, e) _canvas.ShowGrid = btnGrid.Checked
        ts.Items.Add(btnGrid)

        ' Snap toggle
        Dim btnSnap = New ToolStripButton("Snap") With {.ToolTipText = "Snap to grid", .Checked = True, .CheckOnClick = True}
        AddHandler btnSnap.CheckedChanged, Sub(s, e) _canvas.SnapToGrid = btnSnap.Checked
        ts.Items.Add(btnSnap)

        ts.Items.Add(New ToolStripSeparator())

        ' Menu designer
        Dim btnMenuDesigner = New ToolStripButton("Menus...") With {.ToolTipText = "Design menus"}
        AddHandler btnMenuDesigner.Click, AddressOf OnOpenMenuDesigner
        ts.Items.Add(btnMenuDesigner)

        ts.Items.Add(New ToolStripSeparator())

        ' Delete selected
        Dim btnDelete = New ToolStripButton("Delete") With {.ToolTipText = "Delete selected gadget (Del)"}
        AddHandler btnDelete.Click, Sub(s, e) _canvas.DeleteSelectedGadgets()
        ts.Items.Add(btnDelete)

        ts.Items.Add(New ToolStripSeparator())

        ' Z-Order
        Dim btnBringFront = New ToolStripButton("Front") With {.ToolTipText = "Bring to Front"}
        AddHandler btnBringFront.Click, Sub(s, e) _canvas.BringToFront_Gadget()
        ts.Items.Add(btnBringFront)

        Dim btnSendBack = New ToolStripButton("Back") With {.ToolTipText = "Send to Back"}
        AddHandler btnSendBack.Click, Sub(s, e) _canvas.SendToBack_Gadget()
        ts.Items.Add(btnSendBack)

        ts.Items.Add(New ToolStripSeparator())

        ' Alignment dropdown
        Dim btnAlign = New ToolStripDropDownButton("Align") With {.ToolTipText = "Align selected gadgets"}
        btnAlign.DropDownItems.Add("Align Left", Nothing, Sub(s, e) _canvas.AlignSelected(W9DesignerCanvas.AlignDirection.Left))
        btnAlign.DropDownItems.Add("Align Right", Nothing, Sub(s, e) _canvas.AlignSelected(W9DesignerCanvas.AlignDirection.Right))
        btnAlign.DropDownItems.Add("Align Top", Nothing, Sub(s, e) _canvas.AlignSelected(W9DesignerCanvas.AlignDirection.Top))
        btnAlign.DropDownItems.Add("Align Bottom", Nothing, Sub(s, e) _canvas.AlignSelected(W9DesignerCanvas.AlignDirection.Bottom))
        btnAlign.DropDownItems.Add(New ToolStripSeparator())
        btnAlign.DropDownItems.Add("Center Horizontally", Nothing, Sub(s, e) _canvas.AlignSelected(W9DesignerCanvas.AlignDirection.CenterH))
        btnAlign.DropDownItems.Add("Center Vertically", Nothing, Sub(s, e) _canvas.AlignSelected(W9DesignerCanvas.AlignDirection.CenterV))
        btnAlign.DropDownItems.Add(New ToolStripSeparator())
        btnAlign.DropDownItems.Add("Same Width", Nothing, Sub(s, e) _canvas.SizeSelected(W9DesignerCanvas.SizeDirection.Width))
        btnAlign.DropDownItems.Add("Same Height", Nothing, Sub(s, e) _canvas.SizeSelected(W9DesignerCanvas.SizeDirection.Height))
        btnAlign.DropDownItems.Add("Same Both", Nothing, Sub(s, e) _canvas.SizeSelected(W9DesignerCanvas.SizeDirection.Both))
        btnAlign.DropDownItems.Add(New ToolStripSeparator())
        btnAlign.DropDownItems.Add("Space Horizontally", Nothing, Sub(s, e) _canvas.SpaceEvenly(True))
        btnAlign.DropDownItems.Add("Space Vertically", Nothing, Sub(s, e) _canvas.SpaceEvenly(False))
        ts.Items.Add(btnAlign)

        ' Tab order
        Dim btnTabOrder = New ToolStripButton("Tab Order") With {.ToolTipText = "Toggle tab/creation order display", .CheckOnClick = True}
        AddHandler btnTabOrder.CheckedChanged, Sub(s, e)
                                                    _canvas.ToggleTabOrderDisplay()
                                                End Sub
        ts.Items.Add(btnTabOrder)

        ts.Items.Add(New ToolStripSeparator())

        ' Generate code
        Dim btnGenerate = New ToolStripButton("Generate Code") With {
            .ToolTipText = "Generate FreeBASIC + Window9 source code",
            .Font = New Font("Segoe UI", 9, FontStyle.Bold)
        }
        AddHandler btnGenerate.Click, AddressOf OnGenerateCode
        ts.Items.Add(btnGenerate)

        ' Save design
        Dim btnSaveDesign = New ToolStripButton("Save Design") With {.ToolTipText = "Save form design (.w9form)"}
        AddHandler btnSaveDesign.Click, AddressOf OnSaveDesign
        ts.Items.Add(btnSaveDesign)

        ' Load design
        Dim btnLoadDesign = New ToolStripButton("Load Design") With {.ToolTipText = "Load form design (.w9form)"}
        AddHandler btnLoadDesign.Click, AddressOf OnLoadDesign
        ts.Items.Add(btnLoadDesign)

        ts.Items.Add(New ToolStripSeparator())

        ' Preview code
        Dim btnPreview = New ToolStripButton("Preview Code") With {.ToolTipText = "Preview generated code without saving"}
        AddHandler btnPreview.Click, AddressOf OnPreviewCode
        ts.Items.Add(btnPreview)

        Return ts
    End Function

    ' =========================================================================
    ' Wire events
    ' =========================================================================
    Private Sub WireEvents()
        ' Toolbox -> Canvas
        AddHandler _toolbox.GadgetTypeSelected, AddressOf OnToolboxGadgetSelected

        ' Canvas -> Properties
        AddHandler _canvas.GadgetSelected, Sub(g As W9GadgetInstance)
                                               _properties.ShowGadgetProperties(g)
                                               UpdateStatus(g)
                                           End Sub

        AddHandler _canvas.FormSurfaceClicked, Sub()
                                                   _properties.ShowFormProperties()
                                                   SetStatus("Form selected — edit properties on the right")
                                               End Sub

        AddHandler _canvas.GadgetAdded, Sub(g As W9GadgetInstance)
                                            _properties.RefreshGadgetCombo()
                                            _properties.ShowGadgetProperties(g)
                                            _toolbox.ResetToPointer()
                                            MarkDirty()
                                        End Sub

        AddHandler _canvas.GadgetDeleted, Sub(g As W9GadgetInstance)
                                              _properties.RefreshGadgetCombo()
                                              _properties.ShowFormProperties()
                                              MarkDirty()
                                          End Sub

        AddHandler _canvas.GadgetMoved, Sub(g As W9GadgetInstance)
                                            _properties.RefreshProperties()
                                            UpdateStatus(g)
                                        End Sub

        AddHandler _canvas.GadgetResized, Sub(g As W9GadgetInstance)
                                              _properties.RefreshProperties()
                                              UpdateStatus(g)
                                          End Sub

        AddHandler _canvas.DesignChanged, Sub() MarkDirty()

        ' Properties -> Canvas
        AddHandler _properties.PropertyChanged, Sub(g As W9GadgetInstance, propName As String)
                                                    _canvas.RefreshDesign()
                                                    _properties.RefreshGadgetCombo()
                                                    MarkDirty()
                                                End Sub

        AddHandler _properties.FormPropertyChanged, Sub(propName As String)
                                                        UpdateCanvasSize()
                                                        _canvas.RefreshDesign()
                                                        MarkDirty()
                                                    End Sub
    End Sub

    Private Sub OnToolboxGadgetSelected(gadgetType As W9GadgetType)
        _canvas.SetPendingGadgetType(gadgetType)
        Dim tdef = W9GadgetRegistry.GetTypeDef(gadgetType)
        If tdef IsNot Nothing Then
            SetStatus("Draw " & tdef.DisplayName & " — click and drag on the form canvas")
        End If
    End Sub

    ' =========================================================================
    ' Toolbar actions
    ' =========================================================================
    Private Sub OnOpenMenuDesigner(sender As Object, e As EventArgs)
        Dim dlg As New W9MenuDesigner(_formDesign)
        If dlg.ShowDialog(Me.FindForm()) = DialogResult.OK Then
            _formDesign.MenuItems = dlg.ResultMenuItems
            MarkDirty()
            SetStatus("Menus updated — " & _formDesign.MenuItems.Count & " top-level menu(s)")
        End If
        dlg.Dispose()
    End Sub

    Private Sub OnGenerateCode(sender As Object, e As EventArgs)
        Dim code = W9CodeGenerator.GenerateCode(_formDesign)

        ' Ask user where to save
        Using dlg As New SaveFileDialog()
            dlg.Title = "Save Generated Window9 Code"
            dlg.Filter = "FreeBASIC Files (*.bas)|*.bas|All Files (*.*)|*.*"
            dlg.DefaultExt = "bas"
            dlg.FileName = SanitizeFilename(_formDesign.FormTitle) & ".bas"

            If dlg.ShowDialog(Me.FindForm()) = DialogResult.OK Then
                File.WriteAllText(dlg.FileName, code)
                SetStatus("Code generated and saved: " & dlg.FileName)
                MessageBox.Show("Code generated successfully!" & vbCrLf & vbCrLf &
                                "File: " & dlg.FileName & vbCrLf &
                                "Gadgets: " & _formDesign.Gadgets.Count & vbCrLf &
                                "Menus: " & _formDesign.MenuItems.Count,
                                "Code Generated", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Using
    End Sub

    Private Sub OnPreviewCode(sender As Object, e As EventArgs)
        Dim code = W9CodeGenerator.GenerateCode(_formDesign)

        ' Show in a preview window
        Dim preview As New Form() With {
            .Text = "Generated Code Preview",
            .Size = New Size(800, 600),
            .StartPosition = FormStartPosition.CenterParent,
            .Font = New Font("Segoe UI", 9)
        }
        Dim txtCode As New TextBox() With {
            .Dock = DockStyle.Fill,
            .Multiline = True,
            .ScrollBars = ScrollBars.Both,
            .WordWrap = False,
            .Font = New Font("Consolas", 10),
            .BackColor = Color.FromArgb(30, 30, 30),
            .ForeColor = Color.FromArgb(220, 220, 170),
            .Text = code,
            .ReadOnly = True
        }
        Dim btnCopy As New Button() With {
            .Text = "Copy to Clipboard",
            .Dock = DockStyle.Bottom,
            .Height = 35,
            .Font = New Font("Segoe UI", 10)
        }
        AddHandler btnCopy.Click, Sub(s, ev)
                                      Clipboard.SetText(code)
                                      MessageBox.Show("Code copied to clipboard!", "Copied", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                  End Sub
        preview.Controls.Add(txtCode)
        preview.Controls.Add(btnCopy)
        preview.ShowDialog(Me.FindForm())
        preview.Dispose()
    End Sub

    Private Sub OnSaveDesign(sender As Object, e As EventArgs)
        Using dlg As New SaveFileDialog()
            dlg.Title = "Save Form Design"
            dlg.Filter = "Window9 Form Design (*.w9form)|*.w9form|All Files (*.*)|*.*"
            dlg.DefaultExt = "w9form"
            dlg.FileName = SanitizeFilename(_formDesign.FormTitle) & ".w9form"

            If dlg.ShowDialog(Me.FindForm()) = DialogResult.OK Then
                If ProjectManager.SaveFormDesign(_formDesign, dlg.FileName) Then
                    _isDirty = False
                    RaiseEvent DesignDirtyChanged(False)
                    SetStatus("Design saved: " & dlg.FileName)
                Else
                    MessageBox.Show("Failed to save design file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End Using
    End Sub

    Private Sub OnLoadDesign(sender As Object, e As EventArgs)
        If _isDirty Then
            Dim result = MessageBox.Show("Current design has unsaved changes. Continue?",
                                         "Load Design", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If result = DialogResult.No Then Return
        End If

        Using dlg As New OpenFileDialog()
            dlg.Title = "Load Form Design"
            dlg.Filter = "Window9 Form Design (*.w9form)|*.w9form|All Files (*.*)|*.*"

            If dlg.ShowDialog(Me.FindForm()) = DialogResult.OK Then
                Dim loaded = ProjectManager.LoadFormDesign(dlg.FileName)
                If loaded IsNot Nothing Then
                    LoadDesign(loaded)
                    SetStatus("Design loaded: " & dlg.FileName)
                Else
                    MessageBox.Show("Failed to load design file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End Using
    End Sub

    ' =========================================================================
    ' Public methods for MainForm integration
    ' =========================================================================

    ''' <summary>Load a form design into the designer.</summary>
    Public Sub LoadDesign(design As W9FormDesign)
        _formDesign = If(design, New W9FormDesign())
        _canvas.FormDesign = _formDesign
        _properties.SetFormDesign(_formDesign)
        _properties.ShowFormProperties()
        UpdateCanvasSize()
        _isDirty = False
        RaiseEvent DesignDirtyChanged(False)
    End Sub

    ''' <summary>Create a new empty form design.</summary>
    Public Sub NewDesign()
        LoadDesign(New W9FormDesign())
        SetStatus("New form design created")
    End Sub

    ''' <summary>Get the generated code as string (for inserting into editor).</summary>
    Public Function GetGeneratedCode() As String
        Return W9CodeGenerator.GenerateCode(_formDesign)
    End Function

    ''' <summary>Apply theme (dark/light) to all designer sub-panels.</summary>
    Public Sub ApplyTheme(isDark As Boolean)
        _toolbox.ApplyTheme(isDark)
        _properties.ApplyTheme(isDark)

        If isDark Then
            _toolbar.BackColor = Color.FromArgb(45, 45, 48)
            _toolbar.ForeColor = Color.FromArgb(220, 220, 220)
            _statusLabel.BackColor = Color.FromArgb(0, 100, 180)
            If _rightSplitter IsNot Nothing Then
                _rightSplitter.Panel1.BackColor = Color.FromArgb(45, 45, 48)
            End If
        Else
            _toolbar.BackColor = Color.FromArgb(240, 240, 240)
            _toolbar.ForeColor = Color.Black
            _statusLabel.BackColor = Color.FromArgb(0, 120, 215)
            If _rightSplitter IsNot Nothing Then
                _rightSplitter.Panel1.BackColor = Color.FromArgb(60, 60, 60)
            End If
        End If
    End Sub

    ' =========================================================================
    ' Helpers
    ' =========================================================================
    Private Sub UpdateCanvasSize()
        _canvas.Size = New Size(_formDesign.FormWidth + 40, _formDesign.FormHeight + 40)
        If _titleBarLabel IsNot Nothing Then
            _titleBarLabel.Size = New Size(_formDesign.FormWidth, 26)
            _titleBarLabel.Text = "  " & _formDesign.FormTitle
        End If
    End Sub

    Private Sub MarkDirty()
        _isDirty = True
        RaiseEvent DesignDirtyChanged(True)
    End Sub

    Private Sub SetStatus(text As String)
        _statusLabel.Text = "  " & text
        RaiseEvent StatusTextChanged(text)
    End Sub

    Private Sub UpdateStatus(g As W9GadgetInstance)
        If g Is Nothing Then Return
        Dim tdef = W9GadgetRegistry.GetTypeDef(g.GadgetType)
        Dim typeName = If(tdef IsNot Nothing, tdef.DisplayName, "Gadget")
        SetStatus(g.EnumName & " (" & typeName & ") — X:" & g.X & " Y:" & g.Y & " W:" & g.W & " H:" & g.H)
    End Sub

    Private Function SanitizeFilename(s As String) As String
        If s Is Nothing Then Return "untitled"
        Dim result As New System.Text.StringBuilder()
        For Each c In s
            If Char.IsLetterOrDigit(c) OrElse c = "_"c OrElse c = "-"c Then
                result.Append(c)
            ElseIf c = " "c Then
                result.Append("_")
            End If
        Next
        If result.Length = 0 Then Return "untitled"
        Return result.ToString().ToLower()
    End Function
End Class
