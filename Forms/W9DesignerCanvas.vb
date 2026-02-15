Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

''' <summary>
''' Visual form designer canvas — the main design surface where users place,
''' move, resize, and arrange Window9 gadgets. Mimics Visual Studio's form designer.
''' Features: grid snapping, drag to move, resize handles, multi-select,
''' delete, copy/paste, and visual representation of each gadget type.
''' </summary>
Public Class W9DesignerCanvas
    Inherits Panel

    ' ---- Data ----
    Private _formDesign As W9FormDesign
    Private _selectedGadget As W9GadgetInstance = Nothing
    Private _multiSelection As New List(Of W9GadgetInstance)()

    ' ---- Drag state ----
    Private _isDragging As Boolean = False
    Private _isResizing As Boolean = False
    Private _isDrawing As Boolean = False       ' Drawing a new gadget from toolbox
    Private _dragStart As Point
    Private _dragOffset As Point
    Private _resizeHandle As ResizeHandleType = ResizeHandleType.None
    Private _drawStartPoint As Point
    Private _drawCurrentPoint As Point

    ' ---- New gadget being drawn ----
    Private _pendingGadgetType As W9GadgetType? = Nothing

    ' ---- Visual settings ----
    Private _gridSize As Integer = 8
    Private _snapToGrid As Boolean = True
    Private _showGrid As Boolean = True
    Private _zoom As Single = 1.0F

    ' ---- Colors ----
    Private _formBackColor As Color = Color.FromArgb(240, 240, 240)
    Private _gridColor As Color = Color.FromArgb(220, 220, 220)
    Private _selectionColor As Color = Color.FromArgb(0, 120, 215)
    Private _handleColor As Color = Color.White
    Private _handleBorderColor As Color = Color.FromArgb(0, 120, 215)

    ' ---- Resize handles ----
    Private Const HANDLE_SIZE As Integer = 9

    ' ---- Title bar offset (gadgets are below the simulated title bar) ----
    Private Const TITLE_BAR_H As Integer = 30

    ' ---- Non-client area overhead (title bar + borders on Windows 10/11) ----
    ' These approximate the space consumed by the OS window frame
    Private Const NC_TOP As Integer = 31      ' Title bar height
    Private Const NC_BOTTOM As Integer = 8    ' Bottom border
    Private Const NC_SIDES As Integer = 8     ' Left + right border (each side)

    Public Enum ResizeHandleType
        None = 0
        TopLeft
        TopCenter
        TopRight
        MiddleLeft
        MiddleRight
        BottomLeft
        BottomCenter
        BottomRight
    End Enum

    ' ---- Events ----
    Public Event GadgetSelected(gadget As W9GadgetInstance)
    Public Event GadgetMoved(gadget As W9GadgetInstance)
    Public Event GadgetResized(gadget As W9GadgetInstance)
    Public Event GadgetAdded(gadget As W9GadgetInstance)
    Public Event GadgetDeleted(gadget As W9GadgetInstance)
    Public Event FormSurfaceClicked()
    Public Event DesignChanged()

    ' ---- Undo stack ----
    Private _undoStack As New Stack(Of List(Of W9GadgetInstance))()
    Private _redoStack As New Stack(Of List(Of W9GadgetInstance))()

    ' ---- Clipboard ----
    Private _clipboard As W9GadgetInstance = Nothing

    ' =========================================================================
    ' Constructor
    ' =========================================================================
    Public Sub New()
        Me.DoubleBuffered = True
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or
                    ControlStyles.UserPaint Or
                    ControlStyles.OptimizedDoubleBuffer Or
                    ControlStyles.Selectable, True)
        Me.BackColor = _formBackColor
        Me.AllowDrop = True
        Me.AutoScroll = False  ' Wrapper panel handles scrolling
        Me.TabStop = True

        _formDesign = New W9FormDesign()
        BuildContextMenu()
    End Sub

    ' =========================================================================
    ' Properties
    ' =========================================================================
    Public Property FormDesign As W9FormDesign
        Get
            Return _formDesign
        End Get
        Set(value As W9FormDesign)
            _formDesign = If(value, New W9FormDesign())
            _selectedGadget = Nothing
            _multiSelection.Clear()
            Invalidate()
        End Set
    End Property

    Public Property SelectedGadget As W9GadgetInstance
        Get
            Return _selectedGadget
        End Get
        Set(value As W9GadgetInstance)
            _formDesign.ClearSelection()
            _selectedGadget = value
            If value IsNot Nothing Then value.IsSelected = True
            Invalidate()
            RaiseEvent GadgetSelected(value)
        End Set
    End Property

    Public Property GridSize As Integer
        Get
            Return _gridSize
        End Get
        Set(value As Integer)
            _gridSize = Math.Max(4, Math.Min(32, value))
            Invalidate()
        End Set
    End Property

    Public Property SnapToGrid As Boolean
        Get
            Return _snapToGrid
        End Get
        Set(value As Boolean)
            _snapToGrid = value
        End Set
    End Property

    Public Property ShowGrid As Boolean
        Get
            Return _showGrid
        End Get
        Set(value As Boolean)
            _showGrid = value
            Invalidate()
        End Set
    End Property

    ''' <summary>Set the gadget type to draw next (from toolbox click).</summary>
    Public Sub SetPendingGadgetType(gt As W9GadgetType)
        _pendingGadgetType = gt
        Me.Cursor = Cursors.Cross
    End Sub

    Public Sub ClearPendingGadgetType()
        _pendingGadgetType = Nothing
        Me.Cursor = Cursors.Default
    End Sub

    ' =========================================================================
    ' Painting
    ' =========================================================================
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

        ' Canvas represents the CLIENT AREA of the Window9 window.
        ' Y=0 here = Y=0 in generated code = top of client area (below real title bar).
        Dim formRect As New Rectangle(0, 0,
            CInt(_formDesign.FormWidth * _zoom),
            CInt(_formDesign.FormHeight * _zoom))
        Using fb As New SolidBrush(_formBackColor)
            g.FillRectangle(fb, formRect)
        End Using

        ' Grid
        If _showGrid Then DrawGrid(g, formRect)

        ' Form border
        Using fp As New Pen(Color.FromArgb(100, 100, 100), 1)
            fp.DashStyle = DashStyle.Dash
            g.DrawRectangle(fp, formRect)
        End Using

        ' Draw all gadgets at their raw coordinates
        ' Y=10 here = Y=10 in Window9 = 10px below title bar in running app
        For Each gad In _formDesign.Gadgets
            DrawGadget(g, gad)
        Next

        ' Selection handles
        If _selectedGadget IsNot Nothing Then
            DrawSelectionHandles(g, _selectedGadget)
        End If

        ' Multi-selection
        For Each sel In _multiSelection
            If sel IsNot _selectedGadget Then
                DrawSelectionHandles(g, sel)
            End If
        Next

        ' Drawing rectangle (when placing new gadget)
        If _isDrawing Then
            Dim drawRect = GetNormalizedRect(_drawStartPoint, _drawCurrentPoint)
            Using dp As New Pen(_selectionColor, 1)
                dp.DashStyle = DashStyle.Dash
                g.DrawRectangle(dp, drawRect)
            End Using
        End If

        ' Lock indicator on locked gadgets
        For Each gad In _formDesign.Gadgets
            If gad.IsLocked Then
                Dim lx = CInt(gad.X * _zoom) + 2
                Dim ly = CInt(gad.Y * _zoom) + 2
                Using lockBrush As New SolidBrush(Color.FromArgb(180, 255, 100, 100))
                    g.FillEllipse(lockBrush, lx, ly, 12, 12)
                End Using
                Using lockFont As New Font("Segoe UI", 7, FontStyle.Bold)
                    g.DrawString("L", lockFont, Brushes.White, lx + 1, ly)
                End Using
            End If
        Next

        ' Tab order overlay
        If _showTabOrder Then
            For i = 0 To _formDesign.Gadgets.Count - 1
                Dim gad = _formDesign.Gadgets(i)
                Dim cx = CInt((gad.X + gad.W / 2) * _zoom)
                Dim cy = CInt((gad.Y + gad.H / 2) * _zoom)
                Dim numStr = (i + 1).ToString()
                ' Blue circle with number
                Using tabBrush As New SolidBrush(Color.FromArgb(210, 0, 100, 220))
                    g.FillEllipse(tabBrush, cx - 12, cy - 12, 24, 24)
                End Using
                Using tabPen As New Pen(Color.White, 1.5F)
                    g.DrawEllipse(tabPen, cx - 12, cy - 12, 24, 24)
                End Using
                Using tabFont As New Font("Segoe UI", 9, FontStyle.Bold)
                    Dim sz = g.MeasureString(numStr, tabFont)
                    g.DrawString(numStr, tabFont, Brushes.White, cx - sz.Width / 2, cy - sz.Height / 2)
                End Using
            Next
        End If
    End Sub

    Private Sub DrawGrid(g As Graphics, area As Rectangle)
        Using gp As New Pen(_gridColor, 1)
            Dim gs = CInt(_gridSize * _zoom)
            If gs < 4 Then Return
            For x = area.X To area.X + area.Width Step gs
                For y = area.Y To area.Y + area.Height Step gs
                    g.FillRectangle(New SolidBrush(_gridColor), x, y, 1, 1)
                Next
            Next
        End Using
    End Sub

    ''' <summary>Draw a gadget as it would appear on the form (simplified visual).</summary>
    Private Sub DrawGadget(g As Graphics, gad As W9GadgetInstance)
        Dim rect As New Rectangle(
            CInt(gad.X * _zoom), CInt(gad.Y * _zoom),
            CInt(gad.W * _zoom), CInt(gad.H * _zoom))
        Dim tdef = W9GadgetRegistry.GetTypeDef(gad.GadgetType)
        Dim displayText = If(Not String.IsNullOrEmpty(gad.Text), gad.Text, gad.EnumName)
        Dim font = New Font("Segoe UI", Math.Max(7, 8 * _zoom), FontStyle.Regular)

        Select Case gad.GadgetType
            Case W9GadgetType.Button
                DrawButtonGadget(g, rect, displayText, font)

            Case W9GadgetType.TextLabel
                DrawLabelGadget(g, rect, displayText, font)

            Case W9GadgetType.Editor
                DrawEditorGadget(g, rect, displayText, font, gad.IsReadOnly)

            Case W9GadgetType.StringInput
                DrawStringGadget(g, rect, displayText, font)

            Case W9GadgetType.CheckBox
                DrawCheckBoxGadget(g, rect, displayText, font)

            Case W9GadgetType.OptionButton
                DrawOptionGadget(g, rect, displayText, font)

            Case W9GadgetType.ComboBox
                DrawComboBoxGadget(g, rect, font)

            Case W9GadgetType.ListBox
                DrawListBoxGadget(g, rect, font)

            Case W9GadgetType.GroupBox
                DrawGroupGadget(g, rect, displayText, font)

            Case W9GadgetType.ProgressBar
                DrawProgressBarGadget(g, rect)

            Case W9GadgetType.ScrollBar
                DrawScrollBarGadget(g, rect, gad.Orientation)

            Case W9GadgetType.TrackBar
                DrawTrackBarGadget(g, rect)

            Case W9GadgetType.StatusBar
                DrawStatusBarGadget(g, rect, gad)

            Case W9GadgetType.PanelTab
                DrawPanelGadget(g, rect, font)

            Case W9GadgetType.Container
                DrawContainerGadget(g, rect, font, gad.EnumName)

            Case Else
                DrawGenericGadget(g, rect, displayText, font, tdef)
        End Select

        font.Dispose()
    End Sub

    ' ---- Individual gadget painters ----

    Private Sub DrawButtonGadget(g As Graphics, r As Rectangle, text As String, f As Font)
        Using br As New SolidBrush(Color.FromArgb(225, 225, 225))
            g.FillRectangle(br, r)
        End Using
        ControlPaint.DrawBorder3D(g, r, Border3DStyle.Raised)
        Dim sf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
        g.DrawString(text, f, Brushes.Black, New RectangleF(r.X, r.Y, r.Width, r.Height), sf)
    End Sub

    Private Sub DrawLabelGadget(g As Graphics, r As Rectangle, text As String, f As Font)
        g.DrawString(text, f, Brushes.Black, r.X + 2, r.Y + 2)
        Using bp As New Pen(Color.FromArgb(180, 180, 180), 1)
            bp.DashStyle = DashStyle.Dot
            g.DrawRectangle(bp, r)
        End Using
    End Sub

    Private Sub DrawEditorGadget(g As Graphics, r As Rectangle, text As String, f As Font, isReadOnlyEditor As Boolean)
        Dim bgColor = If(isReadOnlyEditor, Color.FromArgb(245, 245, 245), Color.White)
        Using br As New SolidBrush(bgColor)
            g.FillRectangle(br, r)
        End Using
        ControlPaint.DrawBorder3D(g, r, Border3DStyle.Sunken)
        If Not String.IsNullOrEmpty(text) Then
            Dim textRect = Rectangle.Inflate(r, -3, -3)
            g.DrawString(text, f, Brushes.Gray, New RectangleF(textRect.X, textRect.Y, textRect.Width, textRect.Height))
        End If
        ' Editor label
        Using ef = New Font("Segoe UI", 6.5F * _zoom, FontStyle.Italic)
            Dim label = If(isReadOnlyEditor, "Editor (ReadOnly)", "Editor")
            g.DrawString(label, ef, Brushes.Gray, r.X + 3, r.Bottom - 14 * _zoom)
        End Using
    End Sub

    Private Sub DrawStringGadget(g As Graphics, r As Rectangle, text As String, f As Font)
        g.FillRectangle(Brushes.White, r)
        ControlPaint.DrawBorder3D(g, r, Border3DStyle.Sunken)
        If Not String.IsNullOrEmpty(text) Then
            g.DrawString(text, f, Brushes.Gray, r.X + 3, r.Y + 2)
        End If
    End Sub

    Private Sub DrawCheckBoxGadget(g As Graphics, r As Rectangle, text As String, f As Font)
        Using bp As New Pen(Color.FromArgb(180, 180, 180), 1)
            bp.DashStyle = DashStyle.Dot
            g.DrawRectangle(bp, r)
        End Using
        Dim boxRect As New Rectangle(r.X + 2, r.Y + (r.Height - 14) \ 2, 14, 14)
        g.FillRectangle(Brushes.White, boxRect)
        g.DrawRectangle(Pens.Gray, boxRect)
        g.DrawString(text, f, Brushes.Black, r.X + 20, r.Y + 2)
    End Sub

    Private Sub DrawOptionGadget(g As Graphics, r As Rectangle, text As String, f As Font)
        Using bp As New Pen(Color.FromArgb(180, 180, 180), 1)
            bp.DashStyle = DashStyle.Dot
            g.DrawRectangle(bp, r)
        End Using
        Dim circleRect As New Rectangle(r.X + 2, r.Y + (r.Height - 14) \ 2, 14, 14)
        g.FillEllipse(Brushes.White, circleRect)
        g.DrawEllipse(Pens.Gray, circleRect)
        g.DrawString(text, f, Brushes.Black, r.X + 20, r.Y + 2)
    End Sub

    Private Sub DrawComboBoxGadget(g As Graphics, r As Rectangle, f As Font)
        g.FillRectangle(Brushes.White, r)
        ControlPaint.DrawBorder3D(g, r, Border3DStyle.Sunken)
        ' Draw dropdown arrow
        Dim arrowRect As New Rectangle(r.Right - 20, r.Y, 20, r.Height)
        Using br As New SolidBrush(Color.FromArgb(225, 225, 225))
            g.FillRectangle(br, arrowRect)
        End Using
        g.DrawLine(Pens.Gray, arrowRect.X, r.Y, arrowRect.X, r.Bottom)
        ' Arrow triangle
        Dim cx = arrowRect.X + 10
        Dim cy = r.Y + r.Height \ 2
        g.FillPolygon(Brushes.Black, New Point() {New Point(cx - 4, cy - 2), New Point(cx + 4, cy - 2), New Point(cx, cy + 3)})
    End Sub

    Private Sub DrawListBoxGadget(g As Graphics, r As Rectangle, f As Font)
        g.FillRectangle(Brushes.White, r)
        ControlPaint.DrawBorder3D(g, r, Border3DStyle.Sunken)
        Using ef = New Font("Segoe UI", 6.5F * _zoom, FontStyle.Italic)
            g.DrawString("ListBox", ef, Brushes.Gray, r.X + 3, r.Y + 3)
        End Using
    End Sub

    Private Sub DrawGroupGadget(g As Graphics, r As Rectangle, text As String, f As Font)
        Using bp As New Pen(Color.FromArgb(160, 160, 160), 1)
            Dim textSize = g.MeasureString(text, f)
            ' Draw the group border with text gap
            Dim topY = r.Y + CInt(textSize.Height / 2)
            g.DrawLine(bp, r.X, topY, r.X + 8, topY)
            g.DrawString(text, f, Brushes.Black, r.X + 10, r.Y)
            g.DrawLine(bp, r.X + 14 + CInt(textSize.Width), topY, r.Right, topY)
            g.DrawLine(bp, r.X, topY, r.X, r.Bottom)
            g.DrawLine(bp, r.Right, topY, r.Right, r.Bottom)
            g.DrawLine(bp, r.X, r.Bottom, r.Right, r.Bottom)
        End Using
    End Sub

    Private Sub DrawProgressBarGadget(g As Graphics, r As Rectangle)
        g.FillRectangle(Brushes.White, r)
        ControlPaint.DrawBorder3D(g, r, Border3DStyle.Sunken)
        ' Draw partial fill
        Dim fillRect As New Rectangle(r.X + 2, r.Y + 2, CInt((r.Width - 4) * 0.6), r.Height - 4)
        Using br As New SolidBrush(Color.FromArgb(6, 176, 37))
            g.FillRectangle(br, fillRect)
        End Using
    End Sub

    Private Sub DrawScrollBarGadget(g As Graphics, r As Rectangle, orientation As Integer)
        Using br As New SolidBrush(Color.FromArgb(230, 230, 230))
            g.FillRectangle(br, r)
        End Using
        g.DrawRectangle(Pens.Gray, r)
        ' Draw thumb
        If orientation = 0 Then ' Horizontal
            Dim thumbRect As New Rectangle(r.X + r.Width \ 3, r.Y + 2, r.Width \ 3, r.Height - 4)
            Using tb As New SolidBrush(Color.FromArgb(190, 190, 190))
                g.FillRectangle(tb, thumbRect)
            End Using
        Else ' Vertical
            Dim thumbRect As New Rectangle(r.X + 2, r.Y + r.Height \ 3, r.Width - 4, r.Height \ 3)
            Using tb As New SolidBrush(Color.FromArgb(190, 190, 190))
                g.FillRectangle(tb, thumbRect)
            End Using
        End If
    End Sub

    Private Sub DrawTrackBarGadget(g As Graphics, r As Rectangle)
        Using bp As New Pen(Color.FromArgb(180, 180, 180), 1)
            bp.DashStyle = DashStyle.Dot
            g.DrawRectangle(bp, r)
        End Using
        ' Track line
        Dim trackY = r.Y + r.Height \ 2
        g.DrawLine(Pens.Gray, r.X + 10, trackY, r.Right - 10, trackY)
        ' Thumb
        Dim thumbX = r.X + r.Width \ 2
        Dim thumbRect As New Rectangle(thumbX - 5, trackY - 8, 10, 16)
        Using br As New SolidBrush(Color.FromArgb(200, 200, 200))
            g.FillRectangle(br, thumbRect)
        End Using
        g.DrawRectangle(Pens.Gray, thumbRect)
    End Sub

    Private Sub DrawStatusBarGadget(g As Graphics, r As Rectangle, gad As W9GadgetInstance)
        ' StatusBar always at bottom of form
        Dim sbRect As New Rectangle(0, CInt((_formDesign.FormHeight - 24) * _zoom),
                                    CInt(_formDesign.FormWidth * _zoom), CInt(24 * _zoom))
        Using br As New SolidBrush(Color.FromArgb(225, 225, 225))
            g.FillRectangle(br, sbRect)
        End Using
        g.DrawLine(Pens.Gray, sbRect.X, sbRect.Y, sbRect.Right, sbRect.Y)
        Using sf = New Font("Segoe UI", 7 * _zoom)
            If gad.StatusBarFields.Count > 0 Then
                Dim xOff = 4
                For Each field In gad.StatusBarFields
                    g.DrawString(field.Text, sf, Brushes.Black, sbRect.X + xOff, sbRect.Y + 4)
                    If field.Width > 0 Then
                        xOff += CInt(field.Width * _zoom)
                        g.DrawLine(Pens.Gray, sbRect.X + xOff, sbRect.Y + 2, sbRect.X + xOff, sbRect.Bottom - 2)
                        xOff += 4
                    End If
                Next
            Else
                g.DrawString("StatusBar", sf, Brushes.Gray, sbRect.X + 4, sbRect.Y + 4)
            End If
        End Using
    End Sub

    Private Sub DrawPanelGadget(g As Graphics, r As Rectangle, f As Font)
        g.FillRectangle(Brushes.White, r)
        g.DrawRectangle(Pens.Gray, r)
        ' Tab header
        Dim tabRect As New Rectangle(r.X, r.Y, 60, 22)
        Using br As New SolidBrush(Color.FromArgb(240, 240, 240))
            g.FillRectangle(br, tabRect)
        End Using
        g.DrawRectangle(Pens.Gray, tabRect)
        Using ef = New Font("Segoe UI", 6.5F * _zoom)
            g.DrawString("Tab 1", ef, Brushes.Black, tabRect.X + 4, tabRect.Y + 3)
        End Using
    End Sub

    Private Sub DrawContainerGadget(g As Graphics, r As Rectangle, f As Font, name As String)
        Using bp As New Pen(Color.FromArgb(180, 180, 180), 1)
            bp.DashStyle = DashStyle.DashDot
            g.DrawRectangle(bp, r)
        End Using
        Using ef = New Font("Segoe UI", 6.5F * _zoom, FontStyle.Italic)
            g.DrawString(name, ef, Brushes.Gray, r.X + 3, r.Y + 3)
        End Using
    End Sub

    Private Sub DrawGenericGadget(g As Graphics, r As Rectangle, text As String, f As Font, tdef As W9GadgetTypeDef)
        Using br As New SolidBrush(Color.FromArgb(250, 250, 250))
            g.FillRectangle(br, r)
        End Using
        g.DrawRectangle(Pens.Gray, r)
        Dim label = If(tdef IsNot Nothing, tdef.DisplayName, "Gadget")
        Using ef = New Font("Segoe UI", 6.5F * _zoom, FontStyle.Italic)
            g.DrawString(label, ef, Brushes.Gray, r.X + 3, r.Y + 3)
        End Using
        If Not String.IsNullOrEmpty(text) Then
            g.DrawString(text, f, Brushes.Black, r.X + 3, r.Y + 16)
        End If
    End Sub

    ''' <summary>Draw selection handles around a gadget.</summary>
    Private Sub DrawSelectionHandles(g As Graphics, gad As W9GadgetInstance)
        Dim r As New Rectangle(CInt(gad.X * _zoom), CInt(gad.Y * _zoom),
                               CInt(gad.W * _zoom), CInt(gad.H * _zoom))
        ' Selection border
        Using sp As New Pen(_selectionColor, 1)
            sp.DashStyle = DashStyle.Dash
            g.DrawRectangle(sp, r)
        End Using

        ' 8 handles
        Dim handleRects = GetHandleRects(r)
        For Each hr In handleRects
            g.FillRectangle(New SolidBrush(_handleColor), hr)
            g.DrawRectangle(New Pen(_handleBorderColor, 1), hr)
        Next
    End Sub

    Private Function GetHandleRects(r As Rectangle) As Rectangle()
        Dim hs = HANDLE_SIZE
        Dim half = hs \ 2
        Return {
            New Rectangle(r.X - half, r.Y - half, hs, hs),                       ' TopLeft
            New Rectangle(r.X + r.Width \ 2 - half, r.Y - half, hs, hs),         ' TopCenter
            New Rectangle(r.Right - half, r.Y - half, hs, hs),                   ' TopRight
            New Rectangle(r.X - half, r.Y + r.Height \ 2 - half, hs, hs),        ' MiddleLeft
            New Rectangle(r.Right - half, r.Y + r.Height \ 2 - half, hs, hs),    ' MiddleRight
            New Rectangle(r.X - half, r.Bottom - half, hs, hs),                  ' BottomLeft
            New Rectangle(r.X + r.Width \ 2 - half, r.Bottom - half, hs, hs),    ' BottomCenter
            New Rectangle(r.Right - half, r.Bottom - half, hs, hs)               ' BottomRight
        }
    End Function

    ' =========================================================================
    ' Mouse interaction
    ' =========================================================================
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        Me.Focus()
        If e.Button <> MouseButtons.Left Then Return
        Me.Capture = True

        Dim canvasPoint = ScreenToCanvas(e.Location)

        ' If we're placing a new gadget from toolbox
        If _pendingGadgetType.HasValue Then
            _isDrawing = True
            _drawStartPoint = e.Location
            _drawCurrentPoint = e.Location
            Return
        End If

        ' Check if clicking on a resize handle of selected gadget
        If _selectedGadget IsNot Nothing AndAlso Not _selectedGadget.IsLocked Then
            _resizeHandle = HitTestHandles(_selectedGadget, e.Location)
            If _resizeHandle <> ResizeHandleType.None Then
                _isResizing = True
                _dragStart = e.Location
                PushUndo()
                Return
            End If
        End If

        ' Hit test gadgets
        Dim hit = _formDesign.HitTest(canvasPoint)
        If hit IsNot Nothing Then
            ' Ctrl+Click for multi-select
            If (Control.ModifierKeys And Keys.Control) = Keys.Control Then
                If _multiSelection.Contains(hit) Then
                    _multiSelection.Remove(hit)
                    hit.IsSelected = False
                Else
                    _multiSelection.Add(hit)
                    hit.IsSelected = True
                End If
            Else
                _multiSelection.Clear()
                _multiSelection.Add(hit)
            End If

            SelectedGadget = hit
            ' Only allow dragging unlocked gadgets
            If Not hit.IsLocked Then
                _isDragging = True
                _dragStart = e.Location
                _dragOffset = New Point(e.X - CInt(hit.X * _zoom), e.Y - CInt(hit.Y * _zoom))
                PushUndo()
            End If
        Else
            ' Click on empty form surface
            SelectedGadget = Nothing
            _multiSelection.Clear()
            RaiseEvent FormSurfaceClicked()
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)

        If _isDrawing Then
            _drawCurrentPoint = e.Location
            Invalidate()
            Return
        End If

        If _isDragging AndAlso _selectedGadget IsNot Nothing Then
            Dim newX = CInt((e.X - _dragOffset.X) / _zoom)
            Dim newY = CInt((e.Y - _dragOffset.Y) / _zoom)

            ' Move delta for multi-selection (smooth — snap only on mouse-up)
            Dim dx = newX - _selectedGadget.X
            Dim dy = newY - _selectedGadget.Y

            For Each sel In _multiSelection
                sel.X += dx
                sel.Y += dy
                ' Clamp to form bounds
                sel.X = Math.Max(0, sel.X)
                sel.Y = Math.Max(0, sel.Y)
            Next

            Invalidate()
            RaiseEvent GadgetMoved(_selectedGadget)
            RaiseEvent DesignChanged()
            Return
        End If

        If _isResizing AndAlso _selectedGadget IsNot Nothing Then
            ApplyResize(e.Location)
            Invalidate()
            RaiseEvent GadgetResized(_selectedGadget)
            RaiseEvent DesignChanged()
            Return
        End If

        ' Update cursor based on hover
        If _selectedGadget IsNot Nothing Then
            Dim handle = HitTestHandles(_selectedGadget, e.Location)
            Me.Cursor = GetResizeCursor(handle)
        ElseIf _pendingGadgetType.HasValue Then
            Me.Cursor = Cursors.Cross
        Else
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        Me.Capture = False

        If _isDrawing AndAlso _pendingGadgetType.HasValue Then
            _isDrawing = False
            Dim drawRect = GetNormalizedRect(_drawStartPoint, _drawCurrentPoint)
            Dim canvasRect = ScreenToCanvasRect(drawRect)

            ' Enforce minimum size
            If canvasRect.Width < 10 Then canvasRect.Width = W9GadgetRegistry.GetTypeDef(_pendingGadgetType.Value).DefaultWidth
            If canvasRect.Height < 10 Then canvasRect.Height = W9GadgetRegistry.GetTypeDef(_pendingGadgetType.Value).DefaultHeight

            If _snapToGrid Then
                canvasRect.X = SnapToGridValue(canvasRect.X)
                canvasRect.Y = SnapToGridValue(canvasRect.Y)
                canvasRect.Width = SnapToGridValue(canvasRect.Width)
                canvasRect.Height = SnapToGridValue(canvasRect.Height)
            End If

            ' Ensure gadgets are not placed at negative coordinates
            If canvasRect.Y < 0 Then canvasRect.Y = 0

            ' Create the new gadget
            AddGadgetFromToolbox(_pendingGadgetType.Value, canvasRect)
            ClearPendingGadgetType()
            Return
        End If

        _isDragging = False
        _isResizing = False
        _resizeHandle = ResizeHandleType.None

        ' Snap to grid on release (not during drag — keeps movement smooth)
        If _snapToGrid Then
            If _selectedGadget IsNot Nothing Then
                For Each sel In _multiSelection
                    sel.X = SnapToGridValue(sel.X)
                    sel.Y = SnapToGridValue(sel.Y)
                    sel.W = SnapToGridValue(Math.Max(10, sel.W))
                    sel.H = SnapToGridValue(Math.Max(10, sel.H))
                Next
                If Not _multiSelection.Contains(_selectedGadget) Then
                    _selectedGadget.X = SnapToGridValue(_selectedGadget.X)
                    _selectedGadget.Y = SnapToGridValue(_selectedGadget.Y)
                    _selectedGadget.W = SnapToGridValue(Math.Max(10, _selectedGadget.W))
                    _selectedGadget.H = SnapToGridValue(Math.Max(10, _selectedGadget.H))
                End If
                Invalidate()
            End If
        End If
    End Sub

    ' =========================================================================
    ' Keyboard
    ' =========================================================================
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)

        If _selectedGadget Is Nothing Then Return

        Select Case e.KeyCode
            Case Keys.Delete
                DeleteSelectedGadgets()
                e.Handled = True

            Case Keys.Up
                PushUndo()
                Dim step1 = If(_snapToGrid, _gridSize, 1)
                For Each sel In _multiSelection
                    If Not sel.IsLocked Then sel.Y = Math.Max(0, sel.Y - step1)
                Next
                Invalidate()
                RaiseEvent DesignChanged()
                e.Handled = True

            Case Keys.Down
                PushUndo()
                Dim step2 = If(_snapToGrid, _gridSize, 1)
                For Each sel In _multiSelection
                    If Not sel.IsLocked Then sel.Y += step2
                Next
                Invalidate()
                RaiseEvent DesignChanged()
                e.Handled = True

            Case Keys.Left
                PushUndo()
                Dim step3 = If(_snapToGrid, _gridSize, 1)
                For Each sel In _multiSelection
                    If Not sel.IsLocked Then sel.X = Math.Max(0, sel.X - step3)
                Next
                Invalidate()
                RaiseEvent DesignChanged()
                e.Handled = True

            Case Keys.Right
                PushUndo()
                Dim step4 = If(_snapToGrid, _gridSize, 1)
                For Each sel In _multiSelection
                    If Not sel.IsLocked Then sel.X += step4
                Next
                Invalidate()
                RaiseEvent DesignChanged()
                e.Handled = True

            Case Keys.Escape
                ClearPendingGadgetType()
                SelectedGadget = Nothing
                _multiSelection.Clear()
                Invalidate()
                e.Handled = True
        End Select

        ' Ctrl+C / Ctrl+V / Ctrl+X / Ctrl+D / Ctrl+Z / Ctrl+Y / Ctrl+A
        If e.Control Then
            Select Case e.KeyCode
                Case Keys.C
                    CopySelected()
                    e.Handled = True

                Case Keys.V
                    PasteFromClipboard()
                    e.Handled = True

                Case Keys.X
                    CutSelected()
                    e.Handled = True

                Case Keys.D
                    DuplicateSelected()
                    e.Handled = True

                Case Keys.Z
                    PopUndo()
                    e.Handled = True

                Case Keys.Y
                    PopRedo()
                    e.Handled = True

                Case Keys.A
                    SelectAllGadgets()
                    e.Handled = True
            End Select
        End If
    End Sub

    Protected Overrides Function IsInputKey(keyData As Keys) As Boolean
        Select Case keyData
            Case Keys.Up, Keys.Down, Keys.Left, Keys.Right
                Return True
            Case Else
                Return MyBase.IsInputKey(keyData)
        End Select
    End Function

    ' =========================================================================
    ' Public methods
    ' =========================================================================

    ''' <summary>Add a new gadget at the specified position.</summary>
    Public Sub AddGadgetFromToolbox(gadgetType As W9GadgetType, rect As Rectangle)
        PushUndo()
        Dim tdef = W9GadgetRegistry.GetTypeDef(gadgetType)
        Dim idx = _formDesign.Gadgets.Count + 1

        ' Choose default font based on gadget type
        Dim defaultFontName = "Consolas"
        Dim defaultFontSize = 11
        Select Case gadgetType
            Case W9GadgetType.Editor, W9GadgetType.StringInput
                defaultFontName = "Consolas"
                defaultFontSize = 11
            Case W9GadgetType.Button
                defaultFontName = "Segoe UI"
                defaultFontSize = 11
            Case W9GadgetType.TextLabel
                defaultFontName = "Segoe UI"
                defaultFontSize = 11
            Case Else
                defaultFontName = "Segoe UI"
                defaultFontSize = 10
        End Select

        Dim gad As New W9GadgetInstance() With {
            .ID = _formDesign.GetNextGadgetID(),
            .GadgetType = gadgetType,
            .EnumName = W9GadgetRegistry.GenerateEnumName(gadgetType, idx),
            .Text = If(tdef IsNot Nothing, tdef.DefaultText, ""),
            .X = rect.X,
            .Y = rect.Y,
            .W = rect.Width,
            .H = rect.Height,
            .ZOrder = idx,
            .FontName = defaultFontName,
            .FontSize = defaultFontSize
        }
        _formDesign.Gadgets.Add(gad)
        SelectedGadget = gad
        RaiseEvent GadgetAdded(gad)
        RaiseEvent DesignChanged()
        Invalidate()
    End Sub

    ''' <summary>Delete all selected gadgets.</summary>
    Public Sub DeleteSelectedGadgets()
        If _multiSelection.Count = 0 AndAlso _selectedGadget Is Nothing Then Return
        PushUndo()
        If _multiSelection.Count > 0 Then
            For Each sel In _multiSelection
                _formDesign.Gadgets.Remove(sel)
                RaiseEvent GadgetDeleted(sel)
            Next
            _multiSelection.Clear()
        ElseIf _selectedGadget IsNot Nothing Then
            _formDesign.Gadgets.Remove(_selectedGadget)
            RaiseEvent GadgetDeleted(_selectedGadget)
        End If
        _selectedGadget = Nothing
        RaiseEvent DesignChanged()
        Invalidate()
    End Sub

    ''' <summary>Refresh after property changes.</summary>
    Public Sub RefreshDesign()
        Invalidate()
        RaiseEvent DesignChanged()
    End Sub

    ' =========================================================================
    ' Undo/Redo
    ' =========================================================================
    Private Sub PushUndo()
        Dim snapshot = _formDesign.Gadgets.Select(Function(g) g.Clone()).ToList()
        _undoStack.Push(snapshot)
        _redoStack.Clear()
        If _undoStack.Count > 50 Then
            ' Trim stack
        End If
    End Sub

    Private Sub PopUndo()
        If _undoStack.Count = 0 Then Return
        ' Save current for redo
        _redoStack.Push(_formDesign.Gadgets.Select(Function(g) g.Clone()).ToList())
        _formDesign.Gadgets = _undoStack.Pop().ToList()
        _selectedGadget = Nothing
        _multiSelection.Clear()
        Invalidate()
        RaiseEvent DesignChanged()
    End Sub

    Private Sub PopRedo()
        If _redoStack.Count = 0 Then Return
        _undoStack.Push(_formDesign.Gadgets.Select(Function(g) g.Clone()).ToList())
        _formDesign.Gadgets = _redoStack.Pop().ToList()
        _selectedGadget = Nothing
        _multiSelection.Clear()
        Invalidate()
        RaiseEvent DesignChanged()
    End Sub

    ' =========================================================================
    ' Helpers
    ' =========================================================================
    Private Function SnapToGridValue(v As Integer) As Integer
        Return CInt(Math.Round(v / _gridSize) * _gridSize)
    End Function

    Private Function ScreenToCanvas(pt As Point) As Point
        Return New Point(CInt(pt.X / _zoom), CInt(pt.Y / _zoom))
    End Function

    Private Function ScreenToCanvasRect(r As Rectangle) As Rectangle
        Return New Rectangle(CInt(r.X / _zoom), CInt(r.Y / _zoom),
                             CInt(r.Width / _zoom), CInt(r.Height / _zoom))
    End Function

    Private Function GetNormalizedRect(p1 As Point, p2 As Point) As Rectangle
        Return New Rectangle(
            Math.Min(p1.X, p2.X), Math.Min(p1.Y, p2.Y),
            Math.Abs(p2.X - p1.X), Math.Abs(p2.Y - p1.Y))
    End Function

    Private Function HitTestHandles(gad As W9GadgetInstance, pt As Point) As ResizeHandleType
        Dim r As New Rectangle(CInt(gad.X * _zoom), CInt(gad.Y * _zoom),
                               CInt(gad.W * _zoom), CInt(gad.H * _zoom))
        Dim handleRects = GetHandleRects(r)
        Dim types = New ResizeHandleType() {ResizeHandleType.TopLeft, ResizeHandleType.TopCenter, ResizeHandleType.TopRight,
                     ResizeHandleType.MiddleLeft, ResizeHandleType.MiddleRight,
                     ResizeHandleType.BottomLeft, ResizeHandleType.BottomCenter, ResizeHandleType.BottomRight}
        For i = 0 To handleRects.Length - 1
            Dim inflated = Rectangle.Inflate(handleRects(i), 4, 4)
            If inflated.Contains(pt) Then Return types(i)
        Next
        Return ResizeHandleType.None
    End Function

    Private Function GetResizeCursor(handle As ResizeHandleType) As Cursor
        Select Case handle
            Case ResizeHandleType.TopLeft, ResizeHandleType.BottomRight : Return Cursors.SizeNWSE
            Case ResizeHandleType.TopRight, ResizeHandleType.BottomLeft : Return Cursors.SizeNESW
            Case ResizeHandleType.TopCenter, ResizeHandleType.BottomCenter : Return Cursors.SizeNS
            Case ResizeHandleType.MiddleLeft, ResizeHandleType.MiddleRight : Return Cursors.SizeWE
            Case Else
                If _pendingGadgetType.HasValue Then Return Cursors.Cross
                Return Cursors.Default
        End Select
    End Function

    Private Sub ApplyResize(mousePos As Point)
        If _selectedGadget Is Nothing Then Return
        Dim dx = CInt((mousePos.X - _dragStart.X) / _zoom)
        Dim dy = CInt((mousePos.Y - _dragStart.Y) / _zoom)
        _dragStart = mousePos

        Select Case _resizeHandle
            Case ResizeHandleType.TopLeft
                _selectedGadget.X += dx : _selectedGadget.Y += dy
                _selectedGadget.W -= dx : _selectedGadget.H -= dy
            Case ResizeHandleType.TopCenter
                _selectedGadget.Y += dy : _selectedGadget.H -= dy
            Case ResizeHandleType.TopRight
                _selectedGadget.W += dx : _selectedGadget.Y += dy : _selectedGadget.H -= dy
            Case ResizeHandleType.MiddleLeft
                _selectedGadget.X += dx : _selectedGadget.W -= dx
            Case ResizeHandleType.MiddleRight
                _selectedGadget.W += dx
            Case ResizeHandleType.BottomLeft
                _selectedGadget.X += dx : _selectedGadget.W -= dx : _selectedGadget.H += dy
            Case ResizeHandleType.BottomCenter
                _selectedGadget.H += dy
            Case ResizeHandleType.BottomRight
                _selectedGadget.W += dx : _selectedGadget.H += dy
        End Select

        ' Enforce minimums
        If _selectedGadget.W < 10 Then _selectedGadget.W = 10
        If _selectedGadget.H < 10 Then _selectedGadget.H = 10
    End Sub

    ' =========================================================================
    ' Right-Click Context Menu
    ' =========================================================================
    Private _contextMenu As ContextMenuStrip

    Private Sub BuildContextMenu()
        _contextMenu = New ContextMenuStrip()

        Dim mnuCut = _contextMenu.Items.Add("Cut", Nothing, Sub(s, e) CutSelected()) : mnuCut.Name = "mnuCut"
        Dim mnuCopy = _contextMenu.Items.Add("Copy", Nothing, Sub(s, e) CopySelected()) : mnuCopy.Name = "mnuCopy"
        Dim mnuPaste = _contextMenu.Items.Add("Paste", Nothing, Sub(s, e) PasteFromClipboard()) : mnuPaste.Name = "mnuPaste"
        Dim mnuDuplicate = _contextMenu.Items.Add("Duplicate", Nothing, Sub(s, e) DuplicateSelected()) : mnuDuplicate.Name = "mnuDuplicate"
        _contextMenu.Items.Add(New ToolStripSeparator())
        Dim mnuDelete = _contextMenu.Items.Add("Delete", Nothing, Sub(s, e) DeleteSelectedGadgets()) : mnuDelete.Name = "mnuDelete"
        _contextMenu.Items.Add(New ToolStripSeparator())

        ' Z-order
        Dim mnuBringFront = _contextMenu.Items.Add("Bring to Front", Nothing, Sub(s, e) BringToFront_Gadget()) : mnuBringFront.Name = "mnuBringFront"
        Dim mnuSendBack = _contextMenu.Items.Add("Send to Back", Nothing, Sub(s, e) SendToBack_Gadget()) : mnuSendBack.Name = "mnuSendBack"
        _contextMenu.Items.Add(New ToolStripSeparator())

        ' Lock
        Dim mnuLock = _contextMenu.Items.Add("Lock Position", Nothing, Sub(s, e) ToggleLock()) : mnuLock.Name = "mnuLock"
        _contextMenu.Items.Add(New ToolStripSeparator())

        ' Alignment submenu
        Dim mnuAlign = New ToolStripMenuItem("Align")
        mnuAlign.DropDownItems.Add("Align Left", Nothing, Sub(s, e) AlignSelected(AlignDirection.Left))
        mnuAlign.DropDownItems.Add("Align Right", Nothing, Sub(s, e) AlignSelected(AlignDirection.Right))
        mnuAlign.DropDownItems.Add("Align Top", Nothing, Sub(s, e) AlignSelected(AlignDirection.Top))
        mnuAlign.DropDownItems.Add("Align Bottom", Nothing, Sub(s, e) AlignSelected(AlignDirection.Bottom))
        mnuAlign.DropDownItems.Add(New ToolStripSeparator())
        mnuAlign.DropDownItems.Add("Center Horizontally", Nothing, Sub(s, e) AlignSelected(AlignDirection.CenterH))
        mnuAlign.DropDownItems.Add("Center Vertically", Nothing, Sub(s, e) AlignSelected(AlignDirection.CenterV))
        _contextMenu.Items.Add(mnuAlign)

        ' Size submenu
        Dim mnuSize = New ToolStripMenuItem("Make Same Size")
        mnuSize.DropDownItems.Add("Same Width", Nothing, Sub(s, e) SizeSelected(SizeDirection.Width))
        mnuSize.DropDownItems.Add("Same Height", Nothing, Sub(s, e) SizeSelected(SizeDirection.Height))
        mnuSize.DropDownItems.Add("Same Both", Nothing, Sub(s, e) SizeSelected(SizeDirection.Both))
        _contextMenu.Items.Add(mnuSize)

        ' Spacing submenu
        Dim mnuSpace = New ToolStripMenuItem("Space Evenly")
        mnuSpace.DropDownItems.Add("Horizontal Spacing", Nothing, Sub(s, e) SpaceEvenly(True))
        mnuSpace.DropDownItems.Add("Vertical Spacing", Nothing, Sub(s, e) SpaceEvenly(False))
        _contextMenu.Items.Add(mnuSpace)

        _contextMenu.Items.Add(New ToolStripSeparator())

        ' Tab order
        Dim mnuTabOrder = _contextMenu.Items.Add("Show Tab Order", Nothing, Sub(s, e) ToggleTabOrderDisplay()) : mnuTabOrder.Name = "mnuTabOrder"

        ' Select All
        _contextMenu.Items.Add("Select All", Nothing, Sub(s, e) SelectAllGadgets())

        Me.ContextMenuStrip = _contextMenu

        ' Update enabled states when opening
        AddHandler _contextMenu.Opening, Sub(s, e)
                                              Dim hasSelection = _selectedGadget IsNot Nothing
                                              Dim hasMulti = _multiSelection.Count > 1
                                              _contextMenu.Items("mnuCut").Enabled = hasSelection
                                              _contextMenu.Items("mnuCopy").Enabled = hasSelection
                                              _contextMenu.Items("mnuPaste").Enabled = _clipboard IsNot Nothing
                                              _contextMenu.Items("mnuDuplicate").Enabled = hasSelection
                                              _contextMenu.Items("mnuDelete").Enabled = hasSelection
                                              _contextMenu.Items("mnuBringFront").Enabled = hasSelection
                                              _contextMenu.Items("mnuSendBack").Enabled = hasSelection
                                              mnuAlign.Enabled = hasMulti
                                              mnuSize.Enabled = hasMulti
                                              mnuSpace.Enabled = hasMulti AndAlso _multiSelection.Count >= 3
                                              If hasSelection AndAlso _selectedGadget IsNot Nothing Then
                                                  Dim lockItem = _contextMenu.Items("mnuLock")
                                                  lockItem.Text = If(_selectedGadget.IsLocked, "Unlock Position", "Lock Position")
                                              End If
                                          End Sub
    End Sub

    ' =========================================================================
    ' Copy / Cut / Paste / Duplicate
    ' =========================================================================
    Public Sub CopySelected()
        If _selectedGadget IsNot Nothing Then
            _clipboard = _selectedGadget.Clone()
        End If
    End Sub

    Public Sub CutSelected()
        If _selectedGadget IsNot Nothing Then
            _clipboard = _selectedGadget.Clone()
            DeleteSelectedGadgets()
        End If
    End Sub

    Public Sub PasteFromClipboard()
        If _clipboard Is Nothing Then Return
        PushUndo()
        Dim pasted = _clipboard.Clone()
        pasted.X += 20
        pasted.Y += 20
        pasted.ID = _formDesign.GetNextGadgetID()
        Dim idx = _formDesign.Gadgets.Count + 1
        pasted.EnumName = W9GadgetRegistry.GenerateEnumName(pasted.GadgetType, idx)
        _formDesign.Gadgets.Add(pasted)
        SelectedGadget = pasted
        RaiseEvent GadgetAdded(pasted)
        RaiseEvent DesignChanged()
        Invalidate()
    End Sub

    Public Sub DuplicateSelected()
        If _selectedGadget Is Nothing Then Return
        PushUndo()
        Dim duped = _selectedGadget.Clone()
        duped.X += _gridSize * 2
        duped.Y += _gridSize * 2
        duped.ID = _formDesign.GetNextGadgetID()
        Dim idx = _formDesign.Gadgets.Count + 1
        duped.EnumName = W9GadgetRegistry.GenerateEnumName(duped.GadgetType, idx)
        _formDesign.Gadgets.Add(duped)
        SelectedGadget = duped
        RaiseEvent GadgetAdded(duped)
        RaiseEvent DesignChanged()
        Invalidate()
    End Sub

    ' =========================================================================
    ' Z-Order (Bring to Front / Send to Back)
    ' =========================================================================
    Public Sub BringToFront_Gadget()
        If _selectedGadget Is Nothing Then Return
        PushUndo()
        _formDesign.Gadgets.Remove(_selectedGadget)
        _formDesign.Gadgets.Add(_selectedGadget)
        ReassignZOrder()
        RaiseEvent DesignChanged()
        Invalidate()
    End Sub

    Public Sub SendToBack_Gadget()
        If _selectedGadget Is Nothing Then Return
        PushUndo()
        _formDesign.Gadgets.Remove(_selectedGadget)
        _formDesign.Gadgets.Insert(0, _selectedGadget)
        ReassignZOrder()
        RaiseEvent DesignChanged()
        Invalidate()
    End Sub

    Private Sub ReassignZOrder()
        For i = 0 To _formDesign.Gadgets.Count - 1
            _formDesign.Gadgets(i).ZOrder = i + 1
        Next
    End Sub

    Private Sub ToggleLock()
        If _selectedGadget Is Nothing Then Return
        _selectedGadget.IsLocked = Not _selectedGadget.IsLocked
        RaiseEvent DesignChanged()
        Invalidate()
    End Sub

    ' =========================================================================
    ' Alignment Tools
    ' =========================================================================
    Public Enum AlignDirection
        Left
        Right
        Top
        Bottom
        CenterH
        CenterV
    End Enum

    Public Enum SizeDirection
        Width
        Height
        Both
    End Enum

    Public Sub AlignSelected(direction As AlignDirection)
        If _multiSelection.Count < 2 OrElse _selectedGadget Is Nothing Then Return
        PushUndo()

        ' Use the primary selected gadget as anchor
        Dim anchor = _selectedGadget

        For Each g In _multiSelection
            If g Is anchor Then Continue For
            Select Case direction
                Case AlignDirection.Left : g.X = anchor.X
                Case AlignDirection.Right : g.X = anchor.X + anchor.W - g.W
                Case AlignDirection.Top : g.Y = anchor.Y
                Case AlignDirection.Bottom : g.Y = anchor.Y + anchor.H - g.H
                Case AlignDirection.CenterH : g.X = anchor.X + (anchor.W - g.W) \ 2
                Case AlignDirection.CenterV : g.Y = anchor.Y + (anchor.H - g.H) \ 2
            End Select
        Next

        RaiseEvent DesignChanged()
        Invalidate()
    End Sub

    Public Sub SizeSelected(direction As SizeDirection)
        If _multiSelection.Count < 2 OrElse _selectedGadget Is Nothing Then Return
        PushUndo()

        Dim anchor = _selectedGadget

        For Each g In _multiSelection
            If g Is anchor Then Continue For
            Select Case direction
                Case SizeDirection.Width : g.W = anchor.W
                Case SizeDirection.Height : g.H = anchor.H
                Case SizeDirection.Both : g.W = anchor.W : g.H = anchor.H
            End Select
        Next

        RaiseEvent DesignChanged()
        Invalidate()
    End Sub

    Public Sub SpaceEvenly(horizontal As Boolean)
        If _multiSelection.Count < 3 Then Return
        PushUndo()

        If horizontal Then
            ' Sort by X position
            Dim sorted = _multiSelection.OrderBy(Function(g) g.X).ToList()
            Dim totalW = sorted.Sum(Function(g) g.W)
            Dim span = (sorted.Last().X + sorted.Last().W) - sorted.First().X
            Dim gap = (span - totalW) / (_multiSelection.Count - 1)
            Dim curX = sorted.First().X
            For Each g In sorted
                g.X = CInt(curX)
                curX += g.W + gap
            Next
        Else
            ' Sort by Y position
            Dim sorted = _multiSelection.OrderBy(Function(g) g.Y).ToList()
            Dim totalH = sorted.Sum(Function(g) g.H)
            Dim span = (sorted.Last().Y + sorted.Last().H) - sorted.First().Y
            Dim gap = (span - totalH) / (_multiSelection.Count - 1)
            Dim curY = sorted.First().Y
            For Each g In sorted
                g.Y = CInt(curY)
                curY += g.H + gap
            Next
        End If

        RaiseEvent DesignChanged()
        Invalidate()
    End Sub

    ' =========================================================================
    ' Tab Order Display
    ' =========================================================================
    Private _showTabOrder As Boolean = False

    Public Sub ToggleTabOrderDisplay()
        _showTabOrder = Not _showTabOrder
        Invalidate()
    End Sub

    Public ReadOnly Property ShowingTabOrder As Boolean
        Get
            Return _showTabOrder
        End Get
    End Property

    Public Sub SelectAllGadgets()
        _multiSelection.Clear()
        For Each gad In _formDesign.Gadgets
            gad.IsSelected = True
            _multiSelection.Add(gad)
        Next
        If _formDesign.Gadgets.Count > 0 Then
            _selectedGadget = _formDesign.Gadgets.Last()
        End If
        RaiseEvent GadgetSelected(_selectedGadget)
        Invalidate()
    End Sub

End Class
