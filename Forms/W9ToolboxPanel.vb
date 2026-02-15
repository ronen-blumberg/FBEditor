Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Toolbox panel — categorized list of Window9 gadgets the user can select
''' and draw onto the designer canvas. Styled like Visual Studio's Toolbox.
''' </summary>
Public Class W9ToolboxPanel
    Inherits Panel

    Private _listView As ListView
    Private _isDarkTheme As Boolean = False

    ''' <summary>Fired when user selects a gadget type to draw.</summary>
    Public Event GadgetTypeSelected(gadgetType As W9GadgetType)

    ' =========================================================================
    ' Constructor
    ' =========================================================================
    Public Sub New()
        Me.Dock = DockStyle.Fill
        Me.Padding = New Padding(0)

        ' Header
        Dim header As New Label() With {
            .Text = "Toolbox",
            .Dock = DockStyle.Top,
            .Height = 24,
            .TextAlign = ContentAlignment.MiddleLeft,
            .Padding = New Padding(6, 0, 0, 0),
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .BackColor = Color.FromArgb(0, 120, 215),
            .ForeColor = Color.White
        }
        Me.Controls.Add(header)

        ' ListView with gadget types
        _listView = New ListView() With {
            .Dock = DockStyle.Fill,
            .View = View.List,
            .MultiSelect = False,
            .HeaderStyle = ColumnHeaderStyle.None,
            .FullRowSelect = True,
            .ShowGroups = True,
            .HideSelection = False,
            .BorderStyle = BorderStyle.None,
            .Font = New Font("Segoe UI", 9)
        }

        ' Small image list for icons
        Dim imgList As New ImageList() With {
            .ImageSize = New Size(16, 16),
            .ColorDepth = ColorDepth.Depth32Bit
        }
        CreateToolboxIcons(imgList)
        _listView.SmallImageList = imgList

        PopulateToolbox()

        AddHandler _listView.MouseDoubleClick, AddressOf OnToolboxDoubleClick
        AddHandler _listView.ItemActivate, AddressOf OnToolboxItemActivate

        Me.Controls.Add(_listView)
        ' Ensure listview is on top of header
        _listView.BringToFront()
        header.SendToBack()
    End Sub

    ' =========================================================================
    ' Population
    ' =========================================================================
    Private Sub PopulateToolbox()
        _listView.Items.Clear()
        _listView.Groups.Clear()

        ' Create groups
        Dim groupCommon As New ListViewGroup("Common Controls", HorizontalAlignment.Left)
        Dim groupContainers As New ListViewGroup("Containers", HorizontalAlignment.Left)
        Dim groupData As New ListViewGroup("Data Controls", HorizontalAlignment.Left)
        _listView.Groups.AddRange({groupCommon, groupContainers, groupData})

        ' Pointer tool (deselect)
        Dim pointerItem As New ListViewItem("  Pointer (Select)") With {
            .Tag = Nothing,
            .ImageIndex = 0,
            .Group = groupCommon
        }
        _listView.Items.Add(pointerItem)

        ' Add all gadget types
        Dim imgIdx = 1
        For Each tdef In W9GadgetRegistry.AllTypes
            Dim item As New ListViewItem("  " & tdef.DisplayName) With {
                .Tag = tdef.GadgetType,
                .ImageIndex = imgIdx
            }
            Select Case tdef.ToolboxCategory
                Case "Containers" : item.Group = groupContainers
                Case "Data Controls" : item.Group = groupData
                Case Else : item.Group = groupCommon
            End Select
            _listView.Items.Add(item)
            imgIdx += 1
        Next
    End Sub

    ''' <summary>Create simple colored icons for the toolbox.</summary>
    Private Sub CreateToolboxIcons(imgList As ImageList)
        ' Icon 0: Pointer
        imgList.Images.Add(CreateIcon(Color.Gray, "►"))

        ' One icon per gadget type
        For Each tdef In W9GadgetRegistry.AllTypes
            Dim iconColor As Color
            Select Case tdef.ToolboxCategory
                Case "Containers" : iconColor = Color.FromArgb(180, 120, 60)
                Case "Data Controls" : iconColor = Color.FromArgb(60, 120, 180)
                Case Else : iconColor = Color.FromArgb(0, 120, 215)
            End Select
            imgList.Images.Add(CreateIcon(iconColor, tdef.ToolboxIcon))
        Next
    End Sub

    Private Function CreateIcon(bgColor As Color, text As String) As Bitmap
        Dim bmp As New Bitmap(16, 16)
        Using g = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            Using br As New SolidBrush(bgColor)
                g.FillRectangle(br, 1, 1, 14, 14)
            End Using
            g.DrawRectangle(Pens.Gray, 1, 1, 13, 13)
            If Not String.IsNullOrEmpty(text) Then
                Using f As New Font("Segoe UI", 6, FontStyle.Bold)
                    Dim sf As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                    g.DrawString(text, f, Brushes.White, New RectangleF(0, 0, 16, 16), sf)
                End Using
            End If
        End Using
        Return bmp
    End Function

    ' =========================================================================
    ' Events
    ' =========================================================================
    Private Sub OnToolboxDoubleClick(sender As Object, e As MouseEventArgs)
        ActivateSelectedItem()
    End Sub

    Private Sub OnToolboxItemActivate(sender As Object, e As EventArgs)
        ActivateSelectedItem()
    End Sub

    Private Sub ActivateSelectedItem()
        If _listView.SelectedItems.Count = 0 Then Return
        Dim item = _listView.SelectedItems(0)
        If item.Tag Is Nothing Then
            ' Pointer selected — notify to cancel pending
            RaiseEvent GadgetTypeSelected(Nothing)
        Else
            Dim gt = DirectCast(item.Tag, W9GadgetType)
            RaiseEvent GadgetTypeSelected(gt)
        End If
    End Sub

    ''' <summary>Get the currently selected gadget type (Nothing for pointer).</summary>
    Public Function GetSelectedGadgetType() As W9GadgetType?
        If _listView.SelectedItems.Count = 0 Then Return Nothing
        Dim item = _listView.SelectedItems(0)
        If item.Tag Is Nothing Then Return Nothing
        Return DirectCast(item.Tag, W9GadgetType)
    End Function

    ''' <summary>Reset selection to Pointer.</summary>
    Public Sub ResetToPointer()
        If _listView.Items.Count > 0 Then
            _listView.Items(0).Selected = True
        End If
    End Sub

    ' =========================================================================
    ' Theming
    ' =========================================================================
    Public Sub ApplyTheme(isDark As Boolean)
        _isDarkTheme = isDark
        If isDark Then
            Me.BackColor = Color.FromArgb(37, 37, 38)
            _listView.BackColor = Color.FromArgb(37, 37, 38)
            _listView.ForeColor = Color.FromArgb(220, 220, 220)
        Else
            Me.BackColor = Color.FromArgb(246, 246, 246)
            _listView.BackColor = Color.FromArgb(246, 246, 246)
            _listView.ForeColor = Color.FromArgb(30, 30, 30)
        End If
    End Sub
End Class
