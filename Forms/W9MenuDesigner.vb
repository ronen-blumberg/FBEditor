Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Menu designer dialog — tree-based editor for creating/editing Window9 menu structures.
''' Allows adding top-level menus, menu items, separators, and reorganizing via drag.
''' </summary>
Public Class W9MenuDesigner
    Inherits Form

    Private _formDesign As W9FormDesign
    Private _treeView As TreeView
    Private _txtEnumName As TextBox
    Private _txtMenuText As TextBox
    Private _chkSeparator As CheckBox
    Private _btnAddTopMenu As Button
    Private _btnAddItem As Button
    Private _btnAddSeparator As Button
    Private _btnDelete As Button
    Private _btnMoveUp As Button
    Private _btnMoveDown As Button
    Private _btnOK As Button
    Private _btnCancel As Button
    Private _previewLabel As Label

    ''' <summary>The edited menu items result.</summary>
    Public ResultMenuItems As New List(Of W9MenuItemInfo)()

    Public Sub New(design As W9FormDesign)
        _formDesign = design
        InitializeComponent()
        LoadMenuTree()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Menu Designer"
        Me.Size = New Size(600, 480)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Font = New Font("Segoe UI", 9)

        ' Tree view (left side)
        _treeView = New TreeView() With {
            .Location = New Point(12, 12), .Size = New Size(260, 300),
            .HideSelection = False, .FullRowSelect = True
        }
        AddHandler _treeView.AfterSelect, AddressOf OnTreeSelect
        Me.Controls.Add(_treeView)

        ' Buttons (right side of tree)
        Dim btnX = 284
        _btnAddTopMenu = New Button() With {.Text = "Add Menu", .Location = New Point(btnX, 12), .Size = New Size(100, 28)}
        _btnAddItem = New Button() With {.Text = "Add Item", .Location = New Point(btnX, 46), .Size = New Size(100, 28)}
        _btnAddSeparator = New Button() With {.Text = "Add Separator", .Location = New Point(btnX, 80), .Size = New Size(100, 28)}
        _btnDelete = New Button() With {.Text = "Delete", .Location = New Point(btnX, 120), .Size = New Size(100, 28)}
        _btnMoveUp = New Button() With {.Text = "Move Up", .Location = New Point(btnX, 160), .Size = New Size(100, 28)}
        _btnMoveDown = New Button() With {.Text = "Move Down", .Location = New Point(btnX, 194), .Size = New Size(100, 28)}
        Me.Controls.AddRange({_btnAddTopMenu, _btnAddItem, _btnAddSeparator, _btnDelete, _btnMoveUp, _btnMoveDown})

        AddHandler _btnAddTopMenu.Click, AddressOf OnAddTopMenu
        AddHandler _btnAddItem.Click, AddressOf OnAddItem
        AddHandler _btnAddSeparator.Click, AddressOf OnAddSeparator
        AddHandler _btnDelete.Click, AddressOf OnDelete
        AddHandler _btnMoveUp.Click, AddressOf OnMoveUp
        AddHandler _btnMoveDown.Click, AddressOf OnMoveDown

        ' Properties panel (bottom)
        Dim propY = 320
        Me.Controls.Add(New Label() With {.Text = "Enum Name:", .Location = New Point(12, propY + 4), .AutoSize = True})
        _txtEnumName = New TextBox() With {.Location = New Point(100, propY), .Size = New Size(170, 24)}
        AddHandler _txtEnumName.TextChanged, AddressOf OnEnumNameChanged
        Me.Controls.Add(_txtEnumName)

        Me.Controls.Add(New Label() With {.Text = "Menu Text:", .Location = New Point(12, propY + 34), .AutoSize = True})
        _txtMenuText = New TextBox() With {.Location = New Point(100, propY + 30), .Size = New Size(170, 24)}
        AddHandler _txtMenuText.TextChanged, AddressOf OnMenuTextChanged
        Me.Controls.Add(_txtMenuText)

        _chkSeparator = New CheckBox() With {.Text = "Is Separator", .Location = New Point(100, propY + 60), .AutoSize = True, .Enabled = False}
        Me.Controls.Add(_chkSeparator)

        ' Preview
        _previewLabel = New Label() With {
            .Text = "Preview: File  Edit  Help",
            .Location = New Point(12, propY + 90),
            .Size = New Size(360, 24),
            .Font = New Font("Segoe UI", 9, FontStyle.Italic),
            .ForeColor = Color.Gray
        }
        Me.Controls.Add(_previewLabel)

        ' OK / Cancel
        _btnOK = New Button() With {.Text = "OK", .DialogResult = DialogResult.OK, .Location = New Point(400, 400), .Size = New Size(80, 30)}
        _btnCancel = New Button() With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Location = New Point(490, 400), .Size = New Size(80, 30)}
        AddHandler _btnOK.Click, AddressOf OnOK
        Me.Controls.AddRange({_btnOK, _btnCancel})
        Me.AcceptButton = _btnOK
        Me.CancelButton = _btnCancel
    End Sub

    ' =========================================================================
    ' Load/Save
    ' =========================================================================
    Private Sub LoadMenuTree()
        _treeView.Nodes.Clear()
        For Each topMenu In _formDesign.MenuItems
            Dim topNode = New TreeNode(topMenu.Text & " [" & topMenu.EnumName & "]") With {.Tag = topMenu}
            For Each child In topMenu.Children
                Dim label = If(child.IsSeparator, "--- (separator)", child.Text & " [" & child.EnumName & "]")
                Dim childNode = New TreeNode(label) With {.Tag = child}
                topNode.Nodes.Add(childNode)
            Next
            _treeView.Nodes.Add(topNode)
        Next
        If _treeView.Nodes.Count > 0 Then
            _treeView.ExpandAll()
            _treeView.SelectedNode = _treeView.Nodes(0)
        End If
        UpdatePreview()
    End Sub

    Private Sub CollectMenuItems()
        ResultMenuItems.Clear()
        For Each topNode As TreeNode In _treeView.Nodes
            Dim topMenu = DirectCast(topNode.Tag, W9MenuItemInfo)
            topMenu.Children.Clear()
            For Each childNode As TreeNode In topNode.Nodes
                topMenu.Children.Add(DirectCast(childNode.Tag, W9MenuItemInfo))
            Next
            ResultMenuItems.Add(topMenu)
        Next
    End Sub

    ' =========================================================================
    ' Events
    ' =========================================================================
    Private Sub OnTreeSelect(sender As Object, e As TreeViewEventArgs)
        If e.Node Is Nothing OrElse e.Node.Tag Is Nothing Then Return
        Dim mi = DirectCast(e.Node.Tag, W9MenuItemInfo)
        _txtEnumName.Text = mi.EnumName
        _txtMenuText.Text = mi.Text
        _chkSeparator.Checked = mi.IsSeparator
    End Sub

    Private Sub OnEnumNameChanged(sender As Object, e As EventArgs)
        If _treeView.SelectedNode Is Nothing Then Return
        Dim mi = DirectCast(_treeView.SelectedNode.Tag, W9MenuItemInfo)
        mi.EnumName = _txtEnumName.Text
        UpdateNodeText(_treeView.SelectedNode, mi)
    End Sub

    Private Sub OnMenuTextChanged(sender As Object, e As EventArgs)
        If _treeView.SelectedNode Is Nothing Then Return
        Dim mi = DirectCast(_treeView.SelectedNode.Tag, W9MenuItemInfo)
        mi.Text = _txtMenuText.Text
        UpdateNodeText(_treeView.SelectedNode, mi)
        UpdatePreview()
    End Sub

    Private Sub UpdateNodeText(node As TreeNode, mi As W9MenuItemInfo)
        If mi.IsSeparator Then
            node.Text = "--- (separator)"
        Else
            node.Text = mi.Text & " [" & mi.EnumName & "]"
        End If
    End Sub

    Private Sub OnAddTopMenu(sender As Object, e As EventArgs)
        Dim idx = _treeView.Nodes.Count + 1
        Dim mi As New W9MenuItemInfo() With {
            .ID = _formDesign.GetNextMenuID(),
            .EnumName = "miMenu" & idx,
            .Text = "&Menu" & idx,
            .IsTopLevel = True
        }
        Dim node = New TreeNode(mi.Text & " [" & mi.EnumName & "]") With {.Tag = mi}
        _treeView.Nodes.Add(node)
        _treeView.SelectedNode = node
        UpdatePreview()
    End Sub

    Private Sub OnAddItem(sender As Object, e As EventArgs)
        Dim parentNode = _treeView.SelectedNode
        ' If child is selected, use its parent
        If parentNode IsNot Nothing AndAlso parentNode.Parent IsNot Nothing Then
            parentNode = parentNode.Parent
        End If
        If parentNode Is Nothing Then
            MessageBox.Show("Select a top-level menu first.", "Menu Designer", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim idx = parentNode.Nodes.Count + 1
        Dim mi As New W9MenuItemInfo() With {
            .ID = _formDesign.GetNextMenuID(),
            .EnumName = "mi" & SanitizeName(parentNode.Text) & "Item" & idx,
            .Text = "&Item " & idx
        }
        Dim node = New TreeNode(mi.Text & " [" & mi.EnumName & "]") With {.Tag = mi}
        parentNode.Nodes.Add(node)
        parentNode.Expand()
        _treeView.SelectedNode = node
    End Sub

    Private Sub OnAddSeparator(sender As Object, e As EventArgs)
        Dim parentNode = _treeView.SelectedNode
        If parentNode IsNot Nothing AndAlso parentNode.Parent IsNot Nothing Then
            parentNode = parentNode.Parent
        End If
        If parentNode Is Nothing Then
            MessageBox.Show("Select a top-level menu first.", "Menu Designer", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim mi As New W9MenuItemInfo() With {
            .ID = _formDesign.GetNextMenuID(),
            .EnumName = "miSep" & (parentNode.Nodes.Count + 1),
            .Text = "-",
            .IsSeparator = True
        }
        Dim node = New TreeNode("--- (separator)") With {.Tag = mi}
        parentNode.Nodes.Add(node)
        parentNode.Expand()
    End Sub

    Private Sub OnDelete(sender As Object, e As EventArgs)
        If _treeView.SelectedNode IsNot Nothing Then
            _treeView.SelectedNode.Remove()
            UpdatePreview()
        End If
    End Sub

    Private Sub OnMoveUp(sender As Object, e As EventArgs)
        MoveNode(-1)
    End Sub

    Private Sub OnMoveDown(sender As Object, e As EventArgs)
        MoveNode(1)
    End Sub

    Private Sub MoveNode(direction As Integer)
        If _treeView.SelectedNode Is Nothing Then Return
        Dim node = _treeView.SelectedNode
        Dim parent = node.Parent
        Dim collection = If(parent IsNot Nothing, parent.Nodes, _treeView.Nodes)
        Dim idx = collection.IndexOf(node)
        Dim newIdx = idx + direction
        If newIdx < 0 OrElse newIdx >= collection.Count Then Return
        collection.Remove(node)
        collection.Insert(newIdx, node)
        _treeView.SelectedNode = node
        UpdatePreview()
    End Sub

    Private Sub OnOK(sender As Object, e As EventArgs)
        CollectMenuItems()
    End Sub

    Private Sub UpdatePreview()
        Dim parts As New List(Of String)()
        For Each node As TreeNode In _treeView.Nodes
            Dim mi = DirectCast(node.Tag, W9MenuItemInfo)
            parts.Add(mi.Text.Replace("&", ""))
        Next
        _previewLabel.Text = "Preview: " & String.Join("  ", parts)
    End Sub

    Private Function SanitizeName(s As String) As String
        Dim result As New System.Text.StringBuilder()
        For Each c In s
            If Char.IsLetterOrDigit(c) Then result.Append(c)
        Next
        Dim r = result.ToString()
        If r.Length > 10 Then r = r.Substring(0, 10)
        Return r
    End Function
End Class
