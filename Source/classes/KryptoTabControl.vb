'SEE BUILD GUIDE BEFORE USING. 
'THIS CLASS IS NOT COMPLETE, WILL BE UPDATED. 
Imports Gecko

Public Class KryptoTabControlPlus
    Inherits TabControl

#Region "     constants"

    'Windows Constants used in WndProc 
    Private Const WM_NULL As Integer = &H0
    Private Const WM_MOUSEDOWN As Integer = &H201

#End Region

#Region "     class level variables"

    'tracking variables used for highlighting tabs, remove and add buttons 
    Private HotTabIndex As Integer = -1
    Private onCloseButton As Boolean = False
    Private onAddButton As Boolean = False

#End Region

#Region "     property"

    'used for toggling the add button 
    Private _AllowUsersToAddTabPages As Boolean
    Public Property AllowUsersToAddTabPages() As Boolean
        Get
            Return _AllowUsersToAddTabPages
        End Get
        Set(ByVal value As Boolean)
            _AllowUsersToAddTabPages = value
            Me.Invalidate()
        End Set
    End Property

#End Region

#Region "     constructor"

    'creates a new extendedTabControl 
    Public Sub New()
        MyBase.New()
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint, True)
    End Sub

#End Region

#Region "     overridden events"

    Protected Overrides Sub OnControlAdded(e As ControlEventArgs)
        MyBase.OnControlAdded(e)
        If TypeOf e.Control IsNot DoubleBufferedTabPage Then
            Dim t As String = MyBase.TabPages(MyBase.TabPages.IndexOf(DirectCast(e.Control, TabPage))).Text
            MyBase.TabPages.Remove(DirectCast(e.Control, TabPage))
            MyBase.TabPages.Add(New DoubleBufferedTabPage With {.Text = t})
        End If
    End Sub

    ''' <summary> 
    ''' Overridden OnCreateControl method 
    ''' Adds functionality for setting up dynamic event handlers 
    ''' </summary> 
    Protected Overrides Sub OnCreateControl()
        MyBase.OnCreateControl()
        For Each tp As TabPage In MyBase.TabPages
            AddHandler tp.Paint, AddressOf tabpages_paint
            AddHandler tp.MouseMove, AddressOf tabpages_mousemove
        Next
        AddHandler MyBase.Parent.MouseMove, AddressOf me_mousemove
        AddHandler MyBase.Parent.MouseDown, AddressOf parent_mousedown
    End Sub

    ''' <summary> 
    ''' extendedTabControl Parent MouseDown handler 
    ''' Used to capture and respond to Add button clicks 
    ''' </summary> 
    ''' <param name="sender"></param> 
    ''' <param name="e"></param> 
    Private Sub parent_mousedown(sender As Object, e As MouseEventArgs)
        If AllowUsersToAddTabPages Then
            onAddButton = PointerOnAddButton()
            If onAddButton AndAlso e.Button = MouseButtons.Left Then
                If Not Me.DesignMode Then
                    Dim title As String = "New Tab"
                    Dim tp As New DoubleBufferedTabPage With {.Text = title, .BackColor = SystemColors.Window}
                    MyBase.TabPages.Add(tp)
                    AddHandler tp.Paint, AddressOf tabpages_paint
                    AddHandler tp.MouseMove, AddressOf tabpages_mousemove
                    MyBase.SelectedTab = tp

                End If
            End If
        End If
    End Sub

    ''' <summary> 
    ''' Overridden OnMouseMove event 
    ''' Passes MouseEventArgs to me_mousemove which handles TabControl and parent (Form) mousemove events 
    ''' </summary> 
    ''' <param name="e"></param> 
    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseMove(e)
        me_mousemove(Me, e)
    End Sub

    ''' <summary> 
    ''' TabControl and parent (Form) mousemove event handler 
    ''' Sets variables used in painting tabs 
    ''' </summary> 
    ''' <param name="sender"></param> 
    ''' <param name="e"></param> 
    Private Sub me_mousemove(sender As Object, e As MouseEventArgs)

        If MyBase.Bounds.Contains(e.Location) Then
            HotTabIndex = getTabIndex(e.Location)
            If HotTabIndex < MyBase.TabPages.Count Then
                onCloseButton = PointerOnCloseButton(HotTabIndex)
                onAddButton = False
            Else
                onCloseButton = False
                onAddButton = PointerOnAddButton()
            End If

            Me.Invalidate()
        End If
    End Sub

    ''' <summary> 
    ''' TabPages mousemove handler 
    ''' Clears variables used in painting tabs  
    ''' </summary> 
    ''' <param name="sender"></param> 
    ''' <param name="e"></param> 
    Private Sub tabpages_mousemove(sender As Object, e As MouseEventArgs)
        HotTabIndex = -1
        onCloseButton = False
        onAddButton = False
        Me.Invalidate()
    End Sub

    ''' <summary> 
    ''' Overridden OnPaint event 
    ''' Draws Tab backgrounds, title texts, and buttons 
    ''' </summary> 
    ''' <param name="e"></param> 
    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)
        For id As Integer = 0 To MyBase.TabPages.Count - If(AllowUsersToAddTabPages, 0, 1)
            DrawTabContent(e.Graphics, id)
        Next
    End Sub

    ''' <summary> 
    ''' Overridden OnPaintBackground event 
    ''' Draws Tab borders 
    ''' </summary> 
    ''' <param name="pevent"></param> 
    Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaintBackground(pevent)
        For id As Integer = 0 To MyBase.TabPages.Count - If(AllowUsersToAddTabPages, 0, 1)
            DrawTabBackground(pevent.Graphics, id)
        Next
    End Sub

    ''' <summary> 
    ''' TabPages Paint event 
    ''' Draws borders around active tab 
    ''' </summary> 
    ''' <param name="sender"></param> 
    ''' <param name="e"></param> 
    Private Sub tabpages_paint(sender As Object, e As PaintEventArgs)
        If MyBase.SelectedIndex = -1 Then Return
        Dim tabrect As Rectangle = MyBase.GetTabRect(MyBase.SelectedIndex)
        Dim tp As TabPage = DirectCast(sender, TabPage)
        If tp Is MyBase.TabPages(0) Then
            tabrect.X -= 2
        Else
            tabrect.X -= 4
        End If
        tabrect.Width += 3
        Dim r As New Rectangle(0, 0, tp.Bounds.Width - 1, tp.Bounds.Height - 1)

        e.Graphics.DrawLine(Pens.Black, r.Left, r.Top, r.Left, r.Bottom)
        e.Graphics.DrawLine(Pens.Black, r.Left, r.Bottom, r.Right, r.Bottom)
        e.Graphics.DrawLine(Pens.Black, r.Right, r.Top, r.Right, r.Bottom)
        e.Graphics.DrawLine(Pens.Black, r.Left, r.Top, tabrect.Left, r.Top)
        e.Graphics.DrawLine(Pens.Black, tabrect.Right - 4, r.Top, r.Right, r.Top)
    End Sub

    ''' <summary> 
    ''' Overridden OnSelectedIndexChanged event 
    ''' Ensures TabControl painting is processed 
    ''' </summary> 
    ''' <param name="e"></param> 
    Protected Overrides Sub OnSelectedIndexChanged(e As EventArgs)
        MyBase.OnSelectedIndexChanged(e)
        If MyBase.SelectedIndex > -1 Then
            MyBase.SelectedTab.Invalidate()
        End If
        MyBase.Invalidate()
    End Sub

    ''' <summary> 
    ''' Overridden WndProc 
    ''' Reacts to close button clicks 
    ''' </summary> 
    ''' <param name="m"></param> 
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        'If MyBase.TabPages.Count = 0 Then Me.FindForm.Close()

        If m.Msg = WM_MOUSEDOWN Then

            If onCloseButton AndAlso Not Me.DesignMode Then


                MyBase.TabPages(HotTabIndex).Controls(0).Dispose()
                MyBase.TabPages.Remove(MyBase.TabPages(HotTabIndex))




                If MyBase.TabPages.Count = 0 Then Me.FindForm.Close()
                m.Msg = WM_NULL
            End If
        End If
        MyBase.WndProc(m)
    End Sub

#End Region

#Region "     helper drawing methods"

    ''' <summary> 
    ''' Helper method for OnPaintBackground event 
    ''' </summary> 
    ''' <param name="graphics"></param> 
    ''' <param name="id"></param> 
    Private Sub DrawTabBackground(ByVal graphics As Graphics, ByVal id As Integer)
        If id < MyBase.TabPages.Count Then
            Dim tp As TabPage = MyBase.TabPages(id)
            If id = SelectedIndex Then
                Dim r As Rectangle = GetTabRect(id)
                If id = 0 Then r.X += 2 : r.Width -= 2
                graphics.DrawLine(Pens.Black, r.Left, r.Top, r.Left, r.Bottom)
                graphics.DrawLine(Pens.Black, r.Left, r.Top, r.Right, r.Top)
                graphics.DrawLine(Pens.Black, r.Right, r.Top, r.Right, r.Bottom)
            Else
                Dim r As Rectangle = GetTabRect(id)
                If id = 0 Then r.X += 2 : r.Width -= 2
                r.Y += 3
                graphics.DrawLine(Pens.Black, r.Left, r.Top, r.Left, r.Bottom)
                graphics.DrawLine(Pens.Black, r.Left, r.Top, r.Right, r.Top)
                graphics.DrawLine(Pens.Black, r.Right, r.Top, r.Right, r.Bottom)
            End If
        Else
            Dim points As Point() = getPolygon()
            points(points.GetUpperBound(0) - 1) = points(points.GetUpperBound(0))
            ReDim Preserve points(points.GetUpperBound(0) - 1)
            For x As Integer = 0 To points.GetUpperBound(0)
                points(x).Y -= 1
            Next
            graphics.DrawLines(Pens.Black, points)
        End If
    End Sub

    ''' <summary> 
    ''' Helper method for OnPaint event 
    ''' </summary> 
    ''' <param name="graphics"></param> 
    ''' <param name="id"></param> 
    Private Sub DrawTabContent(ByVal graphics As Graphics, ByVal id As Integer)

        If id < MyBase.TabPages.Count Then
            Dim tp As TabPage = MyBase.TabPages(id)
            Dim xFont As New Font("GulimChe", 10, FontStyle.Bold)
            Dim tabRect As Rectangle = GetTabRect(id)
            tabRect.Height += 2
            If id = 0 Then tabRect.X += 2 : tabRect.Width -= 2
            Dim textrect As New Rectangle(Point.Empty, tabRect.Size)
            textrect.X += 2
            If id <> MyBase.SelectedIndex Then textrect.Y += 2

            Dim sf As New StringFormat
            sf.Alignment = StringAlignment.Near
            sf.LineAlignment = StringAlignment.Center

            Using bm As New Bitmap(tabRect.Width, tabRect.Height)
                Using bmGraphics As Graphics = Graphics.FromImage(bm)

                    If id = MyBase.SelectedIndex Then
                        bmGraphics.FillRectangle(SystemBrushes.Window, New Rectangle(1, 1, tabRect.Width - 2, tabRect.Height))
                    ElseIf id <> MyBase.SelectedIndex And id = HotTabIndex Then
                        bmGraphics.FillRectangle(SystemBrushes.GradientActiveCaption, New Rectangle(1, 4, tabRect.Width - 1, tabRect.Height))
                    Else
                        bmGraphics.FillRectangle(New SolidBrush(Me.Parent.BackColor), New Rectangle(1, 4, tabRect.Width - 1, tabRect.Height))
                    End If
                    bmGraphics.DrawString(Me.TabPages(id).Text, Me.Font, SystemBrushes.ControlText, textrect, sf)
                    sf.Alignment = StringAlignment.Center
                    Dim r As Rectangle = GetCloseButtonRect(id)
                    If HotTabIndex <> id OrElse Not onCloseButton Then
                        bmGraphics.DrawString("X", xFont, SystemBrushes.ControlText, r, sf)
                    ElseIf HotTabIndex = id And onCloseButton Then
                        bmGraphics.DrawString("X", xFont, Brushes.Red, r, sf)
                    End If
                End Using
                graphics.DrawImage(bm, tabRect)
            End Using

        Else
            Dim color As Color = If(id = HotTabIndex, SystemColors.GradientActiveCaption, Me.Parent.BackColor)
            Dim gp As New Drawing2D.GraphicsPath
            gp.AddPolygon(getPolygon())
            graphics.FillPath(New SolidBrush(color), gp)
            Dim r As Rectangle = GetPlusButtonRect()
            Dim plusFont As New Font(MyBase.Font.FontFamily, 16)
            Dim sf As New StringFormat
            sf.Alignment = StringAlignment.Center
            sf.LineAlignment = StringAlignment.Center
            If Not onAddButton Then
                graphics.DrawString("+", plusFont, Brushes.Black, r, sf)
            ElseIf onAddButton Then
                graphics.DrawString("+", plusFont, Brushes.Red, r, sf)
            End If
        End If
    End Sub

#End Region

#Region "     helper locating methods"

    ''' <summary> 
    ''' Helper method for drawing close buttons and responding to mouse activity 
    ''' </summary> 
    ''' <param name="x"></param> 
    ''' <returns></returns> 
    Private Function GetCloseButtonRect(x As Integer) As Rectangle
        Dim r As Rectangle = MyBase.GetTabRect(x)
        Dim closeRect As New Rectangle(r.Width - 18, If(x <> MyBase.SelectedIndex, 6, 4), 16, 16)
        Return closeRect
    End Function

    ''' <summary> 
    ''' Helper method for drawing add button and responding to mouse activity 
    ''' </summary> 
    ''' <returns></returns> 
    Private Function GetPlusButtonRect() As Rectangle
        Dim points() As Point = getPolygon()
        If points Is Nothing Then Return Nothing
        Return New Rectangle(points(0).X, 8, 16, 16)
    End Function

    ''' <summary> 
    ''' Helper method for drawing add button and responding to mouse activity 
    ''' </summary> 
    ''' <returns></returns> 
    Private Function getPolygon() As Point()
        Dim points As Point()
        Dim x As Integer
        If MyBase.TabPages.Count > 0 Then
            x = Enumerable.Range(0, MyBase.TabPages.Count).Select(Function(i) MyBase.GetTabRect(i).Right).Max + 1
        Else
            x = 1
        End If
        points = {New Point(x, 6), New Point(x + 12, 6), New Point(x + 16, 10), New Point(x + 16, 26), New Point(x, 26), New Point(x, 6)}
        Return points
    End Function

    ''' <summary> 
    ''' Method used for identifying which tab or button the mouse is over if any 
    ''' </summary> 
    ''' <param name="p"></param> 
    ''' <returns></returns> 
    Private Function getTabIndex(ByVal p As Point) As Integer
        For x As Integer = 0 To MyBase.TabPages.Count - 1
            If MyBase.GetTabRect(x).Contains(p) Then
                Return x
            End If
        Next
        Dim gp As New Drawing2D.GraphicsPath
        gp.AddPolygon(getPolygon())
        If gp.IsVisible(p) Then
            Return MyBase.TabPages.Count
        End If
        Return -1
    End Function

    ''' <summary> 
    ''' Used to capture 'mouse on add button' 
    ''' </summary> 
    ''' <returns></returns> 
    Private Function PointerOnAddButton() As Boolean
        Dim addButtonRect As Rectangle = GetPlusButtonRect()
        Dim pt As Point = MyBase.PointToClient(Cursor.Position)
        Return addButtonRect.Contains(pt)
    End Function

    ''' <summary> 
    ''' Used to capture 'mouse on close button' 
    ''' </summary> 
    ''' <param name="x"></param> 
    ''' <returns></returns> 
    Private Function PointerOnCloseButton(x As Integer) As Boolean
        If x > -1 Then
            Dim closeRect As Rectangle = GetCloseButtonRect(x)
            closeRect.X += MyBase.GetTabRect(x).Left
            Dim pt As Point = MyBase.PointToClient(Cursor.Position)
            If closeRect.Contains(pt) Then
                Return True
            End If
        End If
        Return False
    End Function

#End Region
End Class
