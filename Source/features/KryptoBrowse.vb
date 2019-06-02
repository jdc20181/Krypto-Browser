''WARNING: See the SourceCode directory, for build instructions
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Security
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports Gecko
Imports Gecko.Certificates

Public Class KryptoBrowse
#Region "Browser-Based Events"
    'This region is moving into its own class soon, handling Gecko seperately from the main code will be more effecient and follow better coding practices. 
    Dim int As Integer = 0
   
    Public Sub Loading(ByVal sender As Object, ByVal e As Gecko.GeckoProgressEventArgs)
        KryptoTabControl1.SelectedTab.Text = CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).DocumentTitle

        If KryptoTabControl1.SelectedTab.Text = Nothing Then
            KryptoTabControl1.SelectedTab.Text = CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).Url.ToString
        End If
    End Sub
    'Appropriately handles expired SSL as a temporary measure, they are ignored, however in the future confirmation dialogs will be given. 
    Private Sub IgnoreSSLError(ByVal sender As Object, ByVal e As Gecko.Events.GeckoNSSErrorEventArgs)
        If Not e.Message.ToLower().Contains("certificate") Then
            MsgBox(e.Message)
            e.Handled = False
            Return
        End If
        e.Handled = True
        CertOverrideService.GetService().RememberValidityOverride(e.Uri, e.Certificate, CertOverride.Mismatch Or CertOverride.Time Or CertOverride.Untrusted, True)
        CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).Navigate(e.Uri.AbsoluteUri)
    End Sub
    'Done Event is used to handle Document Completion, initates Favicon loading, and various other flags. 
    Public Sub Done(ByVal sender As Object, ByVal e As Gecko.Events.GeckoDocumentCompletedEventArgs)

        Me.Cursor = Cursors.Default
        Url_Text.Text = CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).Url.ToString

        If My.Settings.Incognito = "Yes" Then
            Dim field = GetType(GeckoWebBrowser).GetField("WebBrowser", BindingFlags.Instance Or BindingFlags.NonPublic)
            Dim nsIWebBrowser As nsIWebBrowser = DirectCast(field.GetValue(CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser)), nsIWebBrowser)

            Xpcom.QueryInterface(Of nsILoadContext)(nsIWebBrowser).SetPrivateBrowsing(True)

        ElseIf My.Settings.Incognito = "No" Then

            My.Settings.History.Add(CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).Url.ToString)

            Dim field = GetType(GeckoWebBrowser).GetField("WebBrowser", BindingFlags.Instance Or BindingFlags.NonPublic)
            Dim nsIWebBrowser As nsIWebBrowser = DirectCast(field.GetValue(CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser)), nsIWebBrowser)
            'this might be null if called right before initialization of browser
            Xpcom.QueryInterface(Of nsILoadContext)(nsIWebBrowser).SetPrivateBrowsing(False)
        End If


        'Force Https sites, for security purposes
        CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).Navigate("javascript: if (location.protocol == 'http:') location.href = location.href.replace(/^http:/, 'https:') return ;")
      

    End Sub
#End Region
#Region "Tab Drag"
    Dim pTabControlClickStartPosition As Point
    Dim iTabControlFlyOffPixelOffset As Integer = 20

    Private Sub TabControlMain_MouseDown(sender As Object, e As MouseEventArgs) Handles KryptoTabControl1.MouseDown
        If (e.Button = MouseButtons.Left) Then
            pTabControlClickStartPosition = e.Location
        End If
    End Sub

    Private Sub TabControlMain_MouseMove(sender As Object, e As MouseEventArgs) Handles KryptoTabControl1.MouseMove
        If (e.Button = MouseButtons.Left) Then
            Dim iMouseOffset_X = pTabControlClickStartPosition.X - e.X
            Dim iMouseOffset_Y = pTabControlClickStartPosition.Y - e.Y

            If iMouseOffset_X > iTabControlFlyOffPixelOffset Or iMouseOffset_Y > iTabControlFlyOffPixelOffset Then
                KryptoTabControl1.DoDragDrop(KryptoTabControl1.SelectedTab, DragDropEffects.Move)
                pTabControlClickStartPosition = New Point
            End If
        Else
            pTabControlClickStartPosition = New Point
        End If
    End Sub

    Private Sub TabControlMain_GiveFeedback(sender As Object, e As GiveFeedbackEventArgs) Handles KryptoTabControl1.GiveFeedback
        e.UseDefaultCursors = False
    End Sub

    Private Sub TabControlMain_QueryContinueDrag(sender As Object, e As QueryContinueDragEventArgs) Handles KryptoTabControl1.QueryContinueDrag
        If Control.MouseButtons <> MouseButtons.Left Then
            e.Action = DragAction.Cancel
            Dim f As New KryptoBrowse
            f.Size = New Size(800, 400)
            f.StartPosition = FormStartPosition.Manual
            f.Location = MousePosition
           
            f.KryptoTabControl1.TabPages.Add(Me.KryptoTabControl1.SelectedTab)

          
            My.Settings.StartUpEnabled = "No"



            AddHandler f.FormClosing, Sub(sender1 As Object, e1 As EventArgs)
                                    end sub
                             
            f.Show()
            Me.Cursor = Cursors.Default
        Else
            e.Action = DragAction.Continue
            Me.Cursor = Cursors.SizeAll
        End If
    End Sub
#End Region

    Public Sub getfav()
        Dim errorlogpath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) &
                         "\.\KryptoData" &
                         My.Settings.Errorlog
        Try
            Dim wc As WebClient = New WebClient()
            '"http://" &
            ServicePointManager.SecurityProtocol = CType(3072, SecurityProtocolType)
            Dim memorystream As MemoryStream = New MemoryStream(wc.DownloadData("http://" & New Uri(Url_Text.ToString()).Host & "/favicon.ico"))
            Dim icon As Icon = New Icon(memorystream)


            Dim i As String = Convert.ToString(ImageList1.Images.Count)

            ImageList1.Images.Add(i, icon.ToBitmap())



            KryptoTabControl1.ImageList = ImageList1
            KryptoTabControl1.SelectedTab.ImageIndex = ImageList1.Images.Count - 1
            'ToolStripButton3.ImageIndex = ImageList1.Images.Count - 1

        Catch e As Exception
            '   File.AppendAllText(errorlogpath, String.Format("{0}{1}", Environment.NewLine, e.ToString()))
            MsgBox("Check error logs")
            Using stwriter As New StreamWriter(errorlogpath, True)
                stwriter.WriteLine("-------------------START-------------" + DateTime.Now)
                ' stwriter.WriteLine(Convert.ToString("Page :"))
                stwriter.WriteLine(e)
                stwriter.WriteLine("-------------------END-------------" + DateTime.Now)
            End Using
        End Try

    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WindowState = FormWindowState.Maximized


        Drop_Menu.DropDownDirection = ToolStripDropDownDirection.BelowLeft
        Drop_Menu.Alignment = ToolStripItemAlignment.Right

        Url_Text.ShortcutsEnabled = False
        Url_Text.Control.ContextMenuStrip = ContextMenuStrip2
     
        If KryptoTabControl1.TabCount = 0 Then
            loadupsettings()
        End If
    End Sub
    Public Sub loadupsettings()
        Try
            Xpcom.Initialize("Firefox")
            Dim tab As New DoubleBufferedTabPage
            Dim brws As New GeckoWebBrowser With {
                .Dock = DockStyle.Fill
            }

            tab.Text = " New Tab"
            tab.Controls.Add(brws)
            Me.KryptoTabControl1.TabPages.Add(tab)
            Me.KryptoTabControl1.SelectedTab = tab
            GeckoPreferences.User("full-screen-api.enabled") = True
            '      brws.ContextMenuStrip = ContextMenuStrip1
            'GeckoPreferences.[Default]("full-screen-api.enabled") = True
            GeckoPreferences.Default("full-screen-api.enabled") = True
            ' Dim sUserAgent As String = "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:60.0) Gecko/20100101 Firefox/60.0.3"
            Dim sUserAgent As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:60.0) Gecko/20100101 Firefox/66.0 Krypto/0.0.1"

            GeckoPreferences.User("general.useragent.override") = sUserAgent
            'GeckoPreferences.Default("extensions.blocklist.enabled") = False
            GeckoPreferences.Default("general.useragent.override") = sUserAgent
            AddHandler brws.NSSError, AddressOf IgnoreSSLError

            brws.Navigate(My.Settings.Home)
            AddHandler brws.ProgressChanged, AddressOf Loading
            AddHandler brws.DocumentCompleted, AddressOf Done
            '   Int = Int() + 1
        Catch x As Exception
            MessageBox.Show("An Error occured" & x.ToString)
        End Try

    End Sub


    Private Sub moveControl(controlToMove As Control, newTab As TabPage)

        Dim findButton() As Control = newTab.Controls.Find(controlToMove.Name, True)

        If findButton.GetUpperBound(0) < 0 Then

            controlToMove.Parent = newTab

        End If

    End Sub
#Region "Naviagtion"
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles Back_Button.Click, BackToolStripMenuItem.Click
        Dim brws As New GeckoWebBrowser
        CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).GoBack()
        AddHandler brws.ProgressChanged, AddressOf Loading
        AddHandler brws.DocumentCompleted, AddressOf Done
    End Sub

    Private Sub Forward_Button_Click(sender As Object, e As EventArgs) Handles Forward_Button.Click, ForwardToolStripMenuItem.Click
        Dim brws As New GeckoWebBrowser
        CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).GoForward()
        AddHandler brws.ProgressChanged, AddressOf Loading
        AddHandler brws.DocumentCompleted, AddressOf Done
    End Sub

    'Private Sub Nav_button_Click(sender As Object, e As EventArgs) Handles Nav_button.Click
    '    UrlNavigate()
    'End Sub

    Private Sub Url_Text_KeyDown(sender As Object, e As KeyEventArgs) Handles Url_Text.KeyDown, Nav_button.Click
        If (e.KeyCode = Keys.Enter) Then
            UrlNavigate()
            e.Handled = True
            e.SuppressKeyPress = True
            'Dim s As String = Url_Text.Text
            'Dim fHasSpace As Boolean = s.Contains(" ")

            'If fHasSpace = True Then
            '    search()
            'ElseIf fHasSpace = False Then
            '    UrlNavigate()
            'End If
        End If
    End Sub
    Public Sub UrlNavigate()

   
        Dim pattern As String = "^(http|https|ftp|)\://|[a-zA-Z0-9\-\.]+\.[a-zA-Z](:[a-zA-Z0-9]*)?/?([a-zA-Z0-9\-\._\?\,\'/\\\+&amp;%\$#\=~])*[^\.\,\)\(\s](?:\d+)?$"
        Dim regex As Regex = New Regex(pattern, RegexOptions.Compiled Or RegexOptions.IgnoreCase)
        Dim url As String = Url_Text.Text.Trim()

        If regex.IsMatch(url) Then
            Dim brws As New GeckoWebBrowser

            AddHandler brws.ProgressChanged, AddressOf Loading
            AddHandler brws.DocumentCompleted, AddressOf Done
            int = int + 1

            CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).Navigate(Url_Text.Text)
        Else
            search()

        End If

    End Sub
    Public Sub search()
        Dim brws As New GeckoWebBrowser
        AddHandler brws.ProgressChanged, AddressOf Loading
        AddHandler brws.DocumentCompleted, AddressOf Done
        '   Int = int + 0.5
        CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).Navigate(My.Settings.SearchP & Url_Text.Text)
    End Sub



    Private Sub Add_Tab_Click(sender As Object, e As EventArgs) Handles NewTabToolStripMenuItem.Click

        Dim tab As New DoubleBufferedTabPage With {
            .Text = " New Tab"
        }

        Me.KryptoTabControl1.TabPages.Add(tab)
        Me.KryptoTabControl1.SelectedTab = tab
    End Sub



    Private Sub Refresh_Event_Click(sender As Object, e As EventArgs) Handles Refresh_Event.Click, RefreshgToolStripMenuItem.Click
        Dim brws As New GeckoWebBrowser
        CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).Refresh()
        AddHandler brws.ProgressChanged, AddressOf Loading
        AddHandler brws.DocumentCompleted, AddressOf Done
    End Sub
#End Region
    Private Sub KryptoBrowse_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim shouldSave = (My.Settings.RestoreSetting = "Enabled")
        Dim hastabs2 = (KryptoTabControl1.TabCount >= 2)
        If shouldSave AndAlso hastabs2 Then
            ' RestoreSave()
            'TODO add restore
        Else

        End If
        Dim shouldWarn = (My.Settings.TabCloseWarning = "Yes")
        Dim hasTabs = (KryptoTabControl1.TabCount >= 2)

        If shouldWarn AndAlso hasTabs Then
            Dim shouldCloseResult = MessageBox.Show("You have 2 or more tabs open. Are you sure you wanna exit?" & vbNewLine & "A Total of" & " " & KryptoTabControl1.TabCount & " " & "Tabs will be closed", "Closing Multi-Tabbed Window", MessageBoxButtons.YesNo)

            If shouldCloseResult = DialogResult.No Then
                e.Cancel = True
            End If
        End If


        '   shutdown the XULRunner services

    End Sub
    Public Sub MakeFullScreen()
        Me.SetVisibleCore(False)
        Me.FormBorderStyle = FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized
        Me.SetVisibleCore(True)
    End Sub

 

    Private Sub SavePageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SavePageToolStripMenuItem.Click
        Dim sfd = New SaveFileDialog()
        sfd.Filter = " Html File | *.html"
        If sfd.ShowDialog() = DialogResult.OK Then



            CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).SaveDocument(sfd.FileName)
        End If
    End Sub



    Private Sub PrintPageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintPageToolStripMenuItem.Click
        CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).Navigate("javascript:print()")

    End Sub
    'Temporary Button for private mode, boot, to move over, add your button, add your image, (disabled by default so add that one) go to my resources, add both images. Then add buttonname.image = my.resources.imagename on true/false 
    Private Sub PrivateModeToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs) Handles PrivateModeToolStripMenuItem.CheckedChanged
        If PrivateModeToolStripMenuItem.Checked = True Then
            My.Settings.Incognito = "Yes"
            My.Settings.Save()
            ' ToolStripLabel1.Visible = True
        ElseIf PrivateModeToolStripMenuItem.Checked = False Then
            My.Settings.Incognito = "No"
            My.Settings.Save()
            ' ToolStripLabel1.Visible = False
        End If
    End Sub












    Private Sub KryptoTabControl1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles KryptoTabControl1.SelectedIndexChanged
        If KryptoTabControl1.TabCount >= 1 Then


            Try



                Xpcom.Initialize("Firefox")
                '  Dim tab As New TabPage
                Dim brws As New GeckoWebBrowser With {
                    .Dock = DockStyle.Fill
                }


                Me.KryptoTabControl1.SelectedTab.Controls.Add(brws)
                'Me.Kryptotabcontrol1.TabPages.Add(tab)
                'Me.Kryptotabcontrol1.SelectedTab = tab
                GeckoPreferences.User("full-screen-api.enabled") = True
                GeckoPreferences.[Default]("full-screen-api.enabled") = True
                ' Dim sUserAgent As String = "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:60.0) Gecko/20100101 Firefox/60.0.3"
                Dim sUserAgent As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:60.0) Gecko/20100101 Firefox/66.0 Krypto/0.0.1"

                GeckoPreferences.User("general.useragent.override") = sUserAgent
                GeckoPreferences.Default("general.useragent.override") = sUserAgent


                brws.Navigate(My.Settings.Home)
                AddHandler brws.ProgressChanged, AddressOf Loading
                AddHandler brws.DocumentCompleted, AddressOf Done
                AddHandler brws.NSSError, AddressOf IgnoreSSLError
                '   Int = Int() + 1
            Catch x As Exception

            End Try
        End If
    End Sub





    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        My.Computer.Clipboard.Clear()
        If Url_Text.SelectionLength > 0 Then
            My.Computer.Clipboard.SetText(Url_Text.SelectedText)

        Else


        End If
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        If My.Computer.Clipboard.ContainsText Then
            Url_Text.Paste()
        End If
    End Sub

    Private Sub PasteGoSearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteGoSearchToolStripMenuItem.Click
        If My.Computer.Clipboard.ContainsText Then
            Url_Text.Paste()
            UrlNavigate()
        End If
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        Url_Text.SelectAll()
    End Sub

    Public Function logerrors(ByVal [error] As String)

        Dim filename As String = "Log_" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt"

        Dim filepath As String = "Log_" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt"

        If File.Exists(filepath) Then
            Using stwriter As New StreamWriter(filepath, True)
                stwriter.WriteLine(" - ------------------START - ------------" + DateTime.Now)
                stwriter.WriteLine(Convert.ToString("Page :"))
                stwriter.WriteLine([error])
                stwriter.WriteLine("-------------------END-------------" + DateTime.Now)
            End Using
        Else
            Dim stwriter As StreamWriter = File.CreateText(filepath)
            stwriter.WriteLine("-------------------START-------------" + DateTime.Now)
            stwriter.WriteLine(Convert.ToString("Page :"))
            stwriter.WriteLine([error])
            stwriter.WriteLine("-------------------END-------------" + DateTime.Now)
            stwriter.Close()
        End If
        Return [error]
    End Function



    Private Sub Url_Text_TextChanged(sender As Object, e As EventArgs) Handles Url_Text.TextChanged
        If Url_Text.Text = "krypto://about" Then
            CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).Navigate("https://beffsbrowser.org/")
        End If
    End Sub

    Private Sub ZoomInToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZoomInToolStripMenuItem.Click, Zoom_IN.Click
        CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).GetDocShellAttribute.GetContentViewerAttribute.SetFullZoomAttribute(CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).GetDocShellAttribute.GetContentViewerAttribute.GetFullZoomAttribute + CSng(0.1))

    End Sub

    Private Sub ZoomOutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZoomOutToolStripMenuItem.Click, Zoom_OUT.Click
        CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).GetDocShellAttribute.GetContentViewerAttribute.SetFullZoomAttribute(CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).GetDocShellAttribute.GetContentViewerAttribute.GetFullZoomAttribute - CSng(0.1))

    End Sub

    Private Sub ResetZoomToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetZoomToolStripMenuItem.Click, Zoom_RESET.Click
        CType(KryptoTabControl1.SelectedTab.Controls.Item(0), GeckoWebBrowser).GetDocShellAttribute.GetContentViewerAttribute.SetFullZoomAttribute(1)
    End Sub


    Private Sub KryptoBrowse_SizeChanged(sender As Object, e As EventArgs) Handles MyBase.SizeChanged
        If Me.WindowState = FormWindowState.Maximized Then
            Refresh_Event.Margin = New Padding(400, 1, 0, 2)
        ElseIf Me.WindowState = FormWindowState.Normal Then
            Refresh_Event.Margin = New Padding(220, 1, 0, 2)
        End If
    End Sub
End Class
