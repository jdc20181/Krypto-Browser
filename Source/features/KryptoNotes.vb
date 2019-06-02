''WARNING See Build Guide before using.

Imports Gecko

Public Class KryptoNotes
    'TODO: Organize Code
    Private Sub KryptoNotes_SizeChanged(sender As Object, e As EventArgs) Handles MyBase.SizeChanged
        'textbox1.Location = New Point(Me.Width \ 2 - textbox1.Width \ 2, Me.Height \ 2 - textbox1.Height \ 2)
        If Me.WindowState = FormWindowState.Maximized Then
            textbox1.Size = New Size(textbox1.Height + 30, textbox1.Width)
            textbox1.Location = New Point(Me.Width \ 2 - textbox1.Width \ 2, Me.Height \ 2 - textbox1.Height \ 2)
        ElseIf Me.WindowState = FormWindowState.Normal Then
            textbox1.Location = New Point(Me.Width \ 2 - textbox1.Width \ 2, Me.Height \ 2 - textbox1.Height \ 2)
            textbox1.Size = New Size(textbox1.Height - 30, textbox1.Width)
        End If

    End Sub
    Private currentFile As String
    Private checkPrint As Integer
    Private Sub NewToolStripButton_Click(sender As Object, e As EventArgs) Handles NewToolStripButton.Click
        If textbox1.Modified Then

            Dim answer As Integer
            answer = MessageBox.Show("The current document has not been saved, would you like to continue without saving?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If answer = DialogResult.Yes Then
                textbox1.Clear()
            Else
                Exit Sub


            End If

        Else

            textbox1.Clear()

        End If

        currentFile = ""
        Me.Text = "Editor: New Document"
    End Sub

    Private Sub OpenToolStripButton_Click(sender As Object, e As EventArgs) Handles OpenToolStripButton.Click
        'Check if there's text added to the textbox
        If textbox1.Modified Then
            'If the text of notepad changed, the program will ask the user if they want to save the changes
            Dim ask As MsgBoxResult
            ask = MsgBox("Do you want to save the changes", MsgBoxStyle.YesNoCancel, "Open Document")
            If ask = MsgBoxResult.No Then
                OpenFileDialog1.ShowDialog()
                textbox1.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
            ElseIf ask = MsgBoxResult.Cancel Then
            ElseIf ask = MsgBoxResult.Yes Then
                SaveFileDialog1.ShowDialog()
                My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, textbox1.Text, False)
                textbox1.Clear()
            End If
        Else
            'If textbox's text is still the same, notepad will show the OpenFileDialog
            OpenFileDialog1.ShowDialog()
            Try
                textbox1.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub SaveToolStripButton_Click(sender As Object, e As EventArgs) Handles SaveToolStripButton.Click
        SaveFile(currentFile)
    End Sub
    Private Sub SaveFile(ByVal strFileName As String)

        If strFileName = String.Empty Then
            'strFileName = "C:\Documents\" & Date.Now.ToString("MM-dd-yyyy HH\hmm\minss\s") & ".rtf"
            strFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments & Date.Now.ToString("MM-dd-yyyy HH\hmm\minss\s") & My.Settings.EasyNoteExt

        End If

        Dim strExt As String = System.IO.Path.GetExtension(strFileName).ToUpper()

        Select Case strExt
            Case ".RTF"
                textbox1.SaveFile(strFileName)
            Case Else
                ' to save as plain text
                Dim txtWriter As System.IO.StreamWriter
                txtWriter = New System.IO.StreamWriter(strFileName)
                txtWriter.Write(textbox1.Text)
                txtWriter.Close()
                txtWriter = Nothing
                textbox1.SelectionStart = 0
                textbox1.SelectionLength = 0
                textbox1.Modified = False
        End Select

        Me.Text = "Editor: " & strFileName

    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        SaveFile(currentFile)
        Timer1.Stop()
        Timer2.Start()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        SaveFile(currentFile)
        'Timer2.Start()


    End Sub

    Private Sub PrintToolStripButton_Click(sender As Object, e As EventArgs) Handles PrintToolStripButton.Click
        PrintDialog1.Document = PrintDocument1

        If PrintDialog1.ShowDialog() = DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub CutToolStripButton_Click(sender As Object, e As EventArgs) Handles CutToolStripButton.Click
        My.Computer.Clipboard.Clear()
        If textbox1.SelectionLength > 0 Then
            My.Computer.Clipboard.SetText(textbox1.SelectedText)

        End If
        textbox1.SelectedText = ""
    End Sub

    Private Sub CopyToolStripButton_Click(sender As Object, e As EventArgs) Handles CopyToolStripButton.Click
        My.Computer.Clipboard.Clear()
        If textbox1.SelectionLength > 0 Then
            My.Computer.Clipboard.SetText(textbox1.SelectedText)
        Else

        End If
    End Sub

    Private Sub PasteToolStripButton_Click(sender As Object, e As EventArgs) Handles PasteToolStripButton.Click
        If My.Computer.Clipboard.ContainsText Then
            textbox1.Paste()
        End If
    End Sub
    Private Sub InitializeMyContextMenu()
        ' Create the contextMenu and the MenuItem to add.
        Dim contextMenu1 As New ContextMenu()
        Dim menuItem1 As New MenuItem("C&ut")
        AddHandler menuItem1.Click, AddressOf CutToolStripButton_Click
        Dim menuItem2 As New MenuItem("&Copy")
        AddHandler menuItem2.Click, AddressOf CopyToolStripButton_Click
        Dim menuItem3 As New MenuItem("&Paste")
        AddHandler menuItem3.Click, AddressOf PasteToolStripButton_Click
        ' Use the MenuItems property to call the Add method
        ' to add the MenuItem to the MainMenu menu item collection.
        contextMenu1.MenuItems.Add(menuItem1)
        contextMenu1.MenuItems.Add(menuItem2)
        contextMenu1.MenuItems.Add(menuItem3)
        ' Assign mainMenu1 to the rich text box.
        textbox1.ContextMenu = contextMenu1
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        textbox1.SelectAll()
    End Sub

    Private Sub KryptoNotes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeMyContextMenu()

        Me.WindowState = FormWindowState.Maximized
    End Sub
    Private Sub IncreaseToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles IncreaseToolStripMenuItem1.Click
        Try
            textbox1.SelectionFont = New Font(textbox1.SelectionFont.FontFamily, Int(textbox1.SelectionFont.SizeInPoints + 5))
        Catch ex As Exception
        End Try
        textbox1.Focus()
    End Sub

    Private Sub DecreaseToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DecreaseToolStripMenuItem1.Click
        Try
            textbox1.SelectionFont = New Font(textbox1.SelectionFont.FontFamily, Int(textbox1.SelectionFont.SizeInPoints - 5))
        Catch ex As Exception
        End Try
        textbox1.Focus()

    End Sub

    Private Sub FontsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles FontsToolStripMenuItem1.Click
        FontDialog1.ShowDialog()
        textbox1.Font = FontDialog1.Font
    End Sub

    Private Sub FontColorsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontColorsToolStripMenuItem.Click
        ColorDialog1.ShowDialog()
        textbox1.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub BoldToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BoldToolStripMenuItem.Click
        If textbox1.SelectionFont.Bold = True Then
            If textbox1.SelectionFont.Italic = True Then
                textbox1.SelectionFont = New Font(Me.textbox1.SelectionFont, FontStyle.Regular + FontStyle.Italic)
            Else
                textbox1.SelectionFont = New Font(Me.textbox1.SelectionFont, FontStyle.Regular)
            End If

        ElseIf textbox1.SelectionFont.Bold = False Then
            If textbox1.SelectionFont.Italic = True Then
                textbox1.SelectionFont = New Font(Me.textbox1.SelectionFont, FontStyle.Bold + FontStyle.Italic)
            Else
                textbox1.SelectionFont = New Font(Me.textbox1.SelectionFont, FontStyle.Bold)
            End If
        End If
        textbox1.Focus()
    End Sub

    Private Sub ItalicsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ItalicsToolStripMenuItem.Click
        If textbox1.SelectionFont.Italic = True Then
            If textbox1.SelectionFont.Bold = True Then
                textbox1.SelectionFont = New Font(Me.textbox1.SelectionFont, FontStyle.Regular + FontStyle.Bold)
            Else
                textbox1.SelectionFont = New Font(Me.textbox1.SelectionFont, FontStyle.Regular)
            End If

        ElseIf textbox1.SelectionFont.Italic = False Then
            If textbox1.SelectionFont.Bold = True Then
                textbox1.SelectionFont = New Font(Me.textbox1.SelectionFont, FontStyle.Italic + FontStyle.Bold)
            Else
                textbox1.SelectionFont = New Font(Me.textbox1.SelectionFont, FontStyle.Italic)
            End If
        End If
        textbox1.Focus()
    End Sub

    Private Sub UnderLineToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UnderLineToolStripMenuItem.Click
        If Not textbox1.SelectionFont Is Nothing Then

            Dim currentFont As System.Drawing.Font = textbox1.SelectionFont
            Dim newFontStyle As System.Drawing.FontStyle

            If textbox1.SelectionFont.Underline = True Then
                newFontStyle = FontStyle.Regular
            Else
                newFontStyle = FontStyle.Underline
            End If

            textbox1.SelectionFont = New Font(currentFont.FontFamily, currentFont.Size, newFontStyle)

        End If
    End Sub

    Private Sub NormalToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NormalToolStripMenuItem.Click
        If Not textbox1.SelectionFont Is Nothing Then

            Dim currentFont As System.Drawing.Font = textbox1.SelectionFont
            Dim newFontStyle As System.Drawing.FontStyle
            newFontStyle = FontStyle.Regular

            textbox1.SelectionFont = New Font(currentFont.FontFamily, currentFont.Size, newFontStyle)

        End If
    End Sub

    Private Sub LeftToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LeftToolStripMenuItem.Click
        textbox1.SelectionAlignment = HorizontalAlignment.Left
        LeftToolStripMenuItem.Checked = True
        CenterToolStripMenuItem.Checked = False
        RightToolStripMenuItem.Checked = False
    End Sub

    Private Sub CenterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CenterToolStripMenuItem.Click
        textbox1.SelectionAlignment = HorizontalAlignment.Center
        LeftToolStripMenuItem.Checked = False
        CenterToolStripMenuItem.Checked = True
        RightToolStripMenuItem.Checked = False
    End Sub

    Private Sub RightToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RightToolStripMenuItem.Click
        textbox1.SelectionAlignment = HorizontalAlignment.Right
        LeftToolStripMenuItem.Checked = False
        CenterToolStripMenuItem.Checked = True
        RightToolStripMenuItem.Checked = False
    End Sub

    Private Sub PrintDocument1_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint


        checkPrint = 0

    End Sub


    Private Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        checkPrint = textbox1.Print(checkPrint, textbox1.TextLength, e)


        If checkPrint < textbox1.TextLength Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If

    End Sub

    Private Sub InsertPictureToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InsertPictureToolStripMenuItem.Click
        Try
            Dim GetPicture As New OpenFileDialog
            GetPicture.Filter = "PNGs (*.png), Bitmaps (*.bmp), GIFs (*.gif), JPEGs (*.jpg)|*.bmp;*.gif;*.jpg;*.png|PNGs (*.png)|*.png|Bitmaps (*.bmp)|*.bmp|GIFs (*.gif)|*.gif|JPEGs (*.jpg)|*.jpg"
            GetPicture.FilterIndex = 1
            GetPicture.InitialDirectory = "C:\"
            If GetPicture.ShowDialog = DialogResult.OK Then
                Dim SelectedPicture As String = GetPicture.FileName
                Dim Picture As Bitmap = New Bitmap(SelectedPicture)
                Dim cboard As Object = Clipboard.GetData(System.Windows.Forms.DataFormats.Text)
                Clipboard.SetImage(Picture)
                Dim PictureFormat As DataFormats.Format = DataFormats.GetFormat(DataFormats.Bitmap)
                If textbox1.CanPaste(PictureFormat) Then
                    textbox1.Paste(PictureFormat)
                End If
                Clipboard.Clear()
                Clipboard.SetText(cboard)
            End If
        Catch ex As Exception
        End Try
    End Sub
#Region "URL Detection stuff"
    Public Property DetectUrls As Boolean
    Private Sub textbox1_LinkClicked(sender As Object, e As LinkClickedEventArgs) Handles textbox1.LinkClicked
        Dim result As Integer = MessageBox.Show("Follow Link", "Open this link?" & vbNewLine & e.LinkText, MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Cancel Then

        ElseIf result = DialogResult.No Then
            ' MessageBox.Show("No pressed")
        ElseIf result = DialogResult.Yes Then
            Dim F As New KryptoBrowse
            F.Show()
            Dim tab As New DoubleBufferedTabPage
            Dim brws As New GeckoWebBrowser

            brws.Dock = DockStyle.Fill
            tab.Text = " New Tab"
            tab.Controls.Add(brws)
            F.KryptoTabControl1.TabPages.Add(tab)
            F.KryptoTabControl1.SelectedTab = tab
            brws.Navigate(e.LinkText) 'change it to your browser control name
            AddHandler brws.ProgressChanged, AddressOf F.Loading
            AddHandler brws.DocumentCompleted, AddressOf F.Done

        End If
    End Sub

    Private Sub SearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchToolStripMenuItem.Click
        Dim a As String
        Dim b As String
        a = InputBox("Enter text to be found")
        b = InStr(textbox1.Text, a)
        If b Then
            textbox1.Focus()
            textbox1.SelectionStart = b - 1
            textbox1.SelectionLength = Len(a)
        Else
            MsgBox("Text not found.")
        End If
    End Sub
    'Private Sub Link_Clicked(sender As Object, e As System.Windows.Forms.LinkClickedEventArgs)
    '    System.Diagnostics.Process.Start(e.LinkText)
    '    BBMain.CheckBox2.Checked = True
    '    BBMain.ToolStripTextBox1.Text = e.LinkText
    'End Sub 'Link_Clicked
#End Region
    Private Sub RichTextBox1_ContentsResized(sender As Object, e As ContentsResizedEventArgs) Handles textbox1.ContentsResized
        If textbox1.TextLength > 500 Then
            Dim mySize As SizeF = New SizeF()
            Dim myFont As Font = New Font(Me.textbox1.Font.FontFamily, Me.textbox1.Font.Size)
            Dim g As Graphics = Graphics.FromHwnd(textbox1.Handle)
            mySize = g.MeasureString(textbox1.Text, myFont)

            Me.textbox1.Height = CInt(Math.Round(mySize.Height + textbox1.Height, 0))
            '        Panel1.Size = New Size(textbox1.Height, textbox1.Width)
        End If
    End Sub

    Private Sub KryptoNotes_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If textbox1.Modified Then

            Dim answer As Integer
            answer = MessageBox.Show("The current document has not been saved, would you like to continue without saving?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If answer = DialogResult.Yes Then
                Close()

            Else
                Exit Sub
            End If

        Else

            textbox1.Clear()

        End If
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        '  textbox1.InsertLink(textbox1.SelectedText, "https://google.com")
        Dim userMsg As String
        userMsg = InputBox("Please Type your link", "Insert a link", "Enter URL", 500, 700)
        If userMsg <> "" Then
            If textbox1.SelectionLength = Not vbEmpty Then
                textbox1.InsertLink(textbox1.SelectedText, userMsg)
            ElseIf textbox1.SelectionLength = vbEmpty Then
                textbox1.InsertLink(userMsg.ToString, userMsg)
            End If
        Else
                MessageBox.Show("Not a URL")
        End If
    End Sub
End Class
