Imports System.Runtime.InteropServices

Public Class Form1
    Public Const MOD_ALT As Integer = &H1 'Alt key
    Public Const WM_HOTKEY As Integer = &H312

    <DllImport("User32.dll")>
    Public Shared Function RegisterHotKey(ByVal hwnd As IntPtr,
                        ByVal id As Integer, ByVal fsModifiers As Integer,
                        ByVal vk As Integer) As Integer
    End Function

    <DllImport("User32.dll")>
    Public Shared Function UnregisterHotKey(ByVal hwnd As IntPtr,
                        ByVal id As Integer) As Integer
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object,
                        ByVal e As System.EventArgs) Handles MyBase.Load
        RegisterHotKey(Me.Handle, 100, MOD_ALT, Keys.P)
        'RegisterHotKey(Me.Handle, 200, MOD_ALT, Keys.C)

        savePath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
        file_Prefix = "ScreenShot_"
        shotNum = 1
        Label1.Text = "Save Path: " & savePath
        Label2.Text = " File Prefix: " & file_Prefix
    End Sub

    Public savePath As String, file_Prefix As String, shotNum As Integer

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = WM_HOTKEY Then
            Dim id As IntPtr = m.WParam
            Select Case (id.ToString)
                Case "100"
                    do_Snap()
                Case "200"
                    MessageBox.Show("You pressed ALT+C key combination")
            End Select
        End If
        MyBase.WndProc(m)
    End Sub


    Private Sub Form1_FormClosing(ByVal sender As System.Object,
                        ByVal e As System.Windows.Forms.FormClosingEventArgs) _
                        Handles MyBase.FormClosing
        UnregisterHotKey(Me.Handle, 100)
        'UnregisterHotKey(Me.Handle, 200)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        do_snap
    End Sub
    Private Sub do_Snap()
        shotNum = shotNumSelector.Value
        Form2.Hide()
        Dim area As Rectangle
        Dim capture As System.Drawing.Bitmap
        Dim graph As Graphics
        Dim strShotNum As String
        area = Form2.Bounds

        capture = New System.Drawing.Bitmap(area.Width, area.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        graph = Graphics.FromImage(capture)
        graph.CopyFromScreen(area.X, area.Y, 0, 0, area.Size, CopyPixelOperation.SourceCopy)
        PictureBox1.Image = capture
        'make sure we do 001 002 etc
        If shotNum < 10 Then
            strShotNum = "00" & shotNum
        ElseIf shotNum < 100 Then
            strShotNum = "0" & shotNum
        Else
            strShotNum = CStr(shotNum)
        End If

        PictureBox1.Image.Save(savePath & "\" & file_Prefix & strShotNum & ".png", System.Drawing.Imaging.ImageFormat.Png)
        shotNum = shotNum + 1
        shotNumSelector.Value = shotNum

    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim save As New FolderBrowserDialog
        Try
            save.Description = "Please select the destination folder for your images."
            save.SelectedPath = savePath
            If save.ShowDialog() = DialogResult.OK Then
                savePath = save.SelectedPath
                file_Prefix = InputBox("Image File Names", "Please specify a base filename:", "Screenshots_")

                Label1.Text = "Save Path: " & savePath
                Label2.Text = " File Prefix: " & file_Prefix
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Form2.Show()
    End Sub
End Class
