Imports System.IO
Imports System.Security.Cryptography
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.VisualBasic.Devices
Imports OffinoTool.My.Resources


Public Class options
    'Global Varibals
    Dim RestoreErrorLevel As Boolean = False


    Dim downloadPath As String = My.Settings.DownloadPath.ToString()
    'LADEN VON GESPEICHERTEN ELEMENTES
    Private Sub options_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtDownloadPath.Text = My.Settings.DownloadPath.ToString()
        chkboxStandardInfo.SetItemChecked(0, My.Settings.chkboxSysteminfo)
        chkboxStandardInfo.SetItemChecked(1, My.Settings.chkboxIPconfigInfo)
        chkboxStandardInfo.SetItemChecked(2, My.Settings.chkboxPrinterInfo)

        chkboxAutoStart.Checked = My.Settings.AutoGetInfo
        cbbLanguage.SelectedIndex = My.Settings.Language

        txtDefaultSaveLocation.Text = My.Settings.DefaultSaveLocation.ToString()
        txtFirebirdPath.Text = My.Settings.FirebirdPath.ToString()
        txtBackupFilePath.Text = My.Settings.FirebirdBackupFilePath.ToString()
        'Load Tab 2
        txtUsername.Text = My.Settings.dbUsername.ToString()
        txtPassword.Text = My.Settings.dbPassword.ToString()
        txtDBPath.Text = My.Settings.dbPath.ToString()
        txtBackupPath.Text = My.Settings.dbBackupPath.ToString()
        txtBackupName.Text = My.Settings.dbBackupName.ToString()
        txtRestoreDBPath.Text = My.Settings.FirebirdRestoreDBPath.ToString()
    End Sub
    Private Sub btnfolderBrowser_Click(sender As Object, e As EventArgs)
        If funcSetPath.ShowDialog() = DialogResult.OK Then
            downloadPath = funcSetPath.SelectedPath
            My.Settings.DownloadPath = downloadPath
        End If
    End Sub

    'SPEICHERN FUNKTION
    Private Sub btnSaveAndClose_Click(sender As Object, e As EventArgs) Handles btnSaveAndClose.Click
        'Setzt checkliste in den Optionen
        My.Settings.chkboxSysteminfo = chkboxStandardInfo.GetItemChecked(0)
        My.Settings.chkboxIPconfigInfo = chkboxStandardInfo.GetItemChecked(1)
        My.Settings.chkboxPrinterInfo = chkboxStandardInfo.GetItemChecked(2)
        'Setzt Auto. Info Holen fest
        My.Settings.AutoGetInfo = chkboxAutoStart.Checked
        'Setzt DownlaodPfad
        My.Settings.DownloadPath = downloadPath
        My.Settings.DefaultSaveLocation = txtDefaultSaveLocation.Text
        My.Settings.FirebirdPath = txtFirebirdPath.Text
        My.Settings.FirebirdBackupFilePath = txtBackupFilePath.Text
        'Ändert die Anwendungs-sprache
        Dim Lang As String = cbbLanguage.SelectedIndex
        Dim cultureInfo As System.Globalization.CultureInfo

        Select Case Lang
            Case "1"
                cultureInfo = New System.Globalization.CultureInfo("en")
                cbbLanguage.SelectedIndex = 1
            Case "0"
                cultureInfo = New System.Globalization.CultureInfo("")
                cbbLanguage.SelectedIndex = 0
            Case Else
                cultureInfo = System.Globalization.CultureInfo.CurrentCulture
                cbbLanguage.SelectedIndex = cbbLanguage.SelectedIndex
        End Select
        My.Settings.Language = cbbLanguage.SelectedIndex
        Thread.CurrentThread.CurrentCulture = cultureInfo
        Thread.CurrentThread.CurrentUICulture = cultureInfo

        'Speichern von Firebird einstellungen
        'Try
        My.Settings.dbUsername = txtUsername.Text
        My.Settings.dbPassword = txtPassword.Text
        My.Settings.dbPath = txtDBPath.Text
        My.Settings.dbBackupPath = txtBackupPath.Text
        My.Settings.dbBackupName = txtBackupName.Text
        My.Settings.FirebirdRestoreDBPath = txtRestoreDBPath.Text
        'Catch ex As Exception
        'MsgBox(GlobalStrings.msgFailedToSave, vbCritical + vbOKOnly, GlobalStrings.txtError)
        'End Try

    End Sub

    Private Sub btnCancelSave_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnResetOptions_Click(sender As Object, e As EventArgs) Handles btnResetOptions.Click
        If MsgBox(GlobalStrings.msgResetOptions & vbCrLf & GlobalStrings.msgtResetOptionsWarning, vbYesNo + 48 + 256, GlobalStrings.lblWarning) = 6 Then
            My.Settings.Reset()
            Application.Restart()
        End If

    End Sub

    Private Sub cbbLanguage_SelectedIndexChanged(sender As Object, e As EventArgs)
        'THIS FUNCTION WILL / EVENT WILL FIRE ON LOAD SINCE CHANGES TO THE ccbLanguage ARE MADE IN THE LAOD FUNCTION!!!
    End Sub

    Private Sub btnTestPaths_Click(sender As Object, e As EventArgs) Handles btnTestPaths.Click
        If My.Computer.FileSystem.FileExists(txtDBPath.Text) Then
            If My.Computer.FileSystem.DirectoryExists(txtBackupPath.Text) Then
                MsgBox(GlobalStrings.msgPathsFound, 0, GlobalStrings.txtConfirmed)
            Else
                MsgBox(GlobalStrings.msgInvalidBackupPath, 0, GlobalStrings.txtError)
            End If
        Else
            MsgBox(GlobalStrings.msgInvalidDB, 0, GlobalStrings.txtError)
        End If
    End Sub

    Private Sub btnFirebirdDBPath_Click(sender As Object, e As EventArgs) Handles btnFirebirdDBPath.Click
        funcSelectFile.ShowDialog()
        If DialogResult.OK Then
            txtDBPath.Text = Me.funcSelectFile.FileName
        End If

    End Sub

    Private Sub btnFirebirdBackupPath_Click(sender As Object, e As EventArgs) Handles btnFirebirdBackupPath.Click
        FirebirdSetBackupPath.ShowDialog()
        If DialogResult.OK Then
            txtBackupPath.Text = Me.FirebirdSetBackupPath.SelectedPath
        End If

    End Sub

    Private Sub btnfolderBrowser_Click_1(sender As Object, e As EventArgs) Handles btnfolderBrowser.Click
        funcSetPath.ShowDialog()
        If DialogResult.OK Then
            txtDownloadPath.Text = Me.funcSetPath.SelectedPath
        End If

    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles btnCreateBackup.Click
        ProgressBar01.Style = ProgressBarStyle.Marquee
        Dim path As String = txtFirebirdPath.Text
        MsgBox(path & "\gbak.exe", 0, "Test")
        If My.Computer.FileSystem.FileExists(System.IO.Path.Combine(path, "gbak.exe")) Then
            If PreCheckForBackupDB() Then
                If txtBackupName.Text = "" Then
                    MsgBox(GlobalStrings.msgEmptyBackupName, 0, GlobalStrings.txtError)
                Else
                    Await CreateBackupFromDB()
                End If
            Else
            End If
        Else
            MsgBox(GlobalStrings.msgGbakNotFound, 0 + vbCritical, GlobalStrings.txtError)
        End If
        ProgressBar01.Style = ProgressBarStyle.Blocks
        ProgressBar01.Value = 0

    End Sub

    'nach click von Create Backup wird 
    Private Function PreCheckForBackupDB()
        Dim bool As Boolean = False
        If My.Computer.FileSystem.FileExists(txtDBPath.Text) Then
            If My.Computer.FileSystem.DirectoryExists(txtBackupPath.Text) Then
                bool = True
            Else
                MsgBox(GlobalStrings.msgInvalidBackupPath, 0, GlobalStrings.txtError)
            End If
        Else
            MsgBox(GlobalStrings.msgInvalidDB, 0, GlobalStrings.txtError)
        End If

        Return bool
    End Function

    'Firebird erstellen einer .fbk von einer datenbank mit fehler Abfangung!
    Async Function CreateBackupFromDB() As Task
        Await Task.Run(Sub()
                           Dim Output As String
                           Using P As New Process()
                               Try
                                   P.StartInfo.WorkingDirectory = txtFirebirdPath.Text
                                   P.StartInfo.FileName = "gbak.exe"
                                   P.StartInfo.Arguments = "-b " & txtDBPath.Text & " " & txtBackupPath.Text & "\" & txtBackupName.Text & ".fbk" & " -user " & My.Settings.dbUsername & " -pass " & My.Settings.dbPassword & " -y gbak.log"
                                   P.StartInfo.UseShellExecute = False
                                   P.StartInfo.RedirectStandardInput = True
                                   P.StartInfo.RedirectStandardOutput = True
                                   P.StartInfo.CreateNoWindow = True
                                   P.Start()

                                   'P.StandardInput.WriteLine("gbak -b " & txtDBPath.Text & " " & txtBackupPath.Text & "\backup01.fbk" & " -user " & My.Settings.dbUsername & " -pass " & My.Settings.dbPassword & " -v")
                                   'P.StandardInput.WriteLine("Exit")
                                   ' Output = P.StandardOutput.ReadToEnd()
                                   P.WaitForExit()
                                   'Output = Output.ToLower()
                                   Output = My.Computer.FileSystem.ReadAllText("gbak.log")
                                   'MsgBox(Output, 0, "Debug")
                                   If Output.Contains("error") Then
                                       MsgBox(GlobalStrings.msgGbakError, 0 + vbCritical, GlobalStrings.txtError)
                                       Exit Try
                                   End If
                               Catch ex As Exception
                                   MsgBox(GlobalStrings.msgGbakNotFound, 0 + vbCritical, GlobalStrings.txtError)
                               End Try
                           End Using

                       End Sub)
    End Function

    'RESTORE START
    Private Async Sub btnFirebirdRestoreStart_Click(sender As Object, e As EventArgs) Handles btnFirebirdRestoreStart.Click
        If RestoreErrorLevel = False Then
            Await shutdownDatabase()
            If RestoreErrorLevel = False Then
                Await restoreDatabase()
                'Await ReadGfixLog(Output)
                If RestoreErrorLevel = False Then
                    Await startDatabase()
                End If
            End If
        End If
        MsgBox("Fertig", 0, "Fertig")
        'zurücksetzen des ErrorLevels
        RestoreErrorLevel = False
    End Sub

    'shutdown der DB
    Async Function shutdownDatabase() As Task
        Await Task.Run(Sub()
                           Using P As New Process()
                               Try
                                   P.StartInfo.FileName = "gfix.exe"
                                   P.StartInfo.WorkingDirectory = txtFirebirdPath.Text
                                   P.StartInfo.Arguments = "-shut full -tran 60 " & txtRestoreDBPath.Text & " -user " & My.Settings.dbUsername & " -pass " & My.Settings.dbPassword
                                   P.StartInfo.UseShellExecute = False
                                   P.StartInfo.RedirectStandardInput = True
                                   P.StartInfo.RedirectStandardOutput = True
                                   P.StartInfo.CreateNoWindow = True
                                   P.Start()

                                   P.WaitForExit()
                               Catch ex As Exception
                                   MsgBox("Fehler beim Runterfahren der Datenbank!", 0 + vbCritical, GlobalStrings.txtError)
                                   RestoreErrorLevel = True
                               End Try
                           End Using
                       End Sub)
    End Function
    'restore
    Async Function restoreDatabase() As Task
        Await Task.Run(Sub()


                           Using P As New Process()
                               Try
                                   Dim Output As String
                                   P.StartInfo.WorkingDirectory = txtFirebirdPath.Text
                                   P.StartInfo.FileName = "gbak.exe"
                                   P.StartInfo.Arguments = "-rep " & txtBackupFilePath.Text & " " & txtRestoreDBPath.Text & " -user " & My.Settings.dbUsername & " -pass " & My.Settings.dbPassword & " -v "
                                   P.StartInfo.UseShellExecute = False
                                   P.StartInfo.RedirectStandardInput = True
                                   P.StartInfo.RedirectStandardOutput = True
                                   P.StartInfo.CreateNoWindow = False
                                   P.Start()
                                   Output = P.StandardOutput.ReadToEnd()
                                   P.WaitForExit()

                                   If Output.Contains("error") Then
                                       MsgBox(GlobalStrings.msgGbakError, 0 + vbCritical, GlobalStrings.txtError)
                                       RestoreErrorLevel = True
                                       Exit Try
                                   End If
                                   'MsgBox(Output, 0, "Debug")
                               Catch ex As Exception
                                   MsgBox("Fehler beim restore der Datenbank!", 0 + vbCritical, GlobalStrings.txtError)
                                   RestoreErrorLevel = True
                               End Try
                           End Using
                       End Sub)
    End Function

    Async Function startDatabase() As Task
        Await Task.Run(Sub()
                           Using P As New Process()
                               Try
                                   P.StartInfo.FileName = "gfix.exe"
                                   P.StartInfo.WorkingDirectory = txtFirebirdPath.Text
                                   P.StartInfo.Arguments = "-o " & txtRestoreDBPath.Text & " -user " & My.Settings.dbUsername & " -pass " & My.Settings.dbPassword
                                   P.StartInfo.UseShellExecute = False
                                   P.StartInfo.RedirectStandardInput = True
                                   P.StartInfo.RedirectStandardOutput = True
                                   P.StartInfo.CreateNoWindow = True
                                   P.Start()
                                   P.WaitForExit()
                               Catch ex As Exception
                                   MsgBox("Fehler Beim Starten der Datenbank!", 0 + vbCritical, GlobalStrings.txtError)
                                   RestoreErrorLevel = True
                               End Try
                           End Using
                       End Sub)
    End Function

    'RESTORE END

    'Passwort anzeige ändern
    Private Sub chbShowHidePass_CheckedChanged(sender As Object, e As EventArgs) Handles chbShowHidePass.CheckedChanged
        Dim bool As Boolean = txtPassword.UseSystemPasswordChar
        Select Case bool
            Case True
                txtPassword.UseSystemPasswordChar = False
                chbShowHidePass.BackgroundImage = eye_slash_regular
            Case False
                txtPassword.UseSystemPasswordChar = True
                chbShowHidePass.BackgroundImage = eye_slash_regularOpen

        End Select
    End Sub

    'Bestätigen der sprache mit refresh der elemente
    Private Sub btnConfirmLanguage_Click(sender As Object, e As EventArgs) Handles btnConfirmLanguage.Click
        Dim Lang As String = cbbLanguage.SelectedIndex
        Dim cultureInfo As System.Globalization.CultureInfo

        Select Case Lang
            Case "1"
                cultureInfo = New System.Globalization.CultureInfo("en")
                cbbLanguage.SelectedIndex = 1
            Case "0"
                cultureInfo = New System.Globalization.CultureInfo("")
                cbbLanguage.SelectedIndex = 0
            Case Else
                cultureInfo = System.Globalization.CultureInfo.CurrentCulture
                cbbLanguage.SelectedIndex = cbbLanguage.SelectedIndex
        End Select
        My.Settings.Language = cbbLanguage.SelectedIndex
        Thread.CurrentThread.CurrentCulture = cultureInfo
        Thread.CurrentThread.CurrentUICulture = cultureInfo

        Me.Controls.Clear()
        InitializeComponent()
        MainForm.LanguageChanged()
        cbbLanguage.SelectedIndex = My.Settings.Language
    End Sub


    Private Sub btnSetRestoreDBPath_Click(sender As Object, e As EventArgs) Handles btnSetRestoreDBPath.Click
        funcSelectFile.ShowDialog()
        If DialogResult.OK Then
            txtRestoreDBPath.Text = Me.funcSelectFile.FileName
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btnSetDefaultSaveLocation.Click
        funcSetPath.ShowDialog()
        If DialogResult.OK Then
            txtDefaultSaveLocation.Text = Me.funcSetPath.SelectedPath
        End If
    End Sub

    Private Sub btnSetFirebirdPath_Click(sender As Object, e As EventArgs) Handles btnSetFirebirdPath.Click
        funcSetPath.ShowDialog()
        If DialogResult.OK Then
            txtFirebirdPath.Text = Me.funcSetPath.SelectedPath
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btnSetBackupFilePath.Click
        funcSelectBackupFile.ShowDialog()
        If DialogResult.OK Then
            txtBackupFilePath.Text = Me.funcSelectBackupFile.FileName
        End If
    End Sub

    Private Sub lblFirebirdRestoreBackup_MouseHover(sender As Object, e As EventArgs) Handles lblFirebirdRestoreBackup.MouseHover
        MainForm.test()
    End Sub
End Class
