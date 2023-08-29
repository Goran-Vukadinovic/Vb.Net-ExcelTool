Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading
Imports Microsoft
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Class MainWindow

    Dim _form As MainWindow
    Dim _fileList As New List(Of String)
    Dim _IsRunning = False
    Dim _CurrentFolder As String = ""

    Dim blackColor = New BrushConverter().ConvertFrom("#000000")
    Dim yellowWhite = New BrushConverter().ConvertFrom("#FFDEEF48")
    Dim darkRedWhite = New BrushConverter().ConvertFrom("#FF730D09")
    Dim watcher As FileSystemWatcher
    Private Sub MainWindow_Load()
        ReadFolderPath()
        tbFolderPath.Text = _CurrentFolder
        Dim allfiles = Directory.GetFiles(_CurrentFolder, "*", SearchOption.TopDirectoryOnly)
        UpdateFileList(allfiles)
        UpdateMainListView()
        'Dispatcher.Invoke(Function()
        '                      MainProgressBar.Value = 0
        '                  End Function)
        'MainProgressBar.IsIndeterminate = True
        'StartWatcher()
    End Sub

    Private Sub UpdateIsRunningUI()
        btnRun.IsEnabled = Not _IsRunning
        btnFolder.IsEnabled = Not _IsRunning
        btnClear.IsEnabled = Not _IsRunning
        btnRun.IsEnabled = Not _IsRunning
    End Sub

    Private Sub UpdateMainListView()
        MainListView.Items.Clear()

        For Each filePath As String In _fileList
            Dim filename = Path.GetFileName(filePath)
            Dim item = New ComboBoxItem()
            item.Content = filename
            item.Foreground = blackColor
            item.FontFamily = New FontFamily("Verdana")
            item.FontSize = 14
            MainListView.Items.Add(item)
        Next
    End Sub

    Private Function AddFileNameListView(filepath As String) As ComboBoxItem
        Dim filename = Path.GetFileName(filepath)
        Dim item = New ComboBoxItem()
        item.Content = filename
        item.Foreground = blackColor
        item.FontFamily = New FontFamily("Verdana")
        item.FontSize = 14
        MainListView.Items.Add(item)
        Return item
    End Function

    Private Sub RunProcessing()
        Dim runbw = New BackgroundWorker
        runbw.WorkerReportsProgress = True
        runbw.WorkerSupportsCancellation = True
        AddHandler runbw.DoWork, AddressOf bw_DoMainWork
        AddHandler runbw.RunWorkerCompleted, AddressOf bw_DoWorkMainCompleted
        runbw.RunWorkerAsync()
    End Sub

    Public Delegate Sub UpdateProgressBarInvoker(ByVal value As Integer)

    Private Sub bw_DoWorkMainCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        Dispatcher.Invoke(Function()
                              MainProgressBar.Value = 100
                          End Function)
        _IsRunning = False
        ClearAllOriginalFiles()
        UpdateIsRunningUI()
        Dispatcher.Invoke(Function()
                              MainProgressBar.Value = 0
                          End Function)
        MainProgressBar.IsIndeterminate = True
        StartWatcher()
    End Sub

    Private Sub bw_DoMainWork(sender As Object, e As DoWorkEventArgs)

        CreateNewSubFolder("Edited")
        CreateNewSubFolder("Original")
        CreateNewSubFolder("Logs")

        Dim index = 1.0
        Dim length = _fileList.Count
        For Each filepath As String In _fileList
            Dim extension = Path.GetExtension(filepath)
            If extension = ".xlsx" Then
                Dim type = CheckExcelAndUpdateWorkBook(filepath)
                UpdateStyleExcel(filepath, type)
            Else
                Continue For
            End If

            Dispatcher.Invoke(Function()
                                  MainProgressBar.Value = (index / (length + 1.0)) * 100
                              End Function)
            index += 1.0
        Next
    End Sub

    Private Function CreateNewSubFolder(newfolder As String)
        If Not Directory.Exists(_CurrentFolder + $"/{newfolder}") Then
            Directory.CreateDirectory(_CurrentFolder + $"/{newfolder}")
        End If
    End Function

    Private Sub btnRun_Click(sender As Object, e As RoutedEventArgs) Handles btnRun.Click
        _IsRunning = True
        If _CurrentFolder = "" Then
            Return
        End If
        UpdateIsRunningUI()
        RunProcessing()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As RoutedEventArgs) Handles btnClear.Click
        _fileList.Clear()
        UpdateMainListView()
    End Sub

    Private Sub btnFolder_Click(sender As Object, e As RoutedEventArgs) Handles btnFolder.Click
        Dim fbDlg = New Forms.FolderBrowserDialog()
        Dim ret = fbDlg.ShowDialog()

        If ret = Forms.DialogResult.OK Then
            _fileList.Clear()
            Dim folderPath = fbDlg.SelectedPath
            _CurrentFolder = folderPath
            tbFolderPath.Text = _CurrentFolder
            Dim allfiles = Directory.GetFiles(folderPath, "*", SearchOption.TopDirectoryOnly)
            UpdateFileList(allfiles)
            UpdateMainListView()
            SaveFolderPath()
        End If
    End Sub

    Private Sub UpdateFileList(files As Object)
        For Each filename As String In files
            If Not filename.Contains("~$") And Path.GetExtension(filename) = ".xlsx" Then
                _fileList.Add(filename)
            End If
        Next
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Application.Current.Shutdown()
    End Sub

    Private Sub MainListView_Drop(sender As Object, e As DragEventArgs)
        Dim files = e.Data.GetData(DataFormats.FileDrop, False)
        UpdateFileList(files)
        UpdateMainListView()
    End Sub

    Private Sub MainListView_DragEnter(sender As Object, e As DragEventArgs)

        Dim a = e.Data.GetData(DataFormats.FileDrop, False)
        If a.Length <> 0 Then
            e.Effects = DragDropEffects.All
        Else
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Sub StartWatcher()
        If _CurrentFolder = "" Then
            Return
        End If
        watcher = New FileSystemWatcher
        watcher.Path = _CurrentFolder
        watcher.NotifyFilter = NotifyFilters.FileName Or NotifyFilters.CreationTime Or NotifyFilters.LastWrite
        watcher.Filter = "*.xlsx"
        AddHandler watcher.Created, AddressOf OnCreated
        watcher.EnableRaisingEvents = True
        GC.KeepAlive(watcher)
    End Sub

    Private Sub OnCreated(sender As Object, e As FileSystemEventArgs)
        Dim _watchFilePath = e.FullPath
        StartWatchBackService(_watchFilePath)
    End Sub

    Private Sub StartWatchBackService(filepath As String)
        Dim bw = New BackgroundWorker
        bw.WorkerReportsProgress = True
        bw.WorkerSupportsCancellation = True
        AddHandler bw.DoWork, AddressOf bw_DoWork
        AddHandler bw.RunWorkerCompleted, AddressOf bw_DoWorkCompleted
        bw.RunWorkerAsync(filepath)
    End Sub
    Private Sub bw_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Dim wathFilePath = e.Argument

        If wathFilePath = "" Or wathFilePath.Contains("~$") Then
            Return
        End If
        _IsRunning = True
        Dim extension = Path.GetExtension(wathFilePath)
        Thread.Sleep(200)
        If extension = ".xlsx" Then
            Dim item As ComboBoxItem
            Dispatcher.Invoke(Function()
                                  item = AddFileNameListView(wathFilePath)
                              End Function)
            Dim type = CheckExcelAndUpdateWorkBook(wathFilePath)
            UpdateStyleExcel(wathFilePath, type)
            File.Delete(wathFilePath)
            Dispatcher.Invoke(Function()
                                  MainListView.Items.Remove(item)
                              End Function)

        End If
    End Sub

    Private Sub bw_DoWorkCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        '_fileList.Clear()
        'MainListView.Items.Clear()
        _IsRunning = False
    End Sub

    Private Sub ClearAllOriginalFiles()
        If _CurrentFolder = "" And _IsRunning Then
            Return
        End If
        Dim allfiles = Directory.GetFiles(_CurrentFolder, "*", SearchOption.TopDirectoryOnly)
        For Each filepath As String In allfiles
            File.Delete(filepath)
        Next
        _fileList.Clear()
        MainListView.Items.Clear()
    End Sub

    Private Sub ReadFolderPath()
        Dim settingPath As String = String.Format("{0}\setting.txt", Environment.CurrentDirectory)
        If File.Exists(settingPath) = True Then
            _CurrentFolder = File.ReadAllText(settingPath, Encoding.UTF8)
        End If

        If _CurrentFolder = "" Then
            _CurrentFolder = String.Format("{0}\scan", Environment.CurrentDirectory)
        End If
        If Directory.Exists(_CurrentFolder) = False Then
            Directory.CreateDirectory(_CurrentFolder)
        End If
    End Sub

    Private Sub SaveFolderPath()
        Dim settingPath As String = String.Format("{0}\setting.txt", Environment.CurrentDirectory)
        File.WriteAllText(settingPath, _CurrentFolder)
    End Sub
End Class
