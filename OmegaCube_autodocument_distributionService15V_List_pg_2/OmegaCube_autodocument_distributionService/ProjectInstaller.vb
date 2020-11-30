Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.IO
Imports System.ServiceProcess

Public Class ProjectInstaller
    Protected Shared AppPath As String = System.AppDomain.CurrentDomain.BaseDirectory
    Protected Shared strLogFilePath As String = AppPath + "\" + "Autodocumentdistribution_ExLog_PG.txt"
    Private Shared sw As StreamWriter = Nothing
    Public Sub New()
        MyBase.New()
        'This call is required by the Component Designer.
        InitializeComponent()
        'Add initialization code after the call to InitializeComponent
    End Sub
    Private Sub ProjectInstaller_BeforeUninstall(sender As Object, e As InstallEventArgs) Handles Me.BeforeUninstall
        Try
            Dim controller As New ServiceController(AutodocumentdistributionServiceInstaller_PG.ServiceName)
            If controller.Status = ServiceControllerStatus.Running Or controller.Status = ServiceControllerStatus.Paused Then
                controller.Stop()
                controller.WaitForStatus(ServiceControllerStatus.Stopped, New TimeSpan(0, 0, 0, 15))
                controller.Close()
            End If
        Catch ex As Exception
            LogExceptions(strLogFilePath, ex)
        End Try
    End Sub

    Public Shared Sub LogExceptions(ByVal filePath As String, ByVal ex As Exception)
        If False = File.Exists(filePath) Then
            Dim fs As New FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite)
            fs.Close()
        End If
        WriteExceptionLog(filePath, ex)
    End Sub

    Private Shared Sub WriteExceptionLog(ByVal strPathName As String, ByVal objException As Exception)
        sw = New StreamWriter(strPathName, True)
        sw.WriteLine("Source     :" & objException.Source.ToString().Trim())
        sw.WriteLine("Method     : " & objException.TargetSite.Name.ToString())
        sw.WriteLine("Date       : " & DateTime.Now.ToLongTimeString())
        sw.WriteLine("Time       : " & DateTime.Now.ToShortDateString())
        sw.WriteLine("Error      : " & objException.Message.ToString().Trim())
        sw.WriteLine("Stack Trace: " & objException.StackTrace.ToString().Trim())
        sw.WriteLine("^^-------------------------------------------------------------------^^")
        sw.Flush()
        sw.Close()
    End Sub
End Class
