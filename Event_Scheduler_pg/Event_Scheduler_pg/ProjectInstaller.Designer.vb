<System.ComponentModel.RunInstaller(True)> Partial Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

    'Installer overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.ServiceProcessevent_scheduler_PG = New System.ServiceProcess.ServiceProcessInstaller()
        Me.ServiceInstallerevent_scheduler_PG = New System.ServiceProcess.ServiceInstaller()
        '
        'ServiceProcessevent_scheduler_PG
        '
        Me.ServiceProcessevent_scheduler_PG.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.ServiceProcessevent_scheduler_PG.Password = Nothing
        Me.ServiceProcessevent_scheduler_PG.Username = Nothing
        '
        'ServiceInstallerevent_scheduler_PG
        '
        Me.ServiceInstallerevent_scheduler_PG.ServiceName = "OmegacubeEDGEevent_scheduler_PG"
        Me.ServiceInstallerevent_scheduler_PG.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.ServiceProcessevent_scheduler_PG, Me.ServiceInstallerevent_scheduler_PG})

    End Sub
    Friend WithEvents ServiceProcessevent_scheduler_PG As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents ServiceInstallerevent_scheduler_PG As System.ServiceProcess.ServiceInstaller

End Class
