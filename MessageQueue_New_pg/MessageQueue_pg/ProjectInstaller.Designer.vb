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
        Me.ServiceProcessMQueue_PG = New System.ServiceProcess.ServiceProcessInstaller()
        Me.ServiceInstallerMQueue_PG = New System.ServiceProcess.ServiceInstaller()
        '
        'ServiceProcessMQueue_PG
        '
        Me.ServiceProcessMQueue_PG.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.ServiceProcessMQueue_PG.Password = Nothing
        Me.ServiceProcessMQueue_PG.Username = Nothing
        '
        'ServiceInstallerMQueue_PG
        '
        Me.ServiceInstallerMQueue_PG.ServiceName = "OmegacubeEDGEMessageQueue_PG"
        Me.ServiceInstallerMQueue_PG.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.ServiceProcessMQueue_PG, Me.ServiceInstallerMQueue_PG})

    End Sub
    Friend WithEvents ServiceProcessMQueue_PG As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents ServiceInstallerMQueue_PG As System.ServiceProcess.ServiceInstaller

End Class
