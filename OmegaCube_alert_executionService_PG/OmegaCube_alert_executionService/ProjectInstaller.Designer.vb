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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.AlertexecutionServiceProcessInstaller_PG = New System.ServiceProcess.ServiceProcessInstaller()
        Me.AlertexecutionServiceInstaller_PG = New System.ServiceProcess.ServiceInstaller()
        '
        'AlertexecutionServiceProcessInstaller_PG
        '
        Me.AlertexecutionServiceProcessInstaller_PG.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.AlertexecutionServiceProcessInstaller_PG.Password = Nothing
        Me.AlertexecutionServiceProcessInstaller_PG.Username = Nothing
        '
        'AlertexecutionServiceInstaller_PG
        '
        Me.AlertexecutionServiceInstaller_PG.DelayedAutoStart = True
        Me.AlertexecutionServiceInstaller_PG.ServiceName = "OmegaCube_AlertExecution_PG"
        Me.AlertexecutionServiceInstaller_PG.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.AlertexecutionServiceProcessInstaller_PG, Me.AlertexecutionServiceInstaller_PG})

    End Sub
    Friend WithEvents AlertexecutionServiceProcessInstaller_PG As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents AlertexecutionServiceInstaller_PG As System.ServiceProcess.ServiceInstaller

End Class
