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
        Me.AutodocumentdistributionServiceProcessInstaller_PG = New System.ServiceProcess.ServiceProcessInstaller()
        Me.AutodocumentdistributionServiceInstaller_PG = New System.ServiceProcess.ServiceInstaller()
        '
        'AutodocumentdistributionServiceProcessInstaller_PG
        '
        Me.AutodocumentdistributionServiceProcessInstaller_PG.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.AutodocumentdistributionServiceProcessInstaller_PG.Password = Nothing
        Me.AutodocumentdistributionServiceProcessInstaller_PG.Username = Nothing
        '
        'AutodocumentdistributionServiceInstaller_PG
        '
        Me.AutodocumentdistributionServiceInstaller_PG.DelayedAutoStart = True
        Me.AutodocumentdistributionServiceInstaller_PG.ServiceName = "OmegaCube_AutoDocumentDistribution_PG"
        Me.AutodocumentdistributionServiceInstaller_PG.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.AutodocumentdistributionServiceProcessInstaller_PG, Me.AutodocumentdistributionServiceInstaller_PG})

    End Sub
    Friend WithEvents AutodocumentdistributionServiceProcessInstaller_PG As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents AutodocumentdistributionServiceInstaller_PG As System.ServiceProcess.ServiceInstaller

End Class
