Public Class ThisAddIn
    Private Sub ThisAddIn_Startup() Handles Me.Startup
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
    End Sub

    Protected Overrides Function RequestComAddInAutomationService() As Object
        Return New ApiAddin1()
    End Function

End Class
