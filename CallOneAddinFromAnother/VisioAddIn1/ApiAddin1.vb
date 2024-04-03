Imports System.Runtime.InteropServices
Imports System.Windows.Forms

<ComVisible(True)>
Public Class ApiAddin1

    Public Sub SayHello(name As String)
        ' access your adding using Globals.ThisAddIn
        MessageBox.Show($"Hello {name}", "Addin1")
    End Sub

End Class
