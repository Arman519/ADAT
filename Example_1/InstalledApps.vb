Option Strict Off
Imports System.Collections.ObjectModel
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Text


Public Class InstalledApps
    Private Function RunScript(ByVal scriptText As String) As Object
        Dim MyRunSpace As Runspace = RunspaceFactory.CreateRunspace()
        MyRunSpace.Open()
        Dim MyPipeline As Pipeline = MyRunSpace.CreatePipeline()
        MyPipeline.Commands.AddScript(scriptText)
        Dim results As Collection(Of PSObject) = MyPipeline.Invoke()
        MyRunSpace.Close()
        Dim MyStringBuilder As New StringBuilder()
        For Each obj As PSObject In results
            MyStringBuilder.AppendLine(obj.ToString())
        Next
        Return MyStringBuilder

    End Function
    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Close()
    End Sub

    Private Sub BtnCopyAppList_Click(sender As Object, e As EventArgs) Handles btnCopyAppList.Click
        Dim str As String
        str = InstalledAppsRtb.Text
        Clipboard.SetText(str)
    End Sub
End Class