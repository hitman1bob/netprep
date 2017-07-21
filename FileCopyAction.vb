Imports System.Threading

Public Class FileCopyAction
    Private fromPathAndName As String
    Private toPathAndName As String
    Private errorException As Exception
    Private completed As Boolean
    Private mutextToNotify As Mutex

    Public copyThread As Thread

    Public Sub New(ByVal fromStr As String, ByVal toStr As String)
        completed = False
        fromPathAndName = fromStr
        toPathAndName = toStr
        copyThread = New Thread(AddressOf copying)
        copyThread.Start()
        Thread.Sleep(0)
    End Sub

    Private Sub copying()
        Try
            System.IO.File.Copy(fromPathAndName, toPathAndName)
        Catch ex As Exception
            errorException = ex
        End Try
        completed = True
    End Sub
    Public Sub Dispose()
        fromPathAndName = Nothing
        toPathAndName = Nothing
        errorException = Nothing
        copyThread = Nothing
    End Sub

    Public Function hasErrored() As Boolean
        If errorException Is Nothing Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Function Ex() As Exception
        Return errorException
    End Function
    Public Function hasCompleted() As Boolean
        Return completed
    End Function
End Class
