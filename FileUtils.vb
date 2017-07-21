Imports System.IO
Imports System.Text.RegularExpressions
Public Class FileUtils
    Private Shared Function WriteToTempFile(ByVal Data As String) As String
        ' Writes text to a temporary file and returns path
        Dim strFilename As String = Path.GetTempFileName()
        Dim objFS As New System.IO.FileStream(strFilename, _
            System.IO.FileMode.Append, _
            System.IO.FileAccess.Write)
        ' Opens stream and begins writing
        Dim Writer As New StreamWriter(objFS)
        Writer.BaseStream.Seek(0, SeekOrigin.End)
        Writer.WriteLine(Data)
        Writer.Flush()
        ' Closes and returns temp path
        Writer.Close()
        Return strFilename
    End Function  'WriteToTempFile
    Public Shared Function GetDirList(ByVal currpath As String) As Object
        Dim oFile As File
        Dim oReader As StreamReader
        Dim oFileStream As FileStream
        Dim oWriter As StreamWriter
        Dim strFilename As String = WriteToTempFile("")
        Dim file As String
        Dim stripreg As String = "^.*\\"
        Dim stripoptions As RegexOptions = ((RegexOptions.IgnorePatternWhitespace _
                Or RegexOptions.Multiline) _
                Or RegexOptions.IgnoreCase)
        Dim stripregex As Regex = New Regex(stripreg, stripoptions)

        oFileStream = New FileStream(strFilename, FileMode.Create)
        oWriter = New StreamWriter(oFileStream)

        For Each file In Directory.GetFiles(currpath)
            file = Regex.Replace(file, stripreg, "")
            oWriter.WriteLine(file)
        Next
        oWriter.Close()
        oFileStream.Close()
        oReader = IO.File.OpenText(strFilename)
        sContent = oReader.ReadToEnd
        oReader.Close()
        oReader = Nothing
        oFile = Nothing
        Return sContent
    End Function  'GetDirList
End Class
