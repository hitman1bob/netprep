Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
Imports System.Configuration

Module Module1
    Public sContent As Object
    Sub Main()

        'Get the current directory
        Dim currpath As String = Directory.GetCurrentDirectory

        'Set Constants
        Dim DPath As String = My.Settings.DrawingPath
        Dim ToolNo As String = GetToolNo(currpath)
        Dim JobName As String = currpath & "\" & ToolNo & "net.bin"
        Dim DrillReportExtension As String = My.Settings.DrillReportExtension
        Dim DrillReport As String = currpath & "\" & ToolNo & DrillReportExtension
        Dim InputDir As String = currpath & "\input\"
        Dim boolGetPDF As Boolean = True
        Dim ChkRetool As New FileSearch()
        Dim rtResult As DialogResult
        Dim gitCmd As String = My.Settings.gitcmd
        Dim GetDrawing As Boolean = True

        ' HACK Show current variables
        ' Recomment block after debugging
        'Console.WriteLine("DrawingPath = " & DPath) 
        'Console.WriteLine("Current Path = " & currpath)
        'Console.WriteLine("Job Name = " & JobName)
        'Console.WriteLine("Drill Report Name = " & DrillReport)
        'Console.Write("Input Directory = " & InputDir)

        'Begin revision control
        ShellandWait(gitCmd, " init", currpath)     'create git repo
        Try
            My.Computer.FileSystem.CopyFile(My.Settings.gitIgnore, _
                                            currpath & "\" & ".gitignore", _
                                            Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, _
                                            FileIO.UICancelOption.DoNothing)
        Catch ex As Exception
            Console.WriteLine("There was an error copying the .gitignore file")
            Console.ReadKey("Press a key to Exit")
            Exit Sub
        End Try
        ShellandWait(gitCmd, " add .", currpath)    'add all files to list
        ShellandWait(gitCmd, " commit -m ""First Save""", currpath)     'commit all files to repo

        GetDrawing = CheckLocal(currpath, boolGetPDF)

        'Check if job exists
        If File.Exists(JobName) Then
            MessageBox.Show("The Job already exists!  I will open the existing bin file.", "Job Found!", _
                      MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Process.Start(JobName)
            GoTo OpenDrill
        End If

        'Check if job is a retool
        ' Does input directory exist?
        If Directory.Exists(InputDir) Then
            ' Search input directory
            ChkRetool.Search(InputDir, "*.sml;*.lrg;*.sed;*.led;*.f04")
            If ChkRetool.Files.Count > 0 Then
                rtResult = MessageBox.Show("This job may be a retool.  Would you like to continue?", "Film Files Found!", _
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If rtResult = DialogResult.No Then
                    Exit Sub
                End If
            End If
        Else
            ' Missing input directory
            rtResult = MessageBox.Show("The input directory does not exist.  Would you like to continue?", "No Input Direcoty", _
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If rtResult = DialogResult.No Then
                Exit Sub
            End If
        End If
        
        'Find CPL
        Dim cpl As String = GetSS(currpath, "")

        'Load Files
        GetLayers(currpath, cpl)

        'Cleanup
        Dim cmpfile As String = currpath & "\compare.txt"
        Dim fi As New FileInfo(cmpfile)
        Console.WriteLine(fi)
        If fi.Exists Then fi.Delete()

OpenDrill:  'Open Drill Report
        If File.Exists(DrillReport) Then Shell(My.Settings.Editor & " " & DrillReport, AppWinStyle.NormalFocus, False)

        'Open Fab Drawing
        If GetDrawing = True Then GetPDF(ToolNo, DPath)

        'That's all folks!

    End Sub 'Main
    Function GetToolNo(ByVal currpath As String) As String
        Dim myPath As String = currpath
        Dim myDirectory As New System.IO.DirectoryInfo(myPath)
        Dim ToolNo As String = Left(myDirectory.Name, 5)
        Return ToolNo
    End Function  'GetToolNo
    Function GetPDF(ByVal ToolNo As String, ByVal DPath As String) As String

        'DPath = "c:\temp\"   ' HACK Override DrawingPath
        'Show message if drawings directory structure not found
        If DPath = "c:\temp" Then
            GetPDF = ""
            Exit Function
        End If
        If Not Directory.Exists(DPath) Then
            MessageBox.Show("Cannot connect to " & DPath)
            GetPDF = ""
            Exit Function
        End If

        'Setup and perform search

        'Dim strDWG As String = ToolNo & "*.jpg"    ' HACK Override GetPDF search extension
        Dim strDWG As String = ToolNo & "*.pdf"
        Dim fsGpdf As New FileSearch
        fsGpdf.Search(DPath, strDWG)

        'Show message if drawing not found
        If Not fsGpdf.Files.Count > 0 Then
            MessageBox.Show("Drawing not found")
        End If

        'Open each matching file with Adobe Reader
        Dim xFile As FileInfo
        For Each xFile In fsGpdf.Files
            Process.Start(xFile.FullName)
        Next
        GetPDF = ""
    End Function  'GetPDF
    Function GetSS(ByVal currpath As String, ByVal SolderSide As String) As String

        'Setup Regex to Find Solder Side Layer number
        Dim ssregex As String = "^\w*[a-z]-[0-9]+[a-z]{1}\b"
        Dim options As RegexOptions = ((RegexOptions.IgnorePatternWhitespace _
                Or RegexOptions.Multiline) _
                Or RegexOptions.IgnoreCase)
        Dim reg As Regex = New Regex(ssregex, options)

        'Get directory list into sContent
        FileUtils.GetDirList(currpath)

        'Find Solder Side Layer number
        Dim oMatches As MatchCollection
        oMatches = reg.Matches(sContent)
        SolderSide = Val(oMatches.Count)
        Return SolderSide

    End Function 'GetSS
    Function CheckLocal(ByVal currpath As String, ByVal boolGetPDF As Boolean) As Boolean
        Dim myIP() As IPAddress = Dns.GetHostAddresses(Dns.GetHostName())
        Dim counter As Integer = 0
        Do Until counter > myIP.Length - 1
            Dim ipsplit() As String = myIP(counter).ToString.Split(".")
            If ipsplit.Length > 1 Then
                If String.Concat(ipsplit(0), ".", ipsplit(1), ".", ipsplit(2)) = My.Settings.LocalIP Then
                    boolGetPDF = False
                End If
            End If
            counter += 1
        Loop
        Return boolGetPDF
    End Function
    Sub GetLayers(ByVal currpath As String, ByVal cpl As String)

        'Launch CAMMaster
        Shell("c:\Pental~1\CAMMas~1\CAMMas~1.exe", AppWinStyle.MinimizedNoFocus, False)

        Dim btImportJob As New CAMMaster.Tool
        Dim lcount As Integer = 1
        Dim lcolor As Integer = 1
        Dim PlnPfx As String = "P0"
        Dim PlnNum As Integer = 1
        Dim ToolNo As String = GetToolNo(currpath)
        Dim F10Name As String = ToolNo & "net."
        Dim JobName As String = ToolNo & "net"

        'Setup Expression
        Dim lnum As Integer = 1
        Dim layerregex As String = "^\w*[a-z]-" & lnum & "[a-z]{1}\b"
        Dim options As RegexOptions = ((RegexOptions.IgnorePatternWhitespace _
                Or RegexOptions.Multiline) _
                Or RegexOptions.IgnoreCase)
        Dim oRegex As Regex = New Regex(layerregex, options)
        Dim oMatch As Match
        Dim oMatches As MatchCollection


        'get gerber layers
        'kill screen updates in CAMMAster
        btImportJob.MessagesMode = "Background"
        oMatches = oRegex.Matches(sContent)
        If oMatches.Count = 0 And lnum = 1 Then
            Console.WriteLine("No Layer 1 Found")
            Console.WriteLine("Press ENTER to QUIT")
            Console.ReadLine()
            End
        End If

        Do While oMatches.Count = 1
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.ImportX(oMatch.Value, "GERBER", lcount, 0, 0)

                If lcount = 1 Then
                    btImportJob.LayerType(lcount) = "CPU"
                    btImportJob.LayerColor(lcount) = 12
                    If UCase$(Right$(oMatch.Value, 1)) = "N" Or UCase$(Right$(oMatch.Value, 1)) = "A" Then
                        btImportJob.LayerNegative(lcount) = True
                        'OldName = btImportJob.LayerName(lcount)
                        btImportJob.LayerName(lcount) = F10Name + "1f"
                        lcount = lcount + 1
                    Else
                        btImportJob.LayerNegative(lcount) = False
                        btImportJob.LayerColor(lcount) = 12
                        btImportJob.LayerName(lcount) = F10Name + "1"
                        lcount = lcount + 1
                    End If
                ElseIf lcount = cpl Then
                    btImportJob.LayerType(lcount) = "CPL"
                    btImportJob.LayerColor(lcount) = 12
                    If UCase$(Right$(oMatch.Value, 1)) = "N" Or UCase$(Right$(oMatch.Value, 1)) = "A" Then
                        btImportJob.LayerNegative(lcount) = True
                        btImportJob.LayerName(lcount) = F10Name + Format(lcount) + "ssf"
                        lcount = lcount + 1
                    Else
                        btImportJob.LayerNegative(lcount) = False
                        btImportJob.LayerColor(lcount) = 12
                        btImportJob.LayerName(lcount) = F10Name + Format(lcount) + "ss"
                        lcount = lcount + 1
                    End If
                Else
                    Dim PlnFmt As String = Format(PlnNum)
                    If UCase$(Right$(oMatch.Value, 1)) = "N" Or UCase$(Right$(oMatch.Value, 1)) = "A" Then
                        btImportJob.LayerNegative(lcount) = True
                        If PlnNum > 9 Then PlnPfx = "P"
                        btImportJob.LayerType(lcount) = PlnPfx + PlnFmt
                        btImportJob.LayerColor(lcount) = lcolor
                        PlnNum = PlnNum + 1
                        btImportJob.LayerName(lcount) = F10Name + "g" + Format(lcount) + "f"
                        'btImportJob.LayerExtension(lcount) = ".g" + lcount + "f"
                    ElseIf UCase$(Right$(oMatch.Value, 1)) = "P" Or UCase$(Right$(oMatch.Value, 1)) = "M" Then
                        btImportJob.LayerNegative(lcount) = False
                        If PlnNum > 9 Then PlnPfx = "P"
                        btImportJob.LayerType(lcount) = PlnPfx + PlnFmt
                        btImportJob.LayerColor(lcount) = lcolor
                        PlnNum = PlnNum + 1
                        btImportJob.LayerName(lcount) = F10Name + "g" + Format(lcount) + "d"
                        'btImportJob.LayerExtension(lcount) = ".g" + lcount + "d"
                    Else
                        btImportJob.LayerNegative(lcount) = False
                        btImportJob.LayerType(lcount) = "SIG"
                        btImportJob.LayerName(lcount) = F10Name + Format(lcount)
                        'btImportJob.LayerExtension(lcount) = "." + lcount
                        btImportJob.LayerColor(lcount) = lcolor

                    End If
                    lcount = lcount + 1
                End If



                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
            lnum = lnum + 1
            lcolor = lcolor + 1
            layerregex = "^\w*[a-z]-" & lnum & "[a-z]{1}\b"
            oRegex = New Regex(layerregex, options)
            oMatches = oRegex.Matches(sContent)
        Loop

        'get m1 files
        layerregex = "^\w*[a-z]-sm1{1}\b"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        For Each oMatch In oMatches
            btImportJob.CurrentDirectory = (currpath)
            btImportJob.ImportX(oMatch.Value, "GERBER", lcount, 0, 0)
            btImportJob.LayerType(lcount) = "MSU"
            btImportJob.LayerColor(lcount) = 9
            btImportJob.LayerName(lcount) = F10Name + "m1"
            lcount = lcount + 1

            'Remove after debugging
            'Console.WriteLine(oMatch.Value)

        Next

        'get et-cvl-drl1 files
        layerregex = "^\w*[a-z]-et-cvrly-drl1"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        For Each oMatch In oMatches
            btImportJob.CurrentDirectory = (currpath)
            btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
            btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
            btImportJob.LayerType(lcount) = "MSU"
            btImportJob.LayerColor(lcount) = 9
            btImportJob.LayerName(lcount) = F10Name + "m1"
            lcount = lcount + 1

            'Remove after debugging
            'Console.WriteLine(oMatch.Value)

        Next
        'get cvl-drl1 files
        layerregex = "^\w*[a-z]-cvrly-drl1"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        For Each oMatch In oMatches
            btImportJob.CurrentDirectory = (currpath)
            btImportJob.ImportX(oMatch.Value, "GERBER", lcount, 0, 0)
            btImportJob.LayerType(lcount) = "MSU"
            btImportJob.LayerColor(lcount) = 9
            btImportJob.LayerName(lcount) = F10Name + "m1"
            lcount = lcount + 1

            'Remove after debugging
            'Console.WriteLine(oMatch.Value)

        Next

        'get msl
        layerregex = "^\w*[a-z]-sm" & (lnum - 1) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        For Each oMatch In oMatches
            btImportJob.CurrentDirectory = (currpath)
            btImportJob.ImportX(oMatch.Value, "GERBER", lcount, 0, 0)
            btImportJob.LayerType(lcount) = "MSL"
            btImportJob.LayerColor(lcount) = 9
            btImportJob.LayerName(lcount) = F10Name + "m" + cpl
            lcount = lcount + 1

            'Remove after debugging
            'Console.WriteLine(oMatch.Value)

        Next

        'get et-cvl-drl? files
        layerregex = "^\w*[a-z]-et-cvrly-drl" & (lnum - 1) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        For Each oMatch In oMatches
            btImportJob.CurrentDirectory = (currpath)
            btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
            btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
            btImportJob.LayerType(lcount) = "MSL"
            btImportJob.LayerName(lcount) = F10Name + "m" + Format(lcount)

            lcount = lcount + 1

            'Remove after debugging
            'Console.WriteLine(oMatch.Value)

        Next

        'get cvl-drl? files
        layerregex = "^\w*[a-z]-cvrly-drl" & (lnum - 1) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        For Each oMatch In oMatches
            btImportJob.CurrentDirectory = (currpath)
            btImportJob.ImportX(oMatch.Value, "GERBER", lcount, 0, 0)
            btImportJob.LayerType(lcount) = "MSL"
            btImportJob.LayerColor(lcount) = 9
            btImportJob.LayerName(lcount) = F10Name + "m" + Format(lcount)
            lcount = lcount + 1

            'Remove after debugging
            'Console.WriteLine(oMatch.Value)

        Next

        'get s1
        layerregex = "^\w*[a-z]-ss1{1}\b"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        For Each oMatch In oMatches
            btImportJob.CurrentDirectory = (currpath)
            btImportJob.ImportX(oMatch.Value, "GERBER", lcount, 0, 0)
            btImportJob.LayerColor(lcount) = 8
            btImportJob.LayerType(lcount) = "SKU"
            btImportJob.LayerName(lcount) = F10Name + "s1"
            lcount = lcount + 1

            'Remove after debugging
            'Console.WriteLine(oMatch.Value)

        Next

        'get skl
        layerregex = "^\w*[a-z]-ss" & (lnum - 1) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        For Each oMatch In oMatches
            btImportJob.CurrentDirectory = (currpath)
            btImportJob.ImportX(oMatch.Value, "GERBER", lcount, 0, 0)
            btImportJob.LayerColor(lcount) = 8
            btImportJob.LayerType(lcount) = "SKL"
            btImportJob.LayerName(lcount) = F10Name + "s" + cpl
            lcount = lcount + 1

            'Remove after debugging
            'Console.WriteLine(oMatch.Value)

        Next

        'get map
        layerregex = "^\w*[a-z]?-?\w*[a-z]-map"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        For Each oMatch In oMatches
            btImportJob.CurrentDirectory = (currpath)
            btImportJob.ImportX(oMatch.Value, "GERBER", lcount, 0, 0)
            btImportJob.LayerColor(lcount) = 5
            btImportJob.LayerType(lcount) = "BOL"
            btImportJob.LayerName(lcount) = F10Name + "oln"
            lcount = lcount + 1

            'Remove after debugging
            'Console.WriteLine(oMatch.Value)

        Next

        'get et-bv files
        layerregex = "^\w*[a-z]-et-bvia[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim et_bvcnt As Integer = oMatches.Count
        If et_bvcnt = 0 Then
            GoTo getbv
        End If
        Do While et_bvcnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "VID"

                'Dim message, title, defaultValue As String
                'Dim myValue As Object
                'message = "Please name the extension"   ' Set prompt.
                'title = "Rename Blind Via File " + oMatch.Value   ' Set title.
                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 13) ' Set default value.

                ' Display message, title, and default value.
                'myValue = InputBox(message, title, defaultValue)
                'myValue = defaultValue
                btImportJob.LayerName(lcount) = F10Name + "bv" + myValue

                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                lcount = lcount + 1
                et_bvcnt = et_bvcnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop
        GoTo getetlsr

getbv:  'get bv files
        layerregex = "^\w*[a-z]-bvia[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim bvcnt As Integer = oMatches.Count
        Do While bvcnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "VID"

                'Dim message, title, defaultValue As String
                'Dim myValue As Object
                'message = "Please name the extension"   ' Set prompt.
                'title = "Rename Blind Via File " + oMatch.Value   ' Set title.
                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 10) ' Set default value.

                ' Display message, title, and default value.
                'myValue = InputBox(message, title, defaultValue)
                'myValue = defaultValue
                btImportJob.LayerName(lcount) = F10Name + "bv" + myValue

                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                lcount = lcount + 1
                bvcnt = bvcnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop


getetlsr:  'get et-lsr files
        layerregex = "^\w*[a-z]-et-lsr[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim et_lsrcnt As Integer = oMatches.Count
        If et_lsrcnt = 0 Then
            GoTo getlsr
        End If
        Do While et_lsrcnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "VID"

                'Dim message, title, defaultValue As String
                'Dim myValue As Object
                'message = "Please name the extension"   ' Set prompt.
                'title = "Rename Blind Via File " + oMatch.Value   ' Set title.
                Dim MyValue As Object = Right$(oMatch.Value, oMatch.Length - 12)   ' Set default value.

                ' Display message, title, and default value.
                'myValue = InputBox(message, title, defaultValue)
                'myValue = defaultValue
                btImportJob.LayerName(lcount) = F10Name + "bv" + MyValue
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                lcount = lcount + 1
                et_lsrcnt = et_lsrcnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop
        GoTo getetcsink


getlsr:  'get lsr files
        layerregex = "^\w*[a-z]-lsr[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim lsrcnt As Integer = oMatches.Count
        Do While lsrcnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "VID"

                'Dim message, title, defaultValue As String
                'Dim myValue As Object
                'message = "Please name the extension"   ' Set prompt.
                'title = "Rename Blind Via File " + oMatch.Value   ' Set title.
                Dim MyValue As Object = Right$(oMatch.Value, oMatch.Length - 9)   ' Set default value.

                ' Display message, title, and default value.
                'myValue = InputBox(message, title, defaultValue)
                'myValue = defaultValue
                btImportJob.LayerName(lcount) = F10Name + "bv" + MyValue
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                lcount = lcount + 1
                lsrcnt = lsrcnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop

getetcsink:  'get et-csink files
        layerregex = "^\w*[a-z]-et-csink[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim et_csinkcnt As Integer = oMatches.Count
        If et_csinkcnt = 0 Then
            GoTo getcsink
        End If
        Do While et_csinkcnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                'btImportJob.ImportX(oMatch.Value, "GERBER", lcount, 0, 0)
                'btImportJob.LayerType(lcount) = "VID"
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "DRI"

                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 9)   ' Set default value.
                btImportJob.LayerName(lcount) = F10Name + myValue

                lcount = lcount + 1
                et_csinkcnt = et_csinkcnt - 1

                'Remove after debugging
                Console.WriteLine(oMatch.Value)

            Next
        Loop
        GoTo getetcbore

getcsink:  'get csink files
        layerregex = "^\w*[a-z]-csink[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim csinkcnt As Integer = oMatches.Count
        Do While csinkcnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                'btImportJob.ImportX(oMatch.Value, "GERBER", lcount, 0, 0)
                'btImportJob.LayerType(lcount) = "VID"
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "DRI"

                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 6)   ' Set default value.
                btImportJob.LayerName(lcount) = F10Name + myValue

                lcount = lcount + 1
                csinkcnt = csinkcnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop

getetcbore:  'get et-cbore files
        layerregex = "^\w*[a-z]-et-cbore[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim et_cborecnt As Integer = oMatches.Count
        If et_cborecnt = 0 Then
            GoTo getcbore
        End If
        Do While et_cborecnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                'btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                'btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                'btImportJob.LayerType(lcount) = "VID"
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "DRI"

                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 6)   ' Set default value.
                btImportJob.LayerName(lcount) = F10Name + myValue

                lcount = lcount + 1
                et_cborecnt = et_cborecnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop
        GoTo getetpth

getcbore:  'get cbore files
        layerregex = "^\w*[a-z]-cbore[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim cborecnt As Integer = oMatches.Count
        Do While cborecnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                'btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                'btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                'btImportJob.LayerType(lcount) = "VID"
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "DRI"

                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 6)   ' Set default value.
                btImportJob.LayerName(lcount) = F10Name + myValue

                lcount = lcount + 1
                cborecnt = cborecnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop

getetpth:  'get et-pth file
        layerregex = "^\w*[a-z]-et-pth"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim et_pcnt As Integer = oMatches.Count
        If et_pcnt = 0 Then
            GoTo getdri_p
        End If
        Do While et_pcnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerColor(lcount) = 10
                btImportJob.LayerType(lcount) = "DRI"
                btImportJob.LayerName(lcount) = F10Name + "dri"
                lcount = lcount + 1
                et_pcnt = et_pcnt - 1

                'Remove after debugging
                'Console.WriteLine("drl_p_" & oMatch.Value)

            Next
        Loop
        GoTo getetnpt

getdri_p:  'get dri_p file
        layerregex = "^\w*[a-z]-drl_p"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim drl_pcnt As Integer = oMatches.Count
        Do While drl_pcnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerColor(lcount) = 10
                btImportJob.LayerType(lcount) = "DRI"
                btImportJob.LayerName(lcount) = F10Name + "dri"
                lcount = lcount + 1
                drl_pcnt = drl_pcnt - 1

                'Remove after debugging
                'Console.WriteLine("drl_p_" & oMatch.Value)
            Next
        Loop

getetnpt:  'get et-npt file
        layerregex = "^\w*[a-z]-et-npt"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim et_ncnt As Integer = oMatches.Count
        If et_ncnt = 0 Then
            GoTo getdri_n
        End If
        Do While et_ncnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerColor(lcount) = 10
                btImportJob.LayerType(lcount) = "DRI"
                btImportJob.LayerName(lcount) = F10Name + "npt"
                ConvertNPTH()
                lcount = lcount + 1
                et_ncnt = et_ncnt - 1

                'Remove after debugging
                'Console.WriteLine("drl_n_" & oMatch.Value)
            Next
        Loop
        GoTo getet2nd

getdri_n:  'get dri_n file
        layerregex = "^\w*[a-z]-drl_n"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim drl_ncnt As Integer = oMatches.Count
        Do While drl_ncnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerColor(lcount) = 10
                btImportJob.LayerType(lcount) = "DRI"
                btImportJob.LayerName(lcount) = F10Name + "npt"
                ConvertNPTH()
                lcount = lcount + 1
                drl_ncnt = drl_ncnt - 1

                'Remove after debugging
                'Console.WriteLine("drl_n_" & oMatch.Value)

            Next
        Loop

getet2nd:  'get et-2nd_drill file
        layerregex = "^\w*[a-z]-et-2nd-drill"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim et_secdrl As Integer = oMatches.Count
        If et_secdrl = 0 Then
            GoTo get2nd
        End If
        Do While et_secdrl > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "DRI"
                btImportJob.LayerName(lcount) = F10Name + "npt"
                ConvertNPTH()
                lcount = lcount + 1
                et_secdrl = et_secdrl - 1

                'Remove after debugging
                'Console.WriteLine("drl_n_" & oMatch.Value)

            Next
        Loop
        GoTo getetepxy

get2nd:  'get 2nd_drill file
        layerregex = "^\w*[a-z]-2nd-drill"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim secdrl As Integer = oMatches.Count
        Do While secdrl > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "DRI"
                btImportJob.LayerName(lcount) = F10Name + "npt"
                ConvertNPTH()
                lcount = lcount + 1
                secdrl = secdrl - 1

                'Remove after debugging
                'Console.WriteLine("drl_n_" & oMatch.Value)

            Next
        Loop


getetepxy:  'get et-epxy files
        layerregex = "^\w*[a-z]-et-epxy[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim et_epxycnt As Integer = oMatches.Count
        If et_epxycnt = 0 Then
            GoTo getepxy
        End If
        Do While et_epxycnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "VID"

                'Dim message, title, defaultValue As String
                'Dim myValue As Object
                'message = "Please name the extension"   ' Set prompt.
                'title = "Rename Blind Via File " + oMatch.Value   ' Set title.
                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 10)   ' Set default value.

                ' Display message, title, and default value.
                'myValue = InputBox(message, title, defaultValue)
                'myValue = defaultValue
                btImportJob.LayerName(lcount) = F10Name + "epxy" + myValue
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                lcount = lcount + 1
                et_epxycnt = et_epxycnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop
        GoTo getetcu

getepxy:  'get epxy files
        layerregex = "^\w*[a-z]-epxy[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim epxycnt As Integer = oMatches.Count
        Do While epxycnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "VID"

                'Dim message, title, defaultValue As String
                'Dim myValue As Object
                'message = "Please name the extension"   ' Set prompt.
                'title = "Rename Blind Via File " + oMatch.Value   ' Set title.
                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 10)   ' Set default value.

                ' Display message, title, and default value.
                'myValue = InputBox(message, title, defaultValue)
                'myValue = defaultValue
                btImportJob.LayerName(lcount) = F10Name + "epxy" + myValue
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                lcount = lcount + 1
                epxycnt = epxycnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop

getetcu:  'get et-cu files
        layerregex = "^\w*[a-z]-et-cu[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim et_cucnt As Integer = oMatches.Count
        If et_cucnt = 0 Then
            GoTo getcu
        End If
        Do While et_cucnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "VID"

                'Dim message, title, defaultValue As String
                'Dim myValue As Object
                'message = "Please name the extension"   ' Set prompt.
                'title = "Rename Blind Via File " + oMatch.Value   ' Set title.
                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 10)   ' Set default value.

                ' Display message, title, and default value.
                'myValue = InputBox(message, title, defaultValue)
                'myValue = defaultValue
                btImportJob.LayerName(lcount) = F10Name + "cu" + myValue
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                lcount = lcount + 1
                et_cucnt = et_cucnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop
        GoTo getetslvr

getcu:  'get cu files
        layerregex = "^\w*[a-z]-cu[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim cucnt As Integer = oMatches.Count
        Do While cucnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "VID"

                'Dim message, title, defaultValue As String
                'Dim myValue As Object
                'message = "Please name the extension"   ' Set prompt.
                'title = "Rename Blind Via File " + oMatch.Value   ' Set title.
                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 10)   ' Set default value.

                ' Display message, title, and default value.
                'myValue = InputBox(message, title, defaultValue)
                'myValue = defaultValue
                btImportJob.LayerName(lcount) = F10Name + "cu" + myValue
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                lcount = lcount + 1
                cucnt = cucnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop


getetslvr:  'get et-slvr files
        layerregex = "^\w*[a-z]-et-slvr[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim et_slvrcnt As Integer = oMatches.Count
        If et_slvrcnt = 0 Then
            GoTo getslvr
        End If
        Do While et_slvrcnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "VID"

                'Dim message, title, defaultValue As String
                'Dim myValue As Object
                'message = "Please name the extension"   ' Set prompt.
                'title = "Rename Blind Via File " + oMatch.Value   ' Set title.
                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 10)   ' Set default value.

                ' Display message, title, and default value.
                'myValue = InputBox(message, title, defaultValue)
                'myValue = defaultValue
                btImportJob.LayerName(lcount) = F10Name + "slvr" + myValue
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                lcount = lcount + 1
                et_slvrcnt = et_slvrcnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop
        GoTo layercount

getslvr:  'get slvr files
        layerregex = "^\w*[a-z]-slvr[1-9][0-9]?-[1-9][0-9]?"
        oRegex = New Regex(layerregex, options)
        oMatches = oRegex.Matches(sContent)
        Dim slvrcnt As Integer = oMatches.Count
        Do While slvrcnt > 0
            For Each oMatch In oMatches
                btImportJob.CurrentDirectory = (currpath)
                btImportJob.DrillDataFormat = "Left=2, Right=4, Absolute, English, ASCII, Trailing, Excellon"
                btImportJob.ImportX(oMatch.Value, "Drill", lcount, 0, 0)
                btImportJob.LayerType(lcount) = "VID"

                'Dim message, title, defaultValue As String
                'Dim myValue As Object
                'message = "Please name the extension"   ' Set prompt.
                'title = "Rename Blind Via File " + oMatch.Value   ' Set title.
                Dim myValue As Object = Right$(oMatch.Value, oMatch.Length - 10)   ' Set default value.

                ' Display message, title, and default value.
                'myValue = InputBox(message, title, defaultValue)
                'myValue = defaultValue
                btImportJob.LayerName(lcount) = F10Name + "slvr" + myValue
                'btImportJob.LayerName(lcount) = F10Name + Format(oMatch)
                lcount = lcount + 1
                slvrcnt = slvrcnt - 1

                'Remove after debugging
                'Console.WriteLine(oMatch.Value)

            Next
        Loop

layercount:
        lcount = lcount + 1

        'tidy up and save both bin files
        btImportJob.OnlyCurrentLayer = False
        btImportJob.MessagesMode = "Interactive"
        btImportJob.SaveJob(currpath & "\" & ToolNo & ".bin")
        btImportJob.SaveJob(currpath & "\" & JobName & ".bin")

    End Sub  'GetLayers
    Sub ConvertNPTH()
        Dim highlightnpth As New CAMMaster.Tool
        With highlightnpth

            .OnlyCurrentDCode = False
            .OnlyCurrentLayer = True
            .SelectGlobal()
            Dim drilltools, dtc, dtc1, drilltcode, toolnum, Len1, ds1, dt1, ds2, dt2 As String

            While .GetSelectionTools("") <> ""    '--While the selected dcodes is something
                drilltools = .GetSelectionTools("")   '-------get the selected dcodes
                dtc = Len(drilltools)     '--Get the length of the returned string of dcodes
                dtc1 = InStr(drilltools, ",")   '--Find the location of the comma in the string

                If dtc1 = 0 Then     '--If there is no comma than this is the last dcode
                    drilltcode = drilltools
                Else
                    drilltcode = Left$(drilltools, InStr(drilltools, ",") - 1) '--If there is a comma trim whatever is in front of it from the left
                End If

                toolnum = Right$(.DCodeShape(drilltcode), Len(.DCodeShape(drilltcode)) - InStr(.DCodeShape(drilltcode), "=")) '--Parse the toolnum
                Len1 = Len(.ToolShape(toolnum))
                ds1 = InStrRev(.ToolShape(toolnum), ",")
                dt1 = Len1 - ds1
                ds2 = Left$(.ToolShape(toolnum), ds1)   '--Parse the drillsize
                dt2 = Right$(.ToolShape(toolnum), dt1)   '--Parse the drilltype
                .ToolShape(toolnum) = ds2 & " Non-plated"
                .OnlyCurrentDCode = True    '--Remove the dcode from the selection
                .CurrentDCode = drilltcode
                .Select("Remove", -1000, -1000, 1000, 1000)
                .OnlyCurrentDCode = False
            End While
        End With
    End Sub 'ConvertNPTH
    Public Sub ShellandWait(ByVal ProcessPath As String, ByVal param As String, ByVal currdir As String)
        Dim objProcess As System.Diagnostics.Process
        Try
            objProcess = New System.Diagnostics.Process()
            objProcess.StartInfo.WorkingDirectory = currdir
            objProcess.StartInfo.FileName = ProcessPath
            objProcess.StartInfo.Arguments = param
            objProcess.StartInfo.CreateNoWindow = False
            objProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
            objProcess.Start()

            'Wait until the process passes back an exit code 
            objProcess.WaitForExit()

            'Free resources associated with this process
            objProcess.Close()
        Catch
            Console.WriteLine("Could not start process " & ProcessPath)
        End Try
    End Sub 'ShellandWait
End Module