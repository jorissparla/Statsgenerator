Imports System.Net

Imports System.IO

Imports System.Diagnostics
Imports System.Configuration
Imports System.Configuration.ClientSettingsSection


Module Module1

    Function GetPage(ByVal pageUrl As String) As String

        Dim s As String = ""
        Dim gPwd As String = ""

        Try
            Dim userid As String = System.Configuration.ConfigurationManager.AppSettings("userid")
            Dim password As String = System.Configuration.ConfigurationManager.AppSettings("password")
            Dim timeoutval As Integer = System.Configuration.ConfigurationManager.AppSettings("timeout")
            '  gPwd = Crypto.Encrypt("******", userid)

            password = Crypto.Decrypt(password, userid)
            Dim request As HttpWebRequest = WebRequest.Create(pageUrl)
            request.Credentials = New NetworkCredential(userid, password, "infor")
            request.Timeout = timeoutval

            Dim response As HttpWebResponse = request.GetResponse()

            Using reader As StreamReader = New StreamReader(response.GetResponseStream())

                s = reader.ReadToEnd()

            End Using

        Catch ex As Exception

            Debug.WriteLine("FAIL: " + ex.Message)
            Throw New Exception(ex.Message)
        End Try
        Return s

    End Function



    Function ExtractBody(ByVal page As String) As String

        Return System.Text.RegularExpressions.Regex.Replace(page, ".*<body[^>]*>(.*)</body>.*", "$1", System.Text.RegularExpressions.RegexOptions.IgnoreCase)

    End Function


    Function MakeNavigatorQueryString(ByVal strFromDate As String, ByVal strToDate As String, ByVal strOwner As String, ByVal strOwnerGroups As String, ByVal strRegion As String)
        Dim qs As String

        Dim templateString = "http://navigator.infor.com/A/stats.asp?Requery=Requery&ClearAll=1&StartDate={0}+00:00:00+AM&EndDate={1}+PM&Reports=111110110"
        qs = String.Format(templateString, strFromDate, strToDate)

        If strOwner <> "" Then
            qs = qs + String.Format("&Owners={0}", strOwner)
        End If
        If strOwnerGroups <> "" Then
            qs = qs + String.Format("&Groups={0}", strOwnerGroups)

        End If
        If strRegion <> "" Then
            qs = qs + String.Format("&RepRegions={0}", strRegion)
        End If


        ' Groups=380,463,478,480,481&RepRegions=2
        'Clipboard.SetText(qs)
        Return qs
    End Function

    Function MakeNavigatorQueryString(ByVal PageURL As String, ByVal strFromDate As String, ByVal strToDate As String, ByVal strOwner As String, ByVal strOwnerGroups As String, ByVal strRegion As String)
        Dim qs As String

        ''Dim templateString = PageURL '"http://navigator.infor.com/A/stats.asp?Requery=Requery&ClearAll=1&StartDate={0}+00:00:00+AM&EndDate={1}+11:59:59+PM&Reports=111110110"
        Dim templateString = PageURL '"http://navigator.infor.com/A/stats.asp?Requery=Requery&ClearAll=1&StartDate={0}+00:00:00+AM&EndDate={1}+PM&Reports=111110110"
        qs = String.Format(templateString, strFromDate, strToDate)

        If strOwner <> "" Then
            qs = qs + String.Format("&Owners={0}", strOwner)
        End If
        If strOwnerGroups <> "" Then
            qs = qs + String.Format("&Groups={0}", strOwnerGroups)

        End If
        If strRegion <> "" Then
            qs = qs + String.Format("&RepRegions={0}", strRegion)
        End If


        ' Groups=380,463,478,480,481&RepRegions=2
        Clipboard.SetText(qs)
        Return qs
    End Function

    'Sub Main()
    '    'Runquery("9/1/2010", "9/21/2010", "797466")
    'End Sub


    Public Sub DeleteQSetResults(ByVal Owner As Rep, ByVal aRange As Range, ByVal anOwnerGroup As OwnerGroup, ByVal aRegion As Region)

    End Sub

    Public Function RunQuery(ByVal MyQuerySet As QuerySet, ByVal selPage As Page)
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim bStartLooking As Boolean = False
 
        Dim strOwner As String = ""
        Dim strOG As String = ""
        Dim strRegion As String = ""
        Dim strUserName As String = ""
        Return RunqueryEx(MyQuerySet, selPage)
    End Function
    Public Function RunqueryEx(ByVal MyQuerySet As QuerySet, ByVal selPage As Page) As Integer

        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim bStartLooking As Boolean = False
        Dim strFromDate As String = MyQuerySet.Range.FromDate
        Dim strToDate As String = MyQuerySet.Range.toDate
        Dim strOwner As String = ""
        Dim strOG As String = ""
        Dim strOGName As String = ""

        Dim strRegion As String = ""
        Dim strRegionName As String = ""
        Dim strUserName As String = ""


        If Not IsNothing(MyQuerySet.Owner) Then
            strOwner = MyQuerySet.Owner.UserId
            strUserName = MyQuerySet.Owner.Name
        End If
        If Not IsNothing(MyQuerySet.OwnerGroup) Then
            strOG = MyQuerySet.OwnerGroup.IDString
            strOGName = MyQuerySet.OwnerGroup.Name
        End If
        If Not IsNothing(MyQuerySet.Region.Name) Then
            strRegionName = MyQuerySet.Region.Name
            strRegion = MyQuerySet.Region.IDSTring
        End If
        Dim qSet As QuerySet = MyQuerySet

        '  Dim qSet As New QuerySet(strOwner, aRange.Name, strOwner, strOG, strRegion)


        Dim dbDatabase As New SqlClient.SqlConnection
        dbDatabase.ConnectionString = strDbConnect
        '    Dim oKPIs() As KPI = GetKPIs()
        Dim KPIs As List(Of KPI) = New DBKPIs(selPage.PageID).List

        Dim page As String = ""
        Dim bFirstValueFound = False

        i = 0
        Try
            page = GetPage(MakeNavigatorQueryString(selPage.PageURL, strFromDate, strToDate, strOwner, strOG, strRegion)) '"http://navigator.infor.com/A/stats.asp?Requery=Requery&ClearAll=1&StartDate=9/1/2010+00:00:00+AM&EndDate=9/19/2010+11:59:59+PM&Reports=111110110&Owners=797466&Groups=380,463,478,480,481")

            ' Dim body As String = ExtractBody(page)
            Dim lines As String() = page.Split(vbLf)
            For Each line As String In lines

                i = i + 1
                If line.IndexOf(selPage.StartString) Then bStartLooking = True
                If bStartLooking Then
                    For Each aKPI In KPIs
                        If aKPI.Value = "0.0" Then
                            If IsNothing(aKPI.OwnerField) Then
                                Dim sTemp As String = GetTagValue(line, aKPI.LookupString)
                                If sTemp <> "NOTFOUND" Then
                                    aKPI.Value = sTemp
                                    'bStartLooking = False
                                    'Next
                                    'Exit For
                                End If
                            Else
                                Dim sTemp As String = GetTagSpecialValue(line, aKPI.OwnerField.LookupString)
                                If sTemp <> "NOTFOUND" Then
                                    aKPI.Value = sTemp
                                End If
                            End If
                        End If
                    Next
                End If
            Next line

            '  Debug.WriteLine(body)
            Dim RunDate As String = Format(Now(), "MM/dd/yyyy hh:m")


            i = 0
            WriteMessage(String.Format("Storing results for {0}:{1}:{2}:{3}", strUserName, MyQuerySet.Range.Name, strOGName, strRegionName))
            For Each aKPI In KPIs
                Dim myDBKPI As New KPIQuerySetResultDB(New KPIQuerySet(aKPI, qSet))

                If aKPI.Name <> "NOTUSED" Then
                    Try
                        myDBKPI.SaveToDB()
                    Catch ex As Exception
                        WriteMessage(String.Format("Error: {0} \n", ex.Message))
                    End Try
                    WriteMessage(String.Format("{0} = {1}", aKPI.Name, aKPI.Value))
                End If
            Next
            dbDatabase.Close()
            '   Dim strUpdateString As String
            '   strUpdateString = " INSERT INTO KPIValues VALUES(KPIName={0}, Value={1}, OwnerList={2})"
            Return 0
        Catch ex1 As Exception
            WriteMessage(ex1.Message)
            Return -1
        End Try
    End Function

    Public Function Runquery(ByVal Owner As Rep, ByVal aRange As Range, ByVal anOwnerGroup As OwnerGroup, ByVal aRegion As Region, ByVal selPage As Page) As Integer

        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim bStartLooking As Boolean = False
        Dim strFromDate As String = aRange.FromDate
        Dim strToDate As String = aRange.toDate
        Dim strOwner As String = ""
        Dim strOG As String = ""
        Dim strRegion As String = ""
        Dim strUserName As String = ""


        If IsNothing(Owner) Then
            strOwner = ""
        Else
            strOwner = Owner.UserId
            strUserName = Owner.Name
        End If
        If IsNothing(anOwnerGroup) Then
            strOG = ""
        Else
            strOG = anOwnerGroup.Name
        End If
        If IsNothing(aRegion) Then
            strRegion = ""
        Else
            strRegion = aRegion.Name
        End If

        Dim qSet As New QuerySet(strUserName, aRange.Name, strOwner, strOG, strRegion)

        Runquery = RunqueryEx(qSet, selPage)
    End Function






    Function GetTagValue(ByVal line As String, ByVal strTagname As String) As String
        Dim StrSearch As String
        StrSearch = strTagname
        If line.IndexOf(StrSearch) > 0 Then
            If StrSearch = "TITLE=" + Chr(34) + "5/87" + Chr(34) + ">" Then
                Dim x As Integer = 0
            End If
            Dim idx As Integer = line.IndexOf(StrSearch)
            Dim stemp As String = Mid(line, idx + Len(StrSearch) + 1)
            idx = stemp.IndexOf("<")
            stemp = Left(stemp, idx)
            If Right(stemp, 1) = "%" Then
                stemp = Mid(stemp, 1, Len(stemp) - 1)
            End If
            Return stemp
        Else
            Return "NOTFOUND"
        End If
    End Function


    Function GetTagValue2(ByVal line As String, ByVal strTagname As String) As String
        Dim StrSearch As String
        StrSearch = strTagname
        If line.IndexOf(StrSearch) > 0 Then
            If StrSearch = "TITLE=" + Chr(34) + "5/87" + Chr(34) + ">" Then
                Dim x As Integer = 0
            End If
            Dim idx As Integer = line.IndexOf(StrSearch)
            Dim stemp As String = Mid(line, idx + Len(StrSearch) + 1)
            idx = stemp.IndexOf("<")
            stemp = Left(stemp, idx)
            If Right(stemp, 1) = "%" Then
                stemp = Mid(stemp, 1, Len(stemp) - 1)
            End If

            Return stemp
        Else
            Return "NOTFOUND"
        End If
    End Function

    Function GetTagSpecialValue(ByVal line As String, ByVal strTagname As String) As String
        Dim StrSearch As String
        StrSearch = strTagname
        If line.IndexOf(StrSearch) > 0 Then
           
            Dim idx As Integer = line.IndexOf(StrSearch)
            Dim stemp As String = Mid(line, idx + Len(StrSearch) + 1)
            For i = 1 To 5
                idx = stemp.IndexOf(">")
                stemp = Mid(stemp, idx + 2)
            Next

            idx = stemp.IndexOf("<")
            stemp = Left(stemp, idx)
            If Right(stemp, 1) = "%" Then
                stemp = Mid(stemp, 1, Len(stemp) - 1)
            End If
            Return stemp
        Else
            Return "NOTFOUND"
        End If
    End Function




    Public Sub ChangePassWord()
        Dim userid As String = System.Configuration.ConfigurationManager.AppSettings("userid")
        Dim password As String = InputBox("Enter password") '= System.Configuration.ConfigurationManager.AppSettings("password")

        password = Crypto.Encrypt(password, userid)
        WriteMessage(password)
        Clipboard.SetText(password)
        Console.ReadLine()
        System.Configuration.ConfigurationManager.AppSettings("password") = password
    End Sub
End Module

