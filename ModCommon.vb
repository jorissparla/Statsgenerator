Imports System.Configuration
Imports System.Configuration.ClientSettingsSection

Module ModCommon
    Public Enum PageType
        ALLSTATS = 0
        SURVEYS = 1
        TRANSFERRED = 2
    End Enum
    Public NullOwnerGroup As OwnerGroup = New OwnerGroup("", "")
    Public strDbConnect As String = GetConnectionString() '"Server=NLBAWPROTDB2;Database=KPI;integrated security=true"
    Public RepList As New DBReps 'List(Of Rep) = New DBReps().List
    Public RangeList As New DBRange ' = GetRanges()
    Public OwnerGroupList As New DBOwnerGroups 'OwnerGroup = GetGroups()
    Public RegionList As New DBRegions ' = GetRegions()
    Public PageList As List(Of Page) = New DBPages().List
    Public Function GetConnectionString() As String
        Dim connstr As String = System.Configuration.ConfigurationManager.AppSettings("connectionstring")
        Return connstr '"Server=NLBAWPROTDB2;Database=KPI;user=report;pwd=report01;MultipleActiveResultSets=True"
        End

    End Function

    Public Sub WriteMessage(ByVal strMessage As String)
        Console.WriteLine(strMessage)
    End Sub


    Public Sub Main(ByVal args As String())
        'expecting <rep> <range> <group> <region>
        Dim SelPage As Page = PageList(PageType.ALLSTATS)
        If args.Count = 0 Then
            WriteMessage(String.Format("Version {1}. Usage: {0} -[switch] [querysetname]", Application.ProductName, Application.ProductVersion))
            WriteMessage(String.Format("[switch] -a (all stats) -s (surveys) -t (transferred)"))
            WriteMessage(String.Format("Valid QuerySets are:"))
            Try
                Dim qsList As List(Of QuerySetDef) = New DBQuerySetDefs().List
                For Each oQs In qsList
                    WriteMessage(String.Format(" {0} : {1} ", oQs.QSetname, oQs.RangeName))
                Next
            Catch ex As Exception
                WriteMessage(String.Format(" {0} : {1} ", ex.Message, ex.InnerException))
            End Try
            'WriteMessage(String.Format("Usage: {0} -[switch] [parameters]", Application.ProductName))
            'WriteMessage(String.Format(" Query ownergroups -o <RangeName> <OwnerGroupName> <Region>"))
            'WriteMessage(String.Format(" Query ownergroups, previous week -oa <OwnerGroupName> <Region>"))
            'WriteMessage(String.Format(" Query users -u <repname> <RangeName> "))
            'WriteMessage(String.Format(" Query users -ua <repname>  "))
            'WriteMessage(String.Format(" Query users -uall  gathers all stats for all users for the past week  "))
            'WriteMessage(String.Format(" Query users -usat  gathers customer sat for all users for the past week  "))
            'WriteMessage(String.Format(" Query users -usur  gathers customer sat for all users for the past week  "))
            'WriteMessage(String.Format(" Query users -ucur  gathers stats for all users for the current week  "))
            'WriteMessage(String.Format(" Query users -osat  gathers customer sat for all ownergroups for the past week  "))
            'WriteMessage(String.Format(" Query users -ostat  <region> gathers stats for all ownergroups or  for the past week  "))
            'WriteMessage(String.Format(" Query users -ostatc <og> <region> gathers stats for all ownergroups or selected og for the current week  "))
        End If
        If args.Count > 0 Then
            For Each arg As String In args

                WriteMessage(arg.ToUpper)

            Next arg
            If args(0) = "-pwd" Then
                ChangePassWord()
                Exit Sub
            End If
            If args(0) = "-a" Then
                SelPage = PageList(PageType.ALLSTATS)
                If args.Count = 2 Then
                    Dim qsList As List(Of QuerySet) = New DBQuerySet(args(1)).List
                    Try
                        For Each oQs In qsList
                            RunQuery(oQs.Owner, oQs.Range, oQs.OwnerGroup, oQs.Region, SelPage)
                        Next
                    Catch ex As Exception
                        WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                    End Try
                End If
            ElseIf args(0) = "-aU" Then
                Dim dbr As New DBReps()
                dbr.UpdateCalcValues()

            ElseIf args(0) = "-s" Then

                SelPage = PageList(PageType.SURVEYS)
                If args.Count = 2 Then
                    Dim qsList As List(Of QuerySet) = New DBQuerySet(args(1)).List
                    Try
                        For Each oQs In qsList
                            RunQuery(oQs.Owner, oQs.Range, oQs.OwnerGroup, oQs.Region, SelPage)
                        Next
                    Catch ex As Exception
                        WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                    End Try
                End If
            ElseIf args(0) = "-t" Then

                SelPage = PageList(PageType.TRANSFERRED)
                If args.Count = 2 Then
                    Dim qsList As List(Of QuerySet) = New DBQuerySet(args(1)).List
                    Try
                        For Each oQs In qsList
                            RunQuery(oQs.Owner, oQs.Range, oQs.OwnerGroup, Nothing, SelPage)
                        Next
                    Catch ex As Exception
                        WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                    End Try
                End If


            ElseIf args(0) = "-o" Then
                If args.Count = 4 Then

                    Try
                        RunQuery(Nothing, RangeList.FindRangeByName(args(1)), OwnerGroupList.FindOwnerGroupByName(args(2)), RegionList.FindRegionByName(args(3)), SelPage)
                    Catch ex As Exception
                        WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                    End Try
                End If
            ElseIf args(0) = "-u" Then
                If args.Count = 3 Then
                    Try
                        RunQuery(RepList.FindRepByName(args(1)), RangeList.FindRangeByName(args(2)), Nothing, Nothing, SelPage)
                    Catch ex As Exception
                        WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                    End Try
                End If

            ElseIf args(0) = "-oa" Then
                If args.Count = 3 Then
                    Try
                        RunQuery(Nothing, RangeList.GetPreviousWeekRange(), OwnerGroupList.FindOwnerGroupByName(args(1)), RegionList.FindRegionByName(args(2)), SelPage)

                    Catch ex As Exception
                        WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                    End Try
                End If
            ElseIf args(0) = "-ua" Then
                If args.Count = 2 Then
                    Try
                        RunQuery(RepList.FindRepByName(args(1)), RangeList.GetPreviousWeekRange(), Nothing, Nothing, SelPage)
                    Catch ex As Exception
                        WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                    End Try
                End If
            ElseIf args(0) = "-usat" Then
                Try
                    Dim pweek As Range = RangeList.GetPreviousWeekRange()
                    For Each aRep In RepList.List
                        WriteMessage(String.Format("Running Customer Sat for {0} for week {1}", aRep.Name, pweek.Name))
                        RunQuery(aRep, pweek, Nothing, Nothing, PageList(1))
                    Next
                    'Runquery(FindRepByName(args(1)), GetPreviousWeekRange(), Nothing, Nothing, PageList(1))
                Catch ex As Exception
                    WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                End Try
            ElseIf args(0) = "-usur" Then
                Try
                    Dim pweek As Range = RangeList.GetPreviousWeekRange()
                    For Each aRep In RepList.List
                        WriteMessage(String.Format("Running Customer Sat for {0} for week {1}", aRep.Name, pweek.Name))
                        RunQuery(aRep, pweek, Nothing, Nothing, PageList(0))
                    Next
                    'Runquery(FindRepByName(args(1)), GetPreviousWeekRange(), Nothing, Nothing, PageList(1))
                Catch ex As Exception
                    WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                End Try
            ElseIf args(0) = "-ucur" Then
                Try
                    Dim cweek As Range = RangeList.GetCurrentWeekRange
                    For Each aRep In RepList.List
                        WriteMessage(String.Format("Running All Stats for {0} for week {1}", aRep.Name, cweek.Name))
                        RunQuery(aRep, cweek, Nothing, Nothing, PageList(0))
                    Next
                    'Runquery(FindRepByName(args(1)), GetPreviousWeekRange(), Nothing, Nothing, PageList(1))
                Catch ex As Exception
                    WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                End Try
            ElseIf args(0) = "-ostat" Then
                'Previous week!!!!
                Try
                    Dim selRegion As Region
                    Dim selGroup As OwnerGroup
                    Dim pweek As Range = RangeList.GetPreviousWeekRange()
                    If args.Count = 3 Then
                        ' no region provided, default to EMEA

                        selRegion = RegionList.FindRegionByName(args(2))
                        selGroup = OwnerGroupList.FindOwnerGroupByName(args(1))
                        WriteMessage(String.Format("Running Stats for {0} for  {1} Region {2}", selGroup.Name, pweek.Name, selRegion.Name))
                        RunQuery(Nothing, pweek, selGroup, selRegion, PageList(0))
                    Else
                        selRegion = RegionList.FindRegionByName("EMEA")
                        For Each aOG In OwnerGroupList.List
                            WriteMessage(String.Format("Running Stats for {0} for  {1} Region {2}", aOG.Name, pweek.Name, selRegion.Name))
                            RunQuery(Nothing, pweek, aOG, selRegion, PageList(0))
                        Next
                    End If



                Catch ex As Exception
                    WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                End Try
            ElseIf args(0) = "-ostatc" Then
                Try
                    Dim selRegion As Region
                    Dim selGroup As OwnerGroup
                    Dim pweek As Range = RangeList.GetCurrentWeekRange()

                    If args.Count = 3 Then
                        ' no region provided, default to EMEA

                        selRegion = RegionList.FindRegionByName(args(2))
                        selGroup = OwnerGroupList.FindOwnerGroupByName(args(1))
                        WriteMessage(String.Format("Running Stats for {0} for  {1} Region {2}", selGroup.Name, pweek.Name, selRegion.Name))
                        RunQuery(Nothing, pweek, selGroup, selRegion, PageList(0))
                    Else
                        selRegion = RegionList.FindRegionByName("EMEA")
                        For Each aOG In OwnerGroupList.List
                            WriteMessage(String.Format("Running Stats for {0} for  {1} Region {2}", aOG.Name, pweek.Name, selRegion.Name))
                            RunQuery(Nothing, pweek, aOG, selRegion, PageList(0))
                        Next
                    End If


                Catch ex As Exception
                    WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                End Try
            ElseIf args(0) = "-osat" Then
                Try
                    Dim selRegion As Region
                    Dim selGroup As OwnerGroup
                    Dim pweek As Range = RangeList.GetPreviousWeekRange()
                    If args.Count = 3 Then
                        ' no region provided, default to EMEA

                        selRegion = RegionList.FindRegionByName(args(2))
                        selGroup = OwnerGroupList.FindOwnerGroupByName(args(1))

                        WriteMessage(String.Format("Running Stats for {0} for  {1} Region {2}", selGroup.Name, pweek.Name, selRegion.Name))
                        RunQuery(Nothing, pweek, selGroup, selRegion, PageList(0))
                        For Each aOG In OwnerGroupList.List
                            WriteMessage(String.Format("Running AllStats for {0} for week {1}", aOG.Name, pweek.Name))
                            RunQuery(Nothing, pweek, aOG, RegionList.FindRegionByName("EMEA"), PageList(0))
                        Next
                    End If
                    'Runquery(FindRepByName(args(1)), GetPreviousWeekRange(), Nothing, Nothing, PageList(1))
                Catch ex As Exception
                    WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
                End Try
            End If

        Else
            Dim mainForm As New Form1
            mainForm.ShowDialog()
        End If

    End Sub


End Module
