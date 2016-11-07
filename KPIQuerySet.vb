Imports System.Data.SqlClient

Public Class QuerySet
    Public Owner As Rep
    Public RangeName As String
    Public Range As Range
    Public OwnerGroup As OwnerGroup
    Public Region As Region

    Public UserId As String
    'Public OwnerGroup As String
    'Public Region As String
    Public Name As String = "Query"
    Sub New(ByVal myOwner As String, ByVal myRange As String, ByVal myUserId As String, ByVal myOwnerGroup As String, ByVal MyRegion As String, Optional ByVal MyName As String = "QUERY")
        Owner = RepList.FindRepByName(myOwner)
        RangeName = myRange
        Range = RangeList.FindRangeByName(myRange)
        UserId = myUserId
        OwnerGroup = OwnerGroupList.FindOwnerGroupByName(myOwnerGroup)
        Region = RegionList.FindRegionByName(MyRegion)
        Name = MyName
    End Sub
    Overrides Function ToString() As String
        Return String.Format("{0} :: {1} ", Range.Name, OwnerGroup.Name)
    End Function
    Function OwnerGroupName() As String
        Return OwnerGroup.Name
    End Function
End Class


Public Class KPIQuerySet
    Public aKPI As KPI
    Public aQuerySet As QuerySet
    Sub New(ByVal myKPI As KPI, ByVal myQuerySet As QuerySet) ' As String, ByVal myRange As String, ByVal myUserId As String, ByVal myOwnerGroup As String, ByVal MyRegion As String)
        aKPI = myKPI
        aQuerySet = myQuerySet

    End Sub
End Class

Public Class QuerySetDef
    Public QSetname As String
    Public RangeName As String
    Public Sub New(ByVal MyName As String, ByVal MyRange As String)
        QSetname = MyName
        RangeName = MyRange

    End Sub
    Public ReadOnly Property Name() As String
        Get
            Return QSetname
        End Get
    End Property
    Public ReadOnly Property DisplayName() As String
        Get
            Return String.Format("{0} ({1}) ", QSetname, RangeName)
        End Get
    End Property
End Class

Public Class DBQuerySetDefs
    Public List As List(Of QuerySetDef)

    Public Sub New()
        List = ListQuerySets()
    End Sub
    Public Function ListQuerySets() As List(Of QuerySetDef)
        Dim QSList As New List(Of QuerySetDef)
        Dim conn As New SqlClient.SqlConnection(strDbConnect)
        Dim cmd As New SqlClient.SqlCommand("[listQuerySets]", conn)
        cmd.CommandType = CommandType.StoredProcedure
        conn.Open()
        Dim reader As SqlClient.SqlDataReader = cmd.ExecuteReader
        Dim i As Integer = 0
        While reader.Read()

            QSList.Add(New QuerySetDef(reader(0), reader(1)))

        End While

        reader.Close()
        conn.Close()
        Return QSList

    End Function
    Public Function ListQuerySetGroups() As String

    End Function
End Class

Public Class DBQuerySet
    Public List As New List(Of QuerySet)
    Sub New(ByVal StrQueryName As String)
        List = GetQuerySets(StrQueryName)
    End Sub

    Sub New()
        ' TODO: Complete member initialization 
    End Sub

    Public Function GetQuerySets(ByVal aName As String) As List(Of QuerySet)
        Dim QSList As New List(Of QuerySet)
        Dim conn As New SqlClient.SqlConnection(strDbConnect)
        Dim cmd As New SqlClient.SqlCommand("[getQuerySet]", conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@name", aName))
        conn.Open()
        Dim reader As SqlClient.SqlDataReader = cmd.ExecuteReader
        Dim i As Integer = 0
        While reader.Read()

            QSList.Add(New QuerySet(reader(0), reader(1), reader(2), reader(3), reader(4), aName))

        End While

        reader.Close()
        conn.Close()
        Return QSList

    End Function

    Public Overrides Function ToString() As String
        Dim sRet As String = ""
        For Each oQ In List
            sRet = oQ.ToString + "::"
        Next
        Return sRet
    End Function

End Class


Public Class KPIQuerySetResultDB

    Dim dbDatabase As New SqlClient.SqlConnection
    Dim InsertCmd As New SqlClient.SqlCommand
    Dim UpdateCmd As New SqlClient.SqlCommand
    Dim SelectCmd As New SqlClient.SqlCommand
    Dim DeleteCmd As New SqlCommand
    Dim SelectReader As SqlDataReader
    Public aKPIQuerySet As KPIQuerySet

    Sub New(ByVal myKPIQuerySet As KPIQuerySet)
        aKPIQuerySet = myKPIQuerySet
        dbDatabase.ConnectionString = strDbConnect
        InsertCmd.Connection = dbDatabase
        InsertCmd.CommandType = CommandType.StoredProcedure
        InsertCmd.CommandText = "insert_kpivalue"
        UpdateCmd.Connection = dbDatabase
        UpdateCmd.CommandType = CommandType.StoredProcedure
        UpdateCmd.CommandText = "Update_kpivalue"
        SelectCmd.Connection = dbDatabase
        SelectCmd.CommandType = CommandType.Text
        SelectCmd.CommandText = "SELECT dbo.exist_kpivalue ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}')"
        DeleteCmd.CommandText = "delete_kpivalues"
        DeleteCmd.CommandType = CommandType.StoredProcedure
        DeleteCmd.Connection = dbDatabase
        dbDatabase.Open()


    End Sub

    Public Sub DeleteALLFromDB()
        With DeleteCmd
            .Parameters.Add(New SqlParameter("@OwnerList", aKPIQuerySet.aQuerySet.Owner.Name))
            .Parameters.Add(New SqlParameter("@RangeName", aKPIQuerySet.aQuerySet.Range.Name))
            .Parameters.Add(New SqlParameter("@UserID", aKPIQuerySet.aQuerySet.UserId))
            .Parameters.Add(New SqlParameter("@OwnerGroupList", aKPIQuerySet.aQuerySet.OwnerGroup.Name))
            .Parameters.Add(New SqlParameter("@Region", aKPIQuerySet.aQuerySet.Region.Name))
        End With
        DeleteCmd.ExecuteNonQuery()
    End Sub

    Public Sub SaveToDB()
        If Exists() > 0 Then
            UpdateToDb()
        Else
            InsertToDb()
        End If
        dbDatabase.Close()
    End Sub

    Public Function Exists() As Integer
        Dim intret As Integer
        SelectCmd.CommandText = String.Format(SelectCmd.CommandText, aKPIQuerySet.aKPI.Name, "", aKPIQuerySet.aQuerySet.RangeName, aKPIQuerySet.aQuerySet.Owner.UserId, aKPIQuerySet.aQuerySet.OwnerGroup.Name, aKPIQuerySet.aQuerySet.Region.Name)
        '  Remove for single quote names SelectCmd.CommandText = String.Format(SelectCmd.CommandText, aKPIQuerySet.aKPI.Name, aKPIQuerySet.aQuerySet.Owner.Name, aKPIQuerySet.aQuerySet.RangeName, aKPIQuerySet.aQuerySet.Owner.UserId, aKPIQuerySet.aQuerySet.OwnerGroup.Name, aKPIQuerySet.aQuerySet.Region.Name)
        SelectReader = SelectCmd.ExecuteReader
        SelectReader.Read()
        intret = SelectReader.Item(0)
        SelectReader.Close()
        Return intret
    End Function

    Public Sub InsertToDb()
        With InsertCmd
            .Parameters.Add(New SqlParameter("@KPIName", aKPIQuerySet.aKPI.Name))
            .Parameters.Add(New SqlParameter("@value", aKPIQuerySet.aKPI.Value))
            .Parameters.Add(New SqlParameter("@OwnerList", aKPIQuerySet.aQuerySet.Owner.Name))
            .Parameters.Add(New SqlParameter("@RangeName", aKPIQuerySet.aQuerySet.Range.Name))
            .Parameters.Add(New SqlParameter("@RunDate", Now()))
            .Parameters.Add(New SqlParameter("@UserID", aKPIQuerySet.aQuerySet.UserId))
            .Parameters.Add(New SqlParameter("@OwnerGroupList", aKPIQuerySet.aQuerySet.OwnerGroup.Name))
            .Parameters.Add(New SqlParameter("@Region", aKPIQuerySet.aQuerySet.Region.Name))
        End With
        InsertCmd.ExecuteNonQuery()

    End Sub



    Public Sub UpdateToDb()
        With UpdateCmd
            .Parameters.Add(New SqlParameter("@KPIName", aKPIQuerySet.aKPI.Name))
            .Parameters.Add(New SqlParameter("@value", aKPIQuerySet.aKPI.Value))
            .Parameters.Add(New SqlParameter("@OwnerList", aKPIQuerySet.aQuerySet.Owner.Name))
            .Parameters.Add(New SqlParameter("@RangeName", aKPIQuerySet.aQuerySet.Range.Name))
            .Parameters.Add(New SqlParameter("@RunDate", Now()))
            .Parameters.Add(New SqlParameter("@UserID", aKPIQuerySet.aQuerySet.UserId))
            .Parameters.Add(New SqlParameter("@OwnerGroupList", aKPIQuerySet.aQuerySet.OwnerGroup.Name))
            .Parameters.Add(New SqlParameter("@Region", aKPIQuerySet.aQuerySet.Region.Name))
        End With
        UpdateCmd.ExecuteNonQuery()
    End Sub
End Class