Public Class Range
    Public Name As String
    Public FromDate As Date
    Public toDate As Date
    Public Type As String
    Public Sub New(ByVal sName As String, ByVal dtFromDate As Date, ByVal dtToDate As Date, ByVal sType As String)
        Name = sName
        FromDate = dtFromDate
        toDate = dtToDate
        Type = sType
    End Sub
End Class

Public Class DBRange
    Public List As List(Of Range)

    Public Sub New()
        List = GetRanges()
    End Sub

    Public Function GetRanges() As List(Of Range)
        Dim aRangeList As New List(Of Range)
        Dim conn As New SqlClient.SqlConnection(strDbConnect)
        '      Dim cmd As New SqlClient.SqlCommand("select Name, fromdate, todate, RangeType from ranges order by Name", conn)
        Dim cmd As New SqlClient.SqlCommand("cboListRanges", conn)
        conn.Open()
        Dim reader As SqlClient.SqlDataReader = cmd.ExecuteReader
        Dim i As Integer = 0
        While reader.Read()

            aRangeList.Add(New Range(reader(0), reader(1), reader(2), reader(3)))

        End While
        reader.Close()
        conn.Close()
        Return aRangeList

    End Function

    Public Function FindRangeByName(ByVal strRange As String) As Range
        For Each aRange As Range In List
            If aRange.Name.ToUpper = strRange.ToUpper Then
                Return aRange
            End If
        Next
        Return Nothing
    End Function

    Public Function GetPreviousWeekRange() As Range
        Dim fRange As Range = Nothing
        Dim conn As New SqlClient.SqlConnection(strDbConnect)
        Dim cmd As New SqlClient.SqlCommand("select Name, fromdate, todate, RangeType from ranges WHERE DateAdd(d, -6, getdate()) between fromdate and todate and rangetype='Week'", conn)
        conn.Open()
        Dim reader As SqlClient.SqlDataReader = cmd.ExecuteReader
        Dim i As Integer = 0
        While reader.Read()
            fRange = New Range(reader(0), reader(1), reader(2), reader(3))
        End While
        ' Call Close when done reading.

        reader.Close()
        conn.Close()
        Return fRange
    End Function

    Public Function GetCurrentWeekRange() As Range
        Dim fRange As Range = Nothing
        Dim conn As New SqlClient.SqlConnection(strDbConnect)
        Dim cmd As New SqlClient.SqlCommand("select Name, fromdate, todate, RangeType from ranges WHERE getdate() between fromdate and todate and rangetype='Week'", conn)
        conn.Open()
        Dim reader As SqlClient.SqlDataReader = cmd.ExecuteReader
        Dim i As Integer = 0
        While reader.Read()
            fRange = New Range(reader(0), reader(1), reader(2), reader(3))
        End While
        ' Call Close when done reading.
        reader.Close()
        conn.Close()
        Return fRange
    End Function

End Class