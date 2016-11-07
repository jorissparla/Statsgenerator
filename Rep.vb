Public Class Rep
    Public Name As String
    Public UserId As String
    Public Location As String
    Public Team As String

    Sub New(ByVal aName As String, ByVal auserid As String, ByVal alocation As String, ByVal ateam As String)
        Name = aName
        UserId = auserid
        Location = alocation
        Team = ateam

    End Sub
End Class

Public Class DBReps
    Public Const strSELECT_REPS = "select Name, [UserID], Location, Team from reps"
    Public Const strWHERE = " WHERE manager like '{0}'  order by Team, location, Name"
    Public Const strOrderBy = " ORDER BY Team, location, Name"
    Public List As New List(Of Rep)


    Public Sub New()
        List = GetReps()
    End Sub

    Public Sub New(ByVal manager As String)
        List = GetReps(manager)
    End Sub
    Public Function GetReps() As List(Of Rep)
        Dim aRepList As New List(Of Rep)
        Dim conn As New SqlClient.SqlConnection(strDbConnect)
        Dim cmd As New SqlClient.SqlCommand(strSELECT_REPS, conn)
        conn.Open()
        Dim reader As SqlClient.SqlDataReader = cmd.ExecuteReader
        Dim i As Integer = 0
        While reader.Read()

            aRepList.Add(New Rep(reader(0), reader(1), reader(2), reader(3)))

        End While
        ' Call Close when done reading.
        reader.Close()
        conn.Close()

        Return aRepList

    End Function

    Public Function GetReps(ByVal manager As String) As List(Of Rep)
        Dim RepList As New List(Of Rep)

        If manager = "(All)" Then manager = "%"
        Dim conn As New SqlClient.SqlConnection(strDbConnect)
        Dim cmd As New SqlClient.SqlCommand(strSELECT_REPS + String.Format(strWHERE, manager), conn)
        conn.Open()
        Dim reader As SqlClient.SqlDataReader = cmd.ExecuteReader
        While reader.Read()
            RepList.Add(New Rep(reader(0), reader(1), reader(2), reader(3)))
        End While
        ' Call Close when done reading.
        reader.Close()
        conn.Close()
        Return RepList

    End Function

    Public Sub UpdateCalcValues()
        Dim conn As New SqlClient.SqlConnection(strDbConnect)
        Dim cmd As New SqlClient.SqlCommand("CalcAllCalculatedValuesCurrentWeek", conn)
        cmd.CommandType = CommandType.StoredProcedure
        conn.Open()
        cmd.ExecuteNonQuery()
        conn.Close()
        'CalcAllCalculatedValuesCurrentWeek
        'ExecuteNonQuery()
    End Sub

    Public Function FindRepByName(ByVal strName As String) As Rep

        For Each aRep As Rep In List
            If aRep.Name = strName Then
                Return aRep
            End If
        Next
        Return New Rep("", "", "", "")
    End Function
End Class