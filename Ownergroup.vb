
Public Class OwnerGroup

    Public Name As String
    Public IDString As String
    Public Sub New(ByVal _name As String, ByVal _IDString As String)
        Name = _name
        IDString = _IDString

    End Sub
End Class

Public Class DBOwnerGroups
    Public List As New List(Of OwnerGroup)

    Public Sub New()
        List = GetGroups()
    End Sub
    Public Function GetGroups() As List(Of OwnerGroup)
        Dim OgList As New List(Of OwnerGroup)
        Dim conn As New SqlClient.SqlConnection(strDbConnect)
        Dim cmd As New SqlClient.SqlCommand("select GroupName, IDString from OwnerGroups", conn)
        conn.Open()
        Dim reader As SqlClient.SqlDataReader = cmd.ExecuteReader
        Dim i As Integer = 0
        While reader.Read()

            OgList.Add(New OwnerGroup(reader(0), reader(1)))

        End While

        reader.Close()
        conn.Close()
        Return OgList

    End Function

    Public Function FindOwnerGroupByName(ByVal strOG As String) As OwnerGroup
        For Each aOG In List
            If aOG.Name.ToUpper = strOG.ToUpper Then
                Return aOG
            End If
        Next
        Return New OwnerGroup("", "")

    End Function

End Class