Public Class Region
    Public Name As String
    Public IDSTring As String

    Sub New(ByVal _name As String, ByVal _IDSTring As String)
        Name = _name
        IDSTring = _IDSTring
    End Sub


End Class

Public Class DBRegions
    Public List As List(Of Region)

    Public Sub New()
        List = GetRegions()
    End Sub

    Public Function GetRegions() As List(Of Region)
        Dim aRegionList As New List(Of Region)
        Dim conn As New SqlClient.SqlConnection(strDbConnect)
        Dim cmd As New SqlClient.SqlCommand("select RegionName, RegionID from regions order by RegionID", conn)
        conn.Open()
        Dim reader As SqlClient.SqlDataReader = cmd.ExecuteReader

        While reader.Read()

            aRegionList.Add(New Region(reader(0), reader(1)))
   
        End While
        ' Call Close when done reading.
        reader.Close()
        conn.Close()
        Return aRegionList

    End Function

    Function FindRegionByName(ByVal strRegion As String) As Region
        For Each aRegion In List
            If aRegion.Name.ToUpper = strRegion.ToUpper Then
                Return aRegion
                Exit Function
            End If
        Next
        Return New Region("", "") 'FindRegionByName("EMEA")

    End Function

End Class

