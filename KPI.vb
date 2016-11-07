Imports System.Data.SqlClient
Public Class KPI
    Public Name As String
    Public Value As String
    Public Type As String
    Public LookupString As String
    Public OwnerField As KPI
    Public Sub New(ByVal aName As String, ByVal aValue As String, ByVal aLUString As String, ByVal aOwnerField As KPI)
        Name = aName
        Value = aValue
        LookupString = aLUString
        OwnerField = aOwnerField
    End Sub
End Class

Public Class DBKPIs
    Public List As List(Of KPI)

    Public Sub New(ByVal ForPageID As Integer)
        List = GetKPIs(ForPageId)
    End Sub
    Public Function GetKPIs(ByVal PageID As Integer) As List(Of KPI)
        Dim KPIList As New List(Of KPI)
        Dim dbDatabase As New SqlClient.SqlConnection
        dbDatabase.ConnectionString = strDbConnect

        'Dim cmd As New SqlClient.SqlCommand("select me.FieldName, 0.0, me.LookupString, p.FieldName, 0.0, p.LookupString FROM KPIDefinitions me LEFT OUTER JOIN KPIDefinitions p ON p.sequence = me.ownerfieldsequence", conn)
        Dim cmd As New SqlClient.SqlCommand("get_kpidefinitions")
        cmd.Connection = dbDatabase
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@pageid", PageID))
        dbDatabase.Open()
        Dim reader As SqlClient.SqlDataReader = cmd.ExecuteReader
        Dim i As Integer = 0
        While reader.Read()
            Dim Parent
            Parent = reader(3)
            If IsDBNull(Parent) Then
                KPIList.Add(New KPI(reader(0), reader(1), reader(2), Nothing))
            Else
                KPIList.Add(New KPI(reader(0), reader(1), reader(2), New KPI(reader(3), reader(4), reader(5), Nothing)))
            End If

        End While

        ' Call Close when done reading.
        reader.Close()
        dbDatabase.Close()
        Return KPIList

    End Function
End Class