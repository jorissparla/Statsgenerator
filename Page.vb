Public Class Page
    Public PageID As Integer
    Public PageName As String
    Public PageURL As String
    Public StartString As String
    Sub New(ByVal myPageId As Integer, ByVal MyPageName As String, ByVal MyPageURL As String, ByVal MyStartString As String)
        PageID = myPageId
        PageName = MyPageName
        PageURL = MyPageURL
        StartString = MyStartString
    End Sub
End Class

Public Class DBPages
    Public List As New List(Of Page)
    Private SQLString As String = "SELECT [ID],[PageName],[PageQueryString], [StartString]  FROM [KPI].[dbo].[QueryPageDef]"

    Public Sub New()
        Try
            List = GetPages()
        Catch ex As Exception
            WriteMessage(ex.Message)
        End Try

    End Sub
    Private Function GetPages() As List(Of Page)
        Dim PageList As New List(Of Page)
        Dim conn As New SqlClient.SqlConnection(strDbConnect)
        Dim cmd As New SqlClient.SqlCommand(SQLString, conn)
        conn.Open()
        Dim reader As SqlClient.SqlDataReader = cmd.ExecuteReader
        Dim i As Integer = 0
        While reader.Read()

            PageList.Add(New Page(reader(0), reader(1), reader(2), reader(3)))
        
        End While
        ' Call Close when done reading.

        reader.Close()
        conn.Close()
        Return PageList
    End Function
End Class