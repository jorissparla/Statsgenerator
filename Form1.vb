Public Class Form1
   

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        UpdateStatus("")
        '   testRegEx()
        RefreshRepList()
        LoadPagesList()
        Dim lstQSetDefs As List(Of QuerySetDef) = New DBQuerySetDefs().ListQuerySets
        Dim lstQSet As List(Of QuerySet)
        LbQuerySets.DataSource = lstQSetDefs
        '  LbQuerySets.DataSource = New DBQuerySet
        LbQuerySets.DisplayMember = "DisplayName"
        LbQuerySets.ValueMember = "Name"
        'LbQuerySets.Set()

        Dim x As Integer
        x = 1

        For Each aRange As Range In RangeList.List
            If Not IsNothing(aRange) Then
                LbRanges.Items.Add(aRange.Name)
            End If
        Next

        For Each aGroup As OwnerGroup In OwnerGroupList.List
            If Not IsNothing(aGroup) Then
                lbOwnerGroups.Items.Add(aGroup.Name)
            End If

        Next

        For Each aRegion As Region In RegionList.List
            If Not IsNothing(aRegion) Then
                lbRegions.Items.Add(aRegion.Name)
            End If
        Next
    End Sub


    Sub RefreshRepList()
        lbReps.Items.Clear()
        For Each aRep As Rep In RepList.List
            If Not IsNothing(aRep) Then
                lbReps.Items.Add(aRep.Name)
            End If
        Next
    End Sub

    Sub LoadPagesList()
        CboPages.Items.Clear()
        For Each aPage As Page In PageList
            CboPages.Items.Add(aPage.PageName)
            CboPages.SelectedIndex = 0
        Next
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRun.Click
        Dim iItem As Integer
        Dim SelOwnerGroup As OwnerGroup = Nothing
        Dim SelRegion As Region = Nothing
        Dim selRep As Rep = Nothing
        Dim selRange As Range = Nothing
        Dim sO As String
        Dim sR As String
        Dim selPage As Page = PageList(CboPages.SelectedIndex)
        Dim intStatus As Integer = 0
        Dim qsSelected As Integer = LbQuerySets.SelectedIndex

        If qsSelected > -1 Then
            Dim qsList As List(Of QuerySet) = New DBQuerySet(LbQuerySets.Items(qsSelected).Name).List
            Try
                For Each oQs In qsList
                    intStatus = intStatus + RunQuery(oQs.Owner, oQs.Range, oQs.OwnerGroup, oQs.Region, selPage)
                Next
            Catch ex As Exception
                WriteMessage(String.Format("{0} : {1}", Now(), ex.Message))
            End Try
            'intStatus = RunqueryEx(LbQuerySets.Items(qsSelected), selPage)

        Else
            Dim jRange As Integer = LbRanges.SelectedIndex
            If jRange = -1 Then
                UpdateStatusError("A Range Needs to be selected")
            Else
                selRange = RangeList.List(jRange)

                If lbOwnerGroups.SelectedIndex > -1 Then
                    SelOwnerGroup = OwnerGroupList.List(lbOwnerGroups.SelectedIndex)
                End If

                If lbRegions.SelectedIndex > -1 Then
                    SelRegion = RegionList.List(lbRegions.SelectedIndex)

                End If
                For iItem = 0 To lbReps.SelectedItems.Count - 1
                    'For iItem = 0 To CheckedListBox1.Items.Count - 1


                    Dim idx As Integer = lbReps.SelectedIndices(iItem)
                    selRep = RepList.List(idx)

                    UpdateStatus(String.Format(" Processing   {0}", selRep.Name))
                    intStatus = RunQuery(selRep, selRange, SelOwnerGroup, SelRegion, selPage)


                Next
                If lbReps.SelectedItems.Count = 0 Then
                    If IsNothing(SelOwnerGroup) Then
                        sO = ""
                    Else
                        sO = SelOwnerGroup.Name
                    End If
                    If IsNothing(SelRegion) Then
                        sR = ""
                    Else
                        sR = SelRegion.Name

                    End If
                    UpdateStatus(String.Format(" Processing   {0} : {1}", sO, sR))
                    intStatus = RunQuery(selRep, selRange, SelOwnerGroup, SelRegion, selPage)
                End If

            End If
        End If
        If intStatus = 0 Then
            UpdateStatus("Done..")
        Else
            UpdateStatus("Finished with errors")
        End If


    End Sub

    Sub UpdateStatus(ByVal strText As String)
        LblStatus.ForeColor = Color.Black

        LblStatus.Text = strText
        Me.Refresh()
    End Sub

    Sub UpdateStatusError(ByVal StrText As String)
        LblStatus.ForeColor = Color.Red
        LblStatus.Text = StrText
    End Sub

    Private Sub ListBox2_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbReps.DoubleClick
        For i As Integer = 1 To lbReps.Items.Count - 1
            lbReps.SelectedItem = lbReps.Items(i)
        Next
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ClearAllSelections()
    End Sub

    Sub ClearAllSelections()
        lbReps.ClearSelected()
        lbOwnerGroups.ClearSelected()
        LbRanges.ClearSelected()
        lbRegions.ClearSelected()
        LbQuerySets.ClearSelected()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboManagers.SelectedIndexChanged
        RepList = New DBReps(CboManagers.Text)
        RefreshRepList()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim i As Integer
        For i = 1 To lbReps.Items.Count
            lbReps.SelectedIndices.Add(i - 1)

        Next
    End Sub
End Class



