
Private Sub btnEnd_Click()
    Unload Me
End Sub

Private Sub btnSend_Click()
    bearbeiter = cbxBearb.Value
    stichwort = txbStichw.Value
    issueType = cbxIssue.Value
    projektId = txbProjId.Value
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    With cbxBearb
        .AddItem "username1"
        .AddItem "username2"
        .AddItem "username3"
    End With
    With cbxIssue
        .AddItem "Task"
        .AddItem "Bug"
        .AddItem "Story"
        .AddItem "Test"
        .AddItem "Improvement"
    End With
    Me.cbxIssue.Text = Me.cbxIssue.List(0)
    With Me.txbProjId
        .Value = "30611" 'Project ID
    End With
End Sub
