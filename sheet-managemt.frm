VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HideUnhideForm 
   Caption         =   "Sheets"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7680
   OleObjectBlob   =   "sheet-managemt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HideUnhideForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cb1_Click()
        If cb1.Value = True Then
            For i = 0 To lbI.ListCount - 1
                lbI.Selected(i) = True
            Next i
        End If
        If cb1.Value = False Then
            For i = 0 To lbI.ListCount - 1
                lbI.Selected(i) = False
            Next i
        End If
End Sub
Private Sub cb2_Click()
        If cb2.Value = True Then
            For i = 0 To lbV.ListCount - 1
                lbV.Selected(i) = True
            Next i
        End If
        If cb2.Value = False Then
            For i = 0 To lbV.ListCount - 1
                lbV.Selected(i) = False
            Next i
        End If
End Sub
Private Sub cbIAZ_Click()
        Dim j As Long
        Dim i As Long
        Dim temp As Variant
        'Use Bubble sort method to put listBox in A-Z order
        With Me.lbI
            For j = 0 To .ListCount - 2
                For i = 0 To .ListCount - 2
                    If LCase(.List(i)) > LCase(.List(i + 1)) Then
                        temp = .List(i)
                        .List(i) = .List(i + 1)
                        .List(i + 1) = temp
                    End If
                Next i
            Next j
        End With
End Sub
Private Sub cbIZA_Click()
        Dim j As Long
        Dim i As Long
        Dim temp As Variant
        'Use Bubble sort method to put listBox in Z-A order
        With Me.lbI
            For j = 0 To .ListCount - 2
                For i = 0 To .ListCount - 2
                    If LCase(.List(i)) < LCase(.List(i + 1)) Then
                        temp = .List(i)
                        .List(i) = .List(i + 1)
                        .List(i + 1) = temp
                    End If
                Next i
            Next j
        End With
End Sub
Private Sub cbVAZ_Click()
        Dim j As Long
        Dim i As Long
        Dim temp As Variant
        'Use Bubble sort method to put listBox in A-Z order
        With Me.lbV
            For j = 0 To .ListCount - 2
                For i = 0 To .ListCount - 2
                    If LCase(.List(i)) > LCase(.List(i + 1)) Then
                        temp = .List(i)
                        .List(i) = .List(i + 1)
                        .List(i + 1) = temp
                    End If
                Next i
            Next j
        End With
End Sub
Private Sub cbVZA_Click()
        Dim j As Long
        Dim i As Long
        Dim temp As Variant
        'Use Bubble sort method to put listBox in Z-A order
        With Me.lbV
            For j = 0 To .ListCount - 2
                For i = 0 To .ListCount - 2
                    If LCase(.List(i)) < LCase(.List(i + 1)) Then
                        temp = .List(i)
                        .List(i) = .List(i + 1)
                        .List(i + 1) = temp
                    End If
                Next i
            Next j
        End With
End Sub

Private Sub Close2_Click()
     Unload Me
End Sub
Private Sub hideAll_Click()
        Call Clear
        Call ActivateForm
        Call HideAllSheet
        Call ActivateForm
End Sub

Private Sub lbI_Click()

End Sub

Private Sub lbV_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        aa = lbV.ListIndex
        ThisWorkbook.Sheets(lbV.List(aa)).Activate
End Sub

Private Sub unhideAll_Click()
        Call Clear
        Call ActivateForm
        Call UnhideAllSheet
        Call ActivateForm
End Sub
Private Sub OK_Click()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
        'Duyet lan luot tung list(i) trong moi listbox, Unhide / hide list(i)
        If Me.lbV.ListCount = 0 Then
            MsgBox "You must choose at least One sheet to be visible !!"
            Exit Sub
        Else
        'Visible sheet
            For ij = 0 To lbV.ListCount - 1
                m = lbV.List(ij)
                ThisWorkbook.Sheets(m).Visible = True
            Next

        'Invisible sheet
            For ii = 0 To lbI.ListCount - 1
                n = lbI.List(ii)
                ThisWorkbook.Sheets(n).Visible = False
            Next
        End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Private Sub Left_Click()
    Me.lbI2.Clear
    u = Me.lbV.ListCount - 1
     On Error Resume Next
     'Add sang ben kia
            For i = 0 To u Step 1
                If Me.lbV.Selected(i) = True Then
                    Me.lbI.AddItem lbV.List(i)
                    Me.lbI2.AddItem lbV.List(i)
                    'Me.lbV.RemoveItem (i)
                End If
            Next
    'Delete o ben kia
            For j = 0 To Me.lbI2.ListCount - 1
                    For i = 0 To Me.lbV.ListCount - 1 Step 1
                        If Me.lbI2.List(j) = Me.lbV.List(i) Then
                            Me.lbV.RemoveItem (i)
                        End If
                    Next
            Next
    On Error GoTo 0
End Sub
Private Sub Right_Click()
    Me.lbV2.Clear
      h = Me.lbI.ListCount - 1
        On Error Resume Next
        'Add sang ben kia
            For i = 0 To h Step 1
                  If Me.lbI.Selected(i) = True Then
                        Me.lbV.AddItem Me.lbI.List(i)
                        Me.lbV2.AddItem Me.lbI.List(i)
                        'Me.lbI.RemoveItem (i)
                  End If
            Next i
        'Delete o ben kia
            For j = 0 To Me.lbV2.ListCount - 1
                    For i = 0 To Me.lbI.ListCount - 1 Step 1
                        If Me.lbV2.List(j) = Me.lbI.List(i) Then
                            Me.lbI.RemoveItem (i)
                        End If
                    Next
            Next
        On Error GoTo 0
End Sub
Private Sub UserForm_Activate()
        Call ActivateForm
End Sub
Sub UnhideAllSheet()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
      u_count = ThisWorkbook.Worksheets.count
      For iu = 1 To u_count
           ThisWorkbook.Sheets(iu).Visible = True
     Next iu
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub HideAllSheet()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
    If Me.lbV.ListCount = 1 Then
        Exit Sub
    Else
         h_count = ThisWorkbook.Worksheets.count
         For ih = 1 To h_count
               ThisWorkbook.Sheets(1).Visible = True
              ThisWorkbook.Sheets(ih).Visible = False
        Next ih
    End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub ActivateForm()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
    Call Clear
      a_count = ThisWorkbook.Worksheets.count
      tbNum.Value = a_count
      For ia = 1 To a_count
           If ThisWorkbook.Sheets(ia).Visible = True Then
                lbV.AddItem ThisWorkbook.Sheets(ia).Name
           Else
                 lbI.AddItem ThisWorkbook.Sheets(ia).Name
           End If
     Next ia
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub Clear()
        Me.lbI.Clear
        Me.lbV.Clear
End Sub
