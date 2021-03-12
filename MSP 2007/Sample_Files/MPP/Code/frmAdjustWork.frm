VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdjustWork 
   Caption         =   "Adjust Work"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "frmAdjustWork.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdjustWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    'Close the UserForm
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim t As Task
Dim a As Assignment
Dim Adjuster As Single
    
    'Make sure the user has typed a number
    If Not IsNumeric(txtPercent.Text) Then
        MsgBox "You must type a valid number in this field.", vbInformation
        Exit Sub
    End If
    
    'Divide the number the user entered by 100, since they entered a percentage
    Adjuster = Val(txtPercent.Text) / 100

    'Loop through each task in the project
    For Each t In ActiveProject.Tasks
        'Skip over blank rows, if any
        If Not t Is Nothing Then
            'Loop through each resource in your project
            For Each a In t.Assignments
                'If the resource name is equal to the resource the user selected,
                'then adjust the work by the Adjuster value
                If a.ResourceName = cboResource.Text Then
                    a.RemainingWork = a.RemainingWork + (a.RemainingWork * Adjuster)
                End If
            Next a
        End If
    Next t
                
    'Display message confirming success
    MsgBox "The work has been adjusted successfully!", vbInformation
    
    'Close the form
    Unload Me

End Sub

Private Sub UserForm_Initialize()
Dim r As Resource
    
    For Each r In ActiveProject.Resources
        cboResource.AddItem r.Name
    Next r
End Sub
