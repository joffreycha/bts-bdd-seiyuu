VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "GESTIO~1.frx":0000
      Left            =   720
      List            =   "GESTIO~1.frx":000D
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    For i = 0 To List1.ListCount
        If Text1.Text <> "" And Text1.Text = List1.List(i) Then
            MsgBox List1.List(i) & " trouvé !", vbInformation + vbOKOnly
            Exit Sub
        End If
    Next i
    MsgBox "aucun element trouvé", vbOKOnly + vbExclamation
End If
End Sub
