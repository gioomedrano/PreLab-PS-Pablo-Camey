VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3840
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   4320
      Top             =   2400
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   1920
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   3000
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRESIONAME"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Shape bote 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   1
      Shape           =   3  'Circle
      Top             =   1
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Timer1.Enabled = True
Timer3.Enabled = True
End Sub

Private Sub Timer1_Timer()

 If bote.Top < 3000 Then
    bote.Top = bote.Top + 100
Else
    bote.Left = bote.Left + 100
End If

If bote.Left > 1000 Then
    Timer2.Enabled = True
    Timer1.Enabled = False
End If
If bote.Left > 3000 Then
  Timer3.Enabled = True
End If
End Sub
Private Sub Timer2_Timer()

If bote.Left > 1000 Then
    bote.Top = bote.Top - 100
End If
If bote.Top < 5 Then
 bote.Left = bote.Left + 100
 bote.Top = bote.Top + 100
End If

If bote.Left > 2000 Then
     bote.Top = bote.Top + 200
End If

If bote.Top > 3000 Then
   bote.Left = bote.Left + 100
   bote.Top = bote.Top - 100
End If

If bote.Left > 3000 Then
    Timer2.Enabled = False
End If

End Sub

Private Sub Timer3_Timer()

If bote.Left > 3000 Then
    bote.Top = bote.Top - 100
End If

If bote.Top < 5 Then
 bote.Left = bote.Left + 100
 bote.Top = bote.Top + 100
End If

If bote.Left > 4000 Then
 bote.Top = bote.Top + 200
End If
If bote.Top > 3000 Then
    bote.Left = bote.Left + 100
    bote.Top = bote.Top - 100
End If
If bote.Left > 5000 Then
 bote.Top = bote.Top - 200
End If

End Sub

