VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Playing with PC Speaker Using Assembly"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   503
      TabIndex        =   7
      Top             =   2760
      Width           =   5055
      Begin VB.CommandButton Command2 
         Caption         =   "Play a tone"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   503
      TabIndex        =   4
      Top             =   360
      Width           =   5055
      Begin VB.CommandButton Command1 
         Caption         =   "Test"
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2880
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Enter the frequency in hertz: "
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   450
         Width           =   2040
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enter the duration in milliseconds:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   930
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PCSpeaker As New clsASMPCSpeaker

Private Sub Command1_Click()
If Val(Text1) > 20000 Then
    MsgBox "Frequency should be less. You are a human (I suppose) and so you cannot hear such high frequency.", vbCritical
    Text1.Text = ""
    Exit Sub
End If
If Val(Text2) > 30000 Then
    Dim Response
    Response = MsgBox("The duration is " & CStr(Val(Text2) / 1000) & " minutes. During this time this running application will seem to have hanged because of using the Sleep API function. But after that duration, it will be ok. Do you want to continue?", vbInformation + vbYesNo + vbDefaultButton2)
    If Response = vbNo Then
        Text2.Text = ""
        Exit Sub
    End If
End If
PCSpeaker.Beep Val(Text1), Val(Text2)
End Sub

Private Sub Command2_Click()
Dim i As Long
For i = 100 To 8000 Step 200
    PCSpeaker.Beep i, 40
Next i
For i = 8000 To 100 Step -200
    PCSpeaker.Beep i, 40
Next i
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub
