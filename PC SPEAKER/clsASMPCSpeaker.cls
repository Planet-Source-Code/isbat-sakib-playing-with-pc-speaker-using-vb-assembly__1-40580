VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsASMPCSpeaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------
'Class for using the PC Speaker using VB and Assembly

'Use it as you like and at your own risks. I am not responsible for
'anything. After all, this is assembly! But give me a little credit
'if you use this class in your apps. I am encouraged by the great
'submissions to planet-source-code showing the techniques to use
'assembly in VB. I decided to make something useful using assembly.

'There is a API function in WinNT based systems - BeepAPI function.
'You can create sounds of any frequency and for duration by using
'that function. But the BeepAPI function in Win98 is very limited.
'It can't do it. So, I made this class that can do the same work
'in Win98 also. I haven't tested this code in other machines. But
'it should work. Please, report to me if it doesn't work. My email
'address is sakib039@hotmail.com

'Don't forget to check all the other Assembly+VB submissions in
'planet-source-code. And comments and votes on this thing are welcome.
'------------------------------------------------------------------------------

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal ptrMC As Long, ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long, ByVal P4 As Long) As Long

'Code for turning the PC Speaker on at the specified frequency
Const HexCode1 = "5589E553565766B8B60066E743B8DC341200BA000000008B5D08F7F389C36631C0660FB6C366E7426631C0660FB6C766E74266E5616631C96689C1660D030066E76131C00FB7C15F5E5B89EC5DC21000"
'Code for turning the PC Speaker off
Const HexCode2 = "5589E55356578B450866E7615F5E5B89EC5DC21000"

Private MachineCode1() As Byte  'This is the byte array that holds the machine code to turn on the speaker
Private MachineCode2() As Byte  'This is the byte array that holds the machine code to turn off the speaker
Private pointerMC1 As Long      'Pointer to MachineCode1
Private pointerMC2 As Long      'Pointer to MachineCode2

Private Sub Class_Initialize()

'In this initialization part, the machine code is loaded
'from the HexCode strings. Then the hex values are converted
'and put to the MachineCode arrays and the pointerMC1 and
'pointerMC2 points to the 1st byte of these arrays.
Dim i As Long

ReDim MachineCode1(Len(HexCode1) \ 2 - 1)
ReDim MachineCode2(Len(HexCode2) \ 2 - 1)

For i = 0 To (Len(HexCode1) \ 2 - 1)
    MachineCode1(i) = CByte("&H" & Mid(HexCode1, i * 2 + 1, 2))
Next i


pointerMC1 = VarPtr(MachineCode1(0))

For i = 0 To (Len(HexCode2) \ 2 - 1)
    MachineCode2(i) = CByte("&H" & Mid(HexCode2, i * 2 + 1, 2))
Next i

pointerMC2 = VarPtr(MachineCode2(0))

End Sub


'Use this function to create the sound
'Duration is in milliseconds and it is optional
'If you don't specify Duration, it is then 300 milliseconds
Public Function Beep(Frequency As Long, Optional Duration As Long = 300)
Dim dl As Long

If Frequency < 0 Or Duration < 0 Then Exit Function  'Surely frequency and duration can't be nagative, can they?

If Duration = 0 Then Exit Function    'No need to proceed if if duration is 0 millisecond

If Frequency = 0 Then     'If frequency is 0 then no need to create sound, just sleep for the specified duration
    Sleep Duration
    Exit Function
End If

dl = CallWindowProc(pointerMC1, Frequency, 0&, 0&, 0&)   'Calling the first code and taking the return value in ret
Sleep Duration   'The speaker is on, keep it in this position for the specified amount of milliseconds

If dl <> 0 Then
    dl = CallWindowProc(pointerMC2, dl, 0&, 0&, 0&)   'Now calling the second code to turn off the speaker
End If

End Function
