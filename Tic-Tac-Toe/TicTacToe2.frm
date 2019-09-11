VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TicTacToe2 
   Caption         =   "Tic-Tac-Toe"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2655
   OleObjectBlob   =   "TicTacToe2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TicTacToe2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cb0_Click()

cb1.Caption = ""
cb2.Caption = ""
cb3.Caption = ""
cb4.Caption = ""
cb5.Caption = ""
cb6.Caption = ""
cb7.Caption = ""
cb8.Caption = ""
cb9.Caption = ""

cb1.Locked = False
cb2.Locked = False
cb3.Locked = False
cb4.Locked = False
cb5.Locked = False
cb6.Locked = False
cb7.Locked = False
cb8.Locked = False
cb9.Locked = False


End Sub

Private Sub cb1_Click()

With cb1
    .Caption = "X"
    .ForeColor = &HFF&
    .Locked = True
End With


Dim a As MSForms.CommandButton
Do
Randomize
random_number = Int(9 * Rnd) + 1
Set a = Me.Controls("cb" & random_number)

If a.Caption = "" Then
    With a
        .Caption = "O"
        .ForeColor = &H8000000D
        .Locked = True
    End With
    Exit Do
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    Exit Do
    MsgBox ("DRAW")
End If
Loop

Dim HL1, HL2, HL3, VL1, VL2, VL3, DL1, DL2, all As String
HL1 = cb1.Caption & cb2.Caption & cb3.Caption
HL2 = cb4.Caption & cb5.Caption & cb6.Caption
HL3 = cb7.Caption & cb8.Caption & cb9.Caption

VL1 = cb1.Caption & cb4.Caption & cb7.Caption
VL2 = cb2.Caption & cb5.Caption & cb8.Caption
VL3 = cb3.Caption & cb6.Caption & cb9.Caption

DL1 = cb1.Caption & cb5.Caption & cb9.Caption
DL2 = cb3.Caption & cb5.Caption & cb7.Caption
all = cb1.Caption & cb2.Caption & cb3.Caption & cb4.Caption & cb5.Caption & cb6.Caption & cb7.Caption & cb8.Caption & cb9.Caption

If HL1 = "XXX" Or HL2 = "XXX" Or HL3 = "XXX" Or VL1 = "XXX" Or VL2 = "XXX" Or VL3 = "XXX" Or DL1 = "XXX" Or DL2 = "XXX" Then
    MsgBox ("'X' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf HL1 = "OOO" Or HL2 = "OOO" Or HL3 = "OOO" Or VL1 = "OOO" Or VL2 = "OOO" Or VL3 = "OOO" Or DL1 = "OOO" Or DL2 = "OOO" Then
    MsgBox ("'O' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    MsgBox ("DRAW")
End If



End Sub

Private Sub cb2_Click()
With cb2
    .Caption = "X"
    .ForeColor = &HFF&
    .Locked = True
End With

Dim a As MSForms.CommandButton
Do
Randomize
random_number = Int(9 * Rnd) + 1
Set a = Me.Controls("cb" & random_number)

If a.Caption = "" Then
    With a
        .Caption = "O"
        .ForeColor = &H8000000D
        .Locked = True
    End With
    Exit Do
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    Exit Do
    MsgBox ("DRAW")
End If
Loop

Dim HL1, HL2, HL3, VL1, VL2, VL3, DL1, DL2, all As String
HL1 = cb1.Caption & cb2.Caption & cb3.Caption
HL2 = cb4.Caption & cb5.Caption & cb6.Caption
HL3 = cb7.Caption & cb8.Caption & cb9.Caption

VL1 = cb1.Caption & cb4.Caption & cb7.Caption
VL2 = cb2.Caption & cb5.Caption & cb8.Caption
VL3 = cb3.Caption & cb6.Caption & cb9.Caption

DL1 = cb1.Caption & cb5.Caption & cb9.Caption
DL2 = cb3.Caption & cb5.Caption & cb7.Caption
all = cb1.Caption & cb2.Caption & cb3.Caption & cb4.Caption & cb5.Caption & cb6.Caption & cb7.Caption & cb8.Caption & cb9.Caption

If HL1 = "XXX" Or HL2 = "XXX" Or HL3 = "XXX" Or VL1 = "XXX" Or VL2 = "XXX" Or VL3 = "XXX" Or DL1 = "XXX" Or DL2 = "XXX" Then
    MsgBox ("'X' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf HL1 = "OOO" Or HL2 = "OOO" Or HL3 = "OOO" Or VL1 = "OOO" Or VL2 = "OOO" Or VL3 = "OOO" Or DL1 = "OOO" Or DL2 = "OOO" Then
    MsgBox ("'O' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    MsgBox ("DRAW")
End If


End Sub

Private Sub cb3_Click()
With cb3
    .Caption = "X"
    .ForeColor = &HFF&
    .Locked = True
End With


Dim a As MSForms.CommandButton
Do
Randomize
random_number = Int(9 * Rnd) + 1
Set a = Me.Controls("cb" & random_number)

If a.Caption = "" Then
    With a
        .Caption = "O"
        .ForeColor = &H8000000D
        .Locked = True
    End With
    Exit Do
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    Exit Do
    MsgBox ("DRAW")
End If
Loop

Dim HL1, HL2, HL3, VL1, VL2, VL3, DL1, DL2, all As String
HL1 = cb1.Caption & cb2.Caption & cb3.Caption
HL2 = cb4.Caption & cb5.Caption & cb6.Caption
HL3 = cb7.Caption & cb8.Caption & cb9.Caption

VL1 = cb1.Caption & cb4.Caption & cb7.Caption
VL2 = cb2.Caption & cb5.Caption & cb8.Caption
VL3 = cb3.Caption & cb6.Caption & cb9.Caption

DL1 = cb1.Caption & cb5.Caption & cb9.Caption
DL2 = cb3.Caption & cb5.Caption & cb7.Caption
all = cb1.Caption & cb2.Caption & cb3.Caption & cb4.Caption & cb5.Caption & cb6.Caption & cb7.Caption & cb8.Caption & cb9.Caption

If HL1 = "XXX" Or HL2 = "XXX" Or HL3 = "XXX" Or VL1 = "XXX" Or VL2 = "XXX" Or VL3 = "XXX" Or DL1 = "XXX" Or DL2 = "XXX" Then
    MsgBox ("'X' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf HL1 = "OOO" Or HL2 = "OOO" Or HL3 = "OOO" Or VL1 = "OOO" Or VL2 = "OOO" Or VL3 = "OOO" Or DL1 = "OOO" Or DL2 = "OOO" Then
    MsgBox ("'O' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    MsgBox ("DRAW")
End If

End Sub

Private Sub cb4_Click()
With cb4
    .Caption = "X"
    .ForeColor = &HFF&
    .Locked = True
End With

Dim a As MSForms.CommandButton
Do
Randomize
random_number = Int(9 * Rnd) + 1
Set a = Me.Controls("cb" & random_number)

If a.Caption = "" Then
    With a
        .Caption = "O"
        .ForeColor = &H8000000D
        .Locked = True
    End With
    Exit Do
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    Exit Do
    MsgBox ("DRAW")
End If
Loop

Dim HL1, HL2, HL3, VL1, VL2, VL3, DL1, DL2, all As String
HL1 = cb1.Caption & cb2.Caption & cb3.Caption
HL2 = cb4.Caption & cb5.Caption & cb6.Caption
HL3 = cb7.Caption & cb8.Caption & cb9.Caption

VL1 = cb1.Caption & cb4.Caption & cb7.Caption
VL2 = cb2.Caption & cb5.Caption & cb8.Caption
VL3 = cb3.Caption & cb6.Caption & cb9.Caption

DL1 = cb1.Caption & cb5.Caption & cb9.Caption
DL2 = cb3.Caption & cb5.Caption & cb7.Caption
all = cb1.Caption & cb2.Caption & cb3.Caption & cb4.Caption & cb5.Caption & cb6.Caption & cb7.Caption & cb8.Caption & cb9.Caption

If HL1 = "XXX" Or HL2 = "XXX" Or HL3 = "XXX" Or VL1 = "XXX" Or VL2 = "XXX" Or VL3 = "XXX" Or DL1 = "XXX" Or DL2 = "XXX" Then
    MsgBox ("'X' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf HL1 = "OOO" Or HL2 = "OOO" Or HL3 = "OOO" Or VL1 = "OOO" Or VL2 = "OOO" Or VL3 = "OOO" Or DL1 = "OOO" Or DL2 = "OOO" Then
    MsgBox ("'O' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    MsgBox ("DRAW")
End If


End Sub

Private Sub cb5_Click()
With cb5
    .Caption = "X"
    .ForeColor = &HFF&
    .Locked = True
End With

Dim a As MSForms.CommandButton
Do
Randomize
random_number = Int(9 * Rnd) + 1
Set a = Me.Controls("cb" & random_number)

If a.Caption = "" Then
    With a
        .Caption = "O"
        .ForeColor = &H8000000D
        .Locked = True
    End With
    Exit Do
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    Exit Do
    MsgBox ("DRAW")
End If
Loop

Dim HL1, HL2, HL3, VL1, VL2, VL3, DL1, DL2, all As String
HL1 = cb1.Caption & cb2.Caption & cb3.Caption
HL2 = cb4.Caption & cb5.Caption & cb6.Caption
HL3 = cb7.Caption & cb8.Caption & cb9.Caption

VL1 = cb1.Caption & cb4.Caption & cb7.Caption
VL2 = cb2.Caption & cb5.Caption & cb8.Caption
VL3 = cb3.Caption & cb6.Caption & cb9.Caption

DL1 = cb1.Caption & cb5.Caption & cb9.Caption
DL2 = cb3.Caption & cb5.Caption & cb7.Caption
all = cb1.Caption & cb2.Caption & cb3.Caption & cb4.Caption & cb5.Caption & cb6.Caption & cb7.Caption & cb8.Caption & cb9.Caption

If HL1 = "XXX" Or HL2 = "XXX" Or HL3 = "XXX" Or VL1 = "XXX" Or VL2 = "XXX" Or VL3 = "XXX" Or DL1 = "XXX" Or DL2 = "XXX" Then
    MsgBox ("'X' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf HL1 = "OOO" Or HL2 = "OOO" Or HL3 = "OOO" Or VL1 = "OOO" Or VL2 = "OOO" Or VL3 = "OOO" Or DL1 = "OOO" Or DL2 = "OOO" Then
    MsgBox ("'O' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    MsgBox ("DRAW")
End If


End Sub

Private Sub cb6_Click()
With cb6
    .Caption = "X"
    .ForeColor = &HFF&
    .Locked = True
End With

Dim a As MSForms.CommandButton
Do
Randomize
random_number = Int(9 * Rnd) + 1
Set a = Me.Controls("cb" & random_number)

If a.Caption = "" Then
    With a
        .Caption = "O"
        .ForeColor = &H8000000D
        .Locked = True
    End With
    Exit Do
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    Exit Do
    MsgBox ("DRAW")
End If
Loop

Dim HL1, HL2, HL3, VL1, VL2, VL3, DL1, DL2, all As String
HL1 = cb1.Caption & cb2.Caption & cb3.Caption
HL2 = cb4.Caption & cb5.Caption & cb6.Caption
HL3 = cb7.Caption & cb8.Caption & cb9.Caption

VL1 = cb1.Caption & cb4.Caption & cb7.Caption
VL2 = cb2.Caption & cb5.Caption & cb8.Caption
VL3 = cb3.Caption & cb6.Caption & cb9.Caption

DL1 = cb1.Caption & cb5.Caption & cb9.Caption
DL2 = cb3.Caption & cb5.Caption & cb7.Caption
all = cb1.Caption & cb2.Caption & cb3.Caption & cb4.Caption & cb5.Caption & cb6.Caption & cb7.Caption & cb8.Caption & cb9.Caption

If HL1 = "XXX" Or HL2 = "XXX" Or HL3 = "XXX" Or VL1 = "XXX" Or VL2 = "XXX" Or VL3 = "XXX" Or DL1 = "XXX" Or DL2 = "XXX" Then
    MsgBox ("'X' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf HL1 = "OOO" Or HL2 = "OOO" Or HL3 = "OOO" Or VL1 = "OOO" Or VL2 = "OOO" Or VL3 = "OOO" Or DL1 = "OOO" Or DL2 = "OOO" Then
    MsgBox ("'O' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    MsgBox ("DRAW")
End If

End Sub

Private Sub cb7_Click()
With cb7
    .Caption = "X"
    .ForeColor = &HFF&
    .Locked = True
End With

Dim a As MSForms.CommandButton
Do
Randomize
random_number = Int(9 * Rnd) + 1
Set a = Me.Controls("cb" & random_number)

If a.Caption = "" Then
    With a
        .Caption = "O"
        .ForeColor = &H8000000D
        .Locked = True
    End With
    Exit Do
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    Exit Do
    MsgBox ("DRAW")
End If
Loop

Dim HL1, HL2, HL3, VL1, VL2, VL3, DL1, DL2, all As String
HL1 = cb1.Caption & cb2.Caption & cb3.Caption
HL2 = cb4.Caption & cb5.Caption & cb6.Caption
HL3 = cb7.Caption & cb8.Caption & cb9.Caption

VL1 = cb1.Caption & cb4.Caption & cb7.Caption
VL2 = cb2.Caption & cb5.Caption & cb8.Caption
VL3 = cb3.Caption & cb6.Caption & cb9.Caption

DL1 = cb1.Caption & cb5.Caption & cb9.Caption
DL2 = cb3.Caption & cb5.Caption & cb7.Caption
all = cb1.Caption & cb2.Caption & cb3.Caption & cb4.Caption & cb5.Caption & cb6.Caption & cb7.Caption & cb8.Caption & cb9.Caption

If HL1 = "XXX" Or HL2 = "XXX" Or HL3 = "XXX" Or VL1 = "XXX" Or VL2 = "XXX" Or VL3 = "XXX" Or DL1 = "XXX" Or DL2 = "XXX" Then
    MsgBox ("'X' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf HL1 = "OOO" Or HL2 = "OOO" Or HL3 = "OOO" Or VL1 = "OOO" Or VL2 = "OOO" Or VL3 = "OOO" Or DL1 = "OOO" Or DL2 = "OOO" Then
    MsgBox ("'O' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    MsgBox ("DRAW")
End If


End Sub

Private Sub cb8_Click()
With cb8
    .Caption = "X"
    .ForeColor = &HFF&
    .Locked = True
End With

Dim a As MSForms.CommandButton
Do
Randomize
random_number = Int(9 * Rnd) + 1
Set a = Me.Controls("cb" & random_number)

If a.Caption = "" Then
    With a
        .Caption = "O"
        .ForeColor = &H8000000D
        .Locked = True
    End With
    Exit Do
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    Exit Do
    MsgBox ("DRAW")
End If
Loop

Dim HL1, HL2, HL3, VL1, VL2, VL3, DL1, DL2, all As String
HL1 = cb1.Caption & cb2.Caption & cb3.Caption
HL2 = cb4.Caption & cb5.Caption & cb6.Caption
HL3 = cb7.Caption & cb8.Caption & cb9.Caption

VL1 = cb1.Caption & cb4.Caption & cb7.Caption
VL2 = cb2.Caption & cb5.Caption & cb8.Caption
VL3 = cb3.Caption & cb6.Caption & cb9.Caption

DL1 = cb1.Caption & cb5.Caption & cb9.Caption
DL2 = cb3.Caption & cb5.Caption & cb7.Caption
all = cb1.Caption & cb2.Caption & cb3.Caption & cb4.Caption & cb5.Caption & cb6.Caption & cb7.Caption & cb8.Caption & cb9.Caption

If HL1 = "XXX" Or HL2 = "XXX" Or HL3 = "XXX" Or VL1 = "XXX" Or VL2 = "XXX" Or VL3 = "XXX" Or DL1 = "XXX" Or DL2 = "XXX" Then
    MsgBox ("'X' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf HL1 = "OOO" Or HL2 = "OOO" Or HL3 = "OOO" Or VL1 = "OOO" Or VL2 = "OOO" Or VL3 = "OOO" Or DL1 = "OOO" Or DL2 = "OOO" Then
    MsgBox ("'O' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    MsgBox ("DRAW")
End If


End Sub

Private Sub cb9_Click()
With cb9
    .Caption = "X"
    .ForeColor = &HFF&
    .Locked = True
End With

Dim a As MSForms.CommandButton
Do
Randomize
random_number = Int(9 * Rnd) + 1
Set a = Me.Controls("cb" & random_number)

If a.Caption = "" Then
    With a
        .Caption = "O"
        .ForeColor = &H8000000D
        .Locked = True
    End With
    Exit Do
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    Exit Do
    MsgBox ("DRAW")
End If
Loop

Dim HL1, HL2, HL3, VL1, VL2, VL3, DL1, DL2, all As String
HL1 = cb1.Caption & cb2.Caption & cb3.Caption
HL2 = cb4.Caption & cb5.Caption & cb6.Caption
HL3 = cb7.Caption & cb8.Caption & cb9.Caption

VL1 = cb1.Caption & cb4.Caption & cb7.Caption
VL2 = cb2.Caption & cb5.Caption & cb8.Caption
VL3 = cb3.Caption & cb6.Caption & cb9.Caption

DL1 = cb1.Caption & cb5.Caption & cb9.Caption
DL2 = cb3.Caption & cb5.Caption & cb7.Caption
all = cb1.Caption & cb2.Caption & cb3.Caption & cb4.Caption & cb5.Caption & cb6.Caption & cb7.Caption & cb8.Caption & cb9.Caption

If HL1 = "XXX" Or HL2 = "XXX" Or HL3 = "XXX" Or VL1 = "XXX" Or VL2 = "XXX" Or VL3 = "XXX" Or DL1 = "XXX" Or DL2 = "XXX" Then
    MsgBox ("'X' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf HL1 = "OOO" Or HL2 = "OOO" Or HL3 = "OOO" Or VL1 = "OOO" Or VL2 = "OOO" Or VL3 = "OOO" Or DL1 = "OOO" Or DL2 = "OOO" Then
    MsgBox ("'O' WIN!")
    cb1.Locked = True
    cb2.Locked = True
    cb3.Locked = True
    cb4.Locked = True
    cb5.Locked = True
    cb6.Locked = True
    cb7.Locked = True
    cb8.Locked = True
    cb9.Locked = True
ElseIf cb1.Caption <> "" And cb2.Caption <> "" And cb3.Caption <> "" And cb4.Caption <> "" And cb5.Caption <> "" And cb6.Caption <> "" And cb7.Caption <> "" And cb8.Caption <> "" And cb9.Caption <> "" Then
    MsgBox ("DRAW")
End If



End Sub

Private Sub UserForm_Click()

End Sub
