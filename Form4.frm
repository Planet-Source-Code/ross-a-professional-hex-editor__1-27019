VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Bytes"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2505
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   2505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add Bytes"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox BytesToAdd 
      Height          =   375
      Left            =   120
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BytesToAdd_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) <> vbBack Then                 'Check for backspace key
    If (KeyAscii >= 48 And KeyAscii <= 57) Then 'Make sure only numbers can be entered
        DoEvents
    Else
        KeyAscii = 0
    End If
End If
End Sub

Private Sub CmdAdd_Click()
If BytesToAdd.Text < 1 Then Exit Sub            'check bytes are greater than 0
AddBytesToEnd (BytesToAdd.Text)                 'Call function to add bytes
Form1.SortHex                                   'Sort hex display
Unload Me                                       'Unload this form
End Sub

Private Sub CmdCancel_Click()
Unload Me                                       'Unload this form
End Sub

Private Sub Form_Load()
AlwaysOnTop Me, True                            'Set alway on top
End Sub
