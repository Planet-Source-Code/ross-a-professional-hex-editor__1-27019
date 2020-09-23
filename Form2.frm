VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Goto Byte"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox ByteText 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Byte number:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ByteText_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) <> vbBack Then                 'check key is not backspace
    If (KeyAscii >= 48 And KeyAscii <= 57) Then 'make sure only numbers can be entered
        DoEvents
    Else
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Command1_Click()
Dim Temp As String
If ByteText.Text > SizeOfFile Then Exit Sub     'check byte is not bigger than size of file
CurrentPos = ByteText.Text                      'get current position
Form1.Edit.Visible = True                       'make edit invisible
Form1.SortHex                                   'sort hex displayed
Form1.Edit.Left = 0                             'set left as 0
Form1.Showtxt.Visible = False                   'make showtxt invisible
Form1.Edit.Top = 0                              'set top as 0
Temp = Hex(HexDisplayed(1))                     'get first hex value
If Len(Temp) = 1 Then Temp = "0" & Temp         'make it 2 chars long
Form1.Edit.Text = Temp                          'set edit text as temp
Unload Me                                       'unload this form
End Sub

Private Sub Command2_Click()
Unload Me                                       'unload this form
End Sub

Private Sub Form_Load()
AlwaysOnTop Me, True                            'make this form always on top
End Sub

