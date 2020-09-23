VERSION 5.00
Begin VB.Form frmload 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load File"
   ClientHeight    =   6570
   ClientLeft      =   10425
   ClientTop       =   2895
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   4680
   Begin VB.CommandButton bntgoto 
      Caption         =   "&Goto dir"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   4680
      Width           =   4695
   End
   Begin VB.TextBox txtgoto 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   4455
   End
   Begin VB.CommandButton bntclose 
      Caption         =   "&Close Picture"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   6240
      Width           =   4695
   End
   Begin VB.CommandButton bntshow 
      Caption         =   "&Show them only"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   6000
      Width           =   4695
   End
   Begin VB.TextBox txtfile 
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Text            =   "*.*"
      Top             =   5640
      Width           =   4695
   End
   Begin VB.CheckBox ckopen 
      Caption         =   "&Stay open"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   4455
   End
   Begin VB.CommandButton bntcancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton bntok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblgoto 
      Caption         =   "Goto dir:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   4455
   End
   Begin VB.Label lblshow 
      Caption         =   $"frmload.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   4455
   End
End
Attribute VB_Name = "frmload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bntcancel_Click()
Unload Me
End Sub

Private Sub bntclose_Click()
frmpic.Closer
End Sub

Private Sub bntgoto_Click()
Dir1.Path = txtgoto.Text
File1.Path = txtgoto.Text
End Sub

Private Sub bntok_Click()
frmpic.load Dir1.Path, File1.FileName
If ckopen.Value = "1" Then

Else
Unload Me
End If
End Sub

Private Sub bntshow_Click()
File1.Pattern = txtfile.Text
End Sub

Private Sub Dir1_Change()
On Error GoTo err:
File1.Path = Dir1.Path
err:
Exit Sub
End Sub

Private Sub Drive1_Change()
On Error GoTo err:
Dir1.Path = Drive1.drive
err:
Exit Sub
End Sub

Private Sub File1_Click()
frmload.Caption = "Load File: " & Dir1.Path & "\" & File1.FileName
End Sub

Private Sub File1_DblClick()
frmpic.load Dir1.Path, File1.FileName
End Sub

Private Sub Form_Load()
frmpic.Show
End Sub
