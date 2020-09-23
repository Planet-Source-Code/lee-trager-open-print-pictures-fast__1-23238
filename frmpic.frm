VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmpic 
   AutoRedraw      =   -1  'True
   Caption         =   "Picture Viewer"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdprint 
      Left            =   0
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picx 
      Height          =   10815
      Left            =   0
      ScaleHeight     =   10755
      ScaleWidth      =   15195
      TabIndex        =   0
      Top             =   0
      Width           =   15255
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuload 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close picture"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "frmpic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub load(drive As String, loc As String)
On Error GoTo err:
Dim pic As String
pic = drive & "\" & loc
err.Clear
On Error GoTo err:
picx.Picture = LoadPicture(pic)
err.Clear
frmpic.Caption = "Viewing Picture " & pic
err:
Exit Sub
End Sub

Private Sub Printer2()
cdprint.ShowPrinter
frmpic.BackColor = vbWhite
picx.BackColor = vbWhite
picx.BorderStyle = 0
frmpic.PrintForm
Printer.EndDoc
frmpic.BackColor = &H8000000F
picx.BackColor = &H8000000F
picx.BorderStyle = 0
End Sub

Public Sub Closer()
picx.Picture = LoadPicture()
frmpic.Caption = "View Picture"
frmload.Caption = "Load File"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End
End Sub

Private Sub mnuclose_Click()
Closer
End Sub

Private Sub mnuload_Click()
frmload.Show
End Sub

Private Sub mnuprint_Click()
Printer2
End Sub

Private Sub mnuquit_Click()
Unload Me
End
End Sub
