VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuat Form Sebagai Array"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'form baru yang merupakan array dari Form1
Private Sub Command1_Click()
   Dim intX As Integer
   Dim frmNew(1 To 5) As New Form1
   For intX = 1 To 5
      frmNew(intX).Show
      frmNew(intX).WindowState = vbMinimized
      'Untuk membuat form yang diminimized tanpa
      'memiliki ukuran normal pada saat tampilan
      'awalnya, ganti urutan coding dari dua baris di
      'atas, sehingga nantinya menjadi:
      'frmNew(intX).WindowState = vbMinimized
      'frmNew(intX).Show
   Next
End Sub


