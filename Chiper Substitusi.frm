VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Chiper Substitusi"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7470
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Chiper Substitusi"
   ScaleHeight     =   7425
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000A&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      MaskColor       =   &H00404040&
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000A&
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MaskColor       =   &H00808080&
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Deskripsi"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enkripsi"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Deskripsi"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   12
      Top             =   4440
      Width           =   6735
      Begin VB.Label Label4 
         Caption         =   "Chipertext"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Plaintext"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enkripsi"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   11
      Top             =   3120
      Width           =   6735
      Begin VB.Label Label8 
         Caption         =   "Chipertext"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Plaintext"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Kelompok 3"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "1-26"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Enkripsi dan Deskripsi Text Dengan Algoritma Substitusi Chiper"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   7095
   End
   Begin VB.Label Label2 
      Caption         =   "plaintext"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Inputkan Bit Geser"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text5.Text = "" Then
MsgBox "Anda Belum Menginputkan Bit Geser...!!", vbCritical, "Warning"
Text5.SetFocus
End If
Var = Text5.Text
Text2 = ""
If Var < 26 Then
N = Len(Trim(Text1))
For i = 1 To N
C = Mid(Text1, i, 1)
P = Chr(Asc(C) + Var)
If Asc(P) > 90 Then
Text2.SelText = Chr(64 + (Asc(P) - 90))
Else
Text2.SelText = P
Text1.Enabled = True
End If
Next
Else
MsgBox " Maksimal Bit Hanya 26...!!!", vbOKOnly + vbInformation, "Peringatan"
Text5 = ""
End If
End Sub

Private Sub Command2_Click()
Var = Text5.Text
Text4 = ""
If Var < 26 Then
N = Len(Trim(Text3))
For i = 1 To N
C = Mid(Text3, i, 1)
P = Chr(Asc(C) + 26 - Var)
If Asc(P) > 90 Then
Text4.SelText = Chr(64 + (Asc(P) - 90))
Else
Text4.SelText = P
End If
Next
Else
MsgBox " Data Anda Salah..!!!", vbOKOnly + vbInformation, "Peringatan"
Text5 = ""
End If
End Sub

Private Sub Command3_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text5.SetFocus
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Text1.SetFocus
If KeyAscii = 8 Then Exit Sub
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not (KeyAscii >= 65 And KeyAscii <= 90) Then
KeyAscii = 0
End If
End Sub
 
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not (KeyAscii >= 65 And KeyAscii <= 90) Then
KeyAscii = 0
End If
End Sub
 
Private Sub Text3_KeyPress(KeyAscii As Integer)
Text3.SetFocus
If KeyAscii = 8 Then Exit Sub
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not (KeyAscii >= 65 And KeyAscii <= 90) Then
KeyAscii = 0
End If
End Sub

