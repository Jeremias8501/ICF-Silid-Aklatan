VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000B&
   Caption         =   "ICF Silid Aklatan   "
   ClientHeight    =   8955
   ClientLeft      =   2535
   ClientTop       =   1530
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   14925
   Begin VB.Timer timDsiplay 
      Interval        =   1000
      Left            =   13200
      Top             =   840
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   9720
      Top             =   3960
      Width           =   3855
   End
   Begin VB.Label lblYear 
      BackStyle       =   0  'Transparent
      Caption         =   "2022"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1095
      Left            =   7560
      TabIndex        =   4
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label lblMonth 
      BackStyle       =   0  'Transparent
      Caption         =   "Jan"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1215
      Left            =   6480
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblDay 
      BackStyle       =   0  'Transparent
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1095
      Left            =   6600
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblNumber 
      BackStyle       =   0  'Transparent
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   90
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1935
      Left            =   6960
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 PM"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   2175
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
   Begin VB.Menu MnuRegister 
      Caption         =   "&REGISTER"
      Begin VB.Menu MnuStudent 
         Caption         =   "&Student"
      End
      Begin VB.Menu MnuBook 
         Caption         =   "&Book"
      End
   End
   Begin VB.Menu MnuBorrow 
      Caption         =   "&BORROW"
   End
   Begin VB.Menu MnuReturn 
      Caption         =   "&RETURN BOOKS"
   End
   Begin VB.Menu Records 
      Caption         =   "&RECORDS"
      Begin VB.Menu Student_History 
         Caption         =   "&Student History"
      End
      Begin VB.Menu Book_History 
         Caption         =   "&Book History"
      End
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "&ABOUT"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu MnuTools 
      Caption         =   "&TOOLS"
      Begin VB.Menu Calendar 
         Caption         =   "&Calendar"
      End
   End
   Begin VB.Menu MnuLogout 
      Caption         =   "&LOGOUT"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MnuBook_Click()
frmBook.Show
Me.Hide
End Sub

Private Sub MnuBorrow_Click()

End Sub

Private Sub MnuLogout_Click()
fraLogin.Show
Me.Hide
End Sub

Private Sub MnuReturn_Click()

End Sub

Private Sub MnuStudent_Click()
frmStudent.Show
Me.Hide
End Sub

Private Sub timDsiplay_Timer()
Dim Today As Variant
Today = Now
lblDay.Caption = Format(Today, "dddd")
lblMonth.Caption = Format(Today, "mmm")
lblYear.Caption = Format(Today, "yyyy")
lblNumber.Caption = Format(Today, "d")
lblTime.Caption = Format(Today, "h:mm:ss ampm")
End Sub
