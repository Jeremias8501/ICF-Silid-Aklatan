VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fraLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   8985
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   5308.635
   ScaleMode       =   0  'User
   ScaleWidth      =   14070.1
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adodclogin 
      Height          =   495
      Left            =   5160
      Top             =   7920
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\maste\OneDrive\Documents\ICF SILID AKLATAN.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\maste\OneDrive\Documents\ICF SILID AKLATAN.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select* from Admin"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&LOGIN"
      Default         =   -1  'True
      Height          =   390
      Left            =   6840
      TabIndex        =   1
      ToolTipText     =   "Login"
      Top             =   5880
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&EXIT"
      Height          =   390
      Left            =   8160
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   5880
      Width           =   1140
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   6840
      TabIndex        =   0
      Top             =   4200
      Width           =   3405
   End
   Begin VB.Frame fraLogin 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   3600
      TabIndex        =   3
      Top             =   1440
      Width           =   7815
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "*"
         TabIndex        =   8
         ToolTipText     =   "Please enter your Password!"
         Top             =   3480
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Create New Account"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3240
         TabIndex        =   7
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label lblUsername 
         BackStyle       =   0  'Transparent
         Caption         =   "&Username:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ICF Silid Aklatan"
         BeginProperty Font 
            Name            =   "Niagara Solid"
            Size            =   84.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   6495
      End
   End
End
Attribute VB_Name = "fraLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
adodclogin.RecordSource = "Select*from Admin where Username ='" + txtUsername.Text + "'and Password ='" + txtPassword.Text + "'"
adodclogin.Refresh
If adodclogin.Recordset.EOF Then
MsgBox "Login Failed, Try Again...!!!!", vbCritical, "Please Enter the correct Username and Password "

Else
MsgBox "Login Successful.", vbInformation, "Successful Attempt"
frmProgress.Show
Unload Me

    End If
End Sub


Private Sub Label1_Click()
Me.Hide
fraRegister.Show
End Sub

