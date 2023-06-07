VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fraRegister 
   Caption         =   "Be a Member!"
   ClientHeight    =   8985
   ClientLeft      =   2685
   ClientTop       =   1380
   ClientWidth     =   14955
   LinkTopic       =   "Form2"
   Picture         =   "fraRegister.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   14955
   Begin VB.Frame fraRegister 
      BackColor       =   &H00004080&
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   6855
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   11895
      Begin VB.TextBox txtUsername 
         DataField       =   "Username"
         DataSource      =   "adodcRecord"
         Height          =   495
         Left            =   6000
         TabIndex        =   14
         Top             =   3840
         Width           =   2775
      End
      Begin VB.TextBox txtAddress 
         DataField       =   "Address"
         DataSource      =   "adodcRecord"
         Height          =   495
         Left            =   6000
         TabIndex        =   7
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtNumber 
         DataField       =   "ContactNumber"
         DataSource      =   "adodcRecord"
         Height          =   495
         Left            =   6000
         TabIndex        =   6
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtPassword 
         DataField       =   "Password"
         DataSource      =   "adodcRecord"
         Height          =   495
         Left            =   6000
         TabIndex        =   5
         Top             =   4440
         Width           =   2775
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   615
         Left            =   4320
         TabIndex        =   4
         Top             =   5400
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   615
         Left            =   6360
         TabIndex        =   3
         Top             =   5400
         Width           =   1815
      End
      Begin VB.TextBox txtFname 
         DataField       =   "Fname"
         DataSource      =   "adodcRecord"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   2
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtLname 
         DataField       =   "Lname"
         DataSource      =   "adodcRecord"
         Height          =   495
         Left            =   6000
         TabIndex        =   1
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblFname 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Segoe UI Emoji"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3480
         TabIndex        =   13
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblLname 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Segoe UI Emoji"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3480
         TabIndex        =   12
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Segoe UI Emoji"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3480
         TabIndex        =   11
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblContact 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number:"
         BeginProperty Font 
            Name            =   "Segoe UI Emoji"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblUsername 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Segoe UI Emoji"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3480
         TabIndex        =   9
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Segoe UI Emoji"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3480
         TabIndex        =   8
         Top             =   4440
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc adodcRecord 
      Height          =   3375
      Left            =   13440
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   5953
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483624
      ForeColor       =   255
      Orientation     =   1
      Enabled         =   -1
      Connect         =   $"fraRegister.frx":3274F
      OLEDBString     =   $"fraRegister.frx":327F2
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Admin"
      Caption         =   "Students Record"
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
End
Attribute VB_Name = "fraRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

adodcRecord.Recordset.Fields("Fname") = txtFname.Text
adodcRecord.Recordset.Fields("Lname") = txtLname.Text
adodcRecord.Recordset.Fields("ContactNumber") = txtNumber.Text
adodcRecord.Recordset.Fields("Address") = txtAddress.Text
adodcRecord.Recordset.Fields("Password") = txtPassword.Text
adodcRecord.Recordset.Fields("Username") = txtUsername.Text
adodcRecord.Recordset.Update

If textFname = "" And txtLname = "" And txtNumber = "" And txtAddress = "" And txtPassword = "" And txtUsername = "" Then
MsgBox "Are you sure?", vbokayonly + vbInformation, "registatrion"
Else
If MsgBox("are you sure?", vbYesNo + vbQuestion, "question1") = vbYes Then

txtFname.Enabled = False
txtLname.Enabled = False
txtNumber.Enabled = False
txtAddress.Enabled = False
txtPassword.Enabled = False
txtUsername.Enabled = False

cmdAdd.Enabled = False
End If
End If

MsgBox "Added Successfully", vbOKOnly, "Successfully Added"
End Sub

Private Sub cmdCancel_Click()
Me.Hide
fraLogin.Show
End Sub

Private Sub Form_Load()
adodcRecord.Recordset.AddNew
End Sub

