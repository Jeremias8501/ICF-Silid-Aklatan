VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBook 
   Caption         =   "Book Section"
   ClientHeight    =   8985
   ClientLeft      =   2685
   ClientTop       =   1230
   ClientWidth     =   14955
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   Picture         =   "frmBook.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   14955
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   15
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&CLEAR"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&UPDATE"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   12
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox txtStatus 
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   4800
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc BookRecord 
      Height          =   615
      Left            =   12120
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
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
      RecordSource    =   "BookRecord"
      Caption         =   "BookRecord"
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
   Begin VB.TextBox txtPublication 
      DataField       =   "PublicationName"
      DataSource      =   "BookRecord"
      Height          =   405
      Left            =   2760
      TabIndex        =   9
      ToolTipText     =   "Publication"
      Top             =   3960
      Width           =   4335
   End
   Begin VB.TextBox txtAuthor 
      DataField       =   "Author"
      DataSource      =   "BookRecord"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      ToolTipText     =   "Author"
      Top             =   3120
      Width           =   4335
   End
   Begin VB.TextBox txtTitle 
      DataField       =   "BookTitle"
      DataSource      =   "BookRecord"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      ToolTipText     =   "Book Title"
      Top             =   2280
      Width           =   4335
   End
   Begin VB.TextBox txtID 
      DataField       =   "BookID"
      DataSource      =   "BookRecord"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      ToolTipText     =   "Book ID"
      Top             =   1440
      Width           =   4335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmBook.frx":3274F
      Height          =   5055
      Left            =   7440
      TabIndex        =   4
      Top             =   1440
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8916
      _Version        =   393216
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSearch 
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   9015
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  'Transparent
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   13440
      TabIndex        =   17
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Author:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblPublication 
      BackStyle       =   0  'Transparent
      Caption         =   "Publication:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Title:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblBookID 
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "frmBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Adodc1.Recordset.Fields("BookID") = txtID.Text
Adodc1.Recordset.Fields("BookTitle") = txtTitle.Text
Adodc1.Recordset.Fields("Author") = txtAuthor.Text
Adodc1.Recordset.Fields("PublicationName") = txtPublication.Text
Adodc1.Recordset.Fields("Status") = txtStatus.Text
Adodc1.Recordset.MoveNext
MsgBox "Added Successfully"
Exit Sub
On Error GoTo errmag
errmag:
MsgBox "Adding Error!"
End Sub

Private Sub cmdClear_Click()
txtID.Text = ""
txtTitle.Text = ""
txtAuthor.Text = ""
txtPublication.Text = ""
txtStatus.Text = ""
End Sub

Private Sub cmdUpdate_Click()
Adodc1.Recordset.Fields("BookID") = txtID.Text
Adodc1.Recordset.Fields("BookTitle") = txtTitle.Text
Adodc1.Recordset.Fields("Author") = txtAuthor.Text
Adodc1.Recordset.Fields("PublicationName") = txtPublication.Text
Adodc1.Recordset.Fields("Status") = txtStatus.Text
Adodc1.Recordset.Update
MsgBox "You successfully updated the information"
End Sub

Private Sub Form_Load()
BookRecord.Recordset.AddNew
End Sub

Private Sub lblBack_Click()
frmMain.Show
Me.Hide
End Sub
