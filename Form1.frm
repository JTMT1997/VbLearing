VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13665
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H00FFFF80&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   13665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   495
      Left            =   12240
      TabIndex        =   24
      Top             =   5280
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   3615
      Left            =   5040
      TabIndex        =   23
      Top             =   1440
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6376
      _Version        =   393216
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   6840
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   495
      Left            =   2160
      TabIndex        =   22
      Top             =   5280
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2880
      TabIndex        =   21
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   134938625
      CurrentDate     =   44672
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form1.frx":0015
      Left            =   2880
      List            =   "Form1.frx":0028
      TabIndex        =   20
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Cancel          =   -1  'True
      Caption         =   "New"
      Height          =   495
      Left            =   720
      MaskColor       =   &H00FFFF80&
      TabIndex        =   18
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2880
      TabIndex        =   17
      Top             =   4560
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form1.frx":0054
      Left            =   2880
      List            =   "Form1.frx":0061
      TabIndex        =   15
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":007D
      Left            =   2880
      List            =   "Form1.frx":0087
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   645
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFF80&
      Caption         =   "Agama"
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF80&
      Caption         =   "Kelas"
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF80&
      Caption         =   "Status"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF80&
      Caption         =   "Email"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      Caption         =   "No.HP"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      Caption         =   "Alamat"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF80&
      Caption         =   "Jenis Kelamin"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Tanggal Lahir"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Tempat Lahir"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Nama Lengkap"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Formulir Pendaftaran"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim koneksi As String
Sub clean()
Text1 = ""
Text2 = ""
Combo1 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Combo2 = " "
Combo3 = " "
Text6 = ""
End Sub

Sub readdata()
Dim result As String
result = "select * from mahasiswa order by NamaLengkap asc "
conn.Execute (result)
Adodc1.RecordSource = result
Set DataGrid1.DataSource = Adodc1
Adodc1.Refresh
End Sub
Sub keluarapps()
If MsgBox("Apakah Ingin keluar dari aplikasi? ", 36, "Informasi") = vbYes Then
            Unload Me
End If
End Sub
Private Sub Command1_Click()
If MsgBox("Apakah Ingin mereset formulir ? ", 36, "Informasi") = vbYes Then
            Call clean
Else
   Call keluarapps
End If
End Sub
Private Sub Command2_Click()
result = "select * from mahasiswa where NamaLengkap = '" & Text1 & "' "

    Set RS = conn.Execute(result)
    If RS.EOF Then
        If MsgBox("Data Akan Disimpan", 36, "Informasi") = vbYes Then
            saves = "insert into mahasiswa values('" & Text1 & "', '" & Text2 & "', '" & DTPicker1 & "', '" & Combo1 & "', '" & Text3 & "', '" & Text4 & "', '" & Text5 & "', '" & Combo2 & "', '" & Combo3 & "', '" & Text6 & "')"
            conn.Execute (saves)
            End If
  Call readdata
  Else
    MsgBox ("Nama anda sudah ada")
End If
Call clean
End Sub

Private Sub Command3_Click()
Call keluarapps
End Sub

Private Sub Form_Load()
koneksi = "provider = microsoft.jet.oledb.4.0; data source = " & App.Path & "\mahasiswa.mdb"
Adodc1.ConnectionString = koneksi
conn.Open koneksi
Call readdata
End Sub

