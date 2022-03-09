VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proses Awal Identifikasi dengan JST-LM"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6600
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "inisialisasi_lm.frx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "KEMBALI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   28
      Top             =   6360
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "inisialisasi_lm.frx":6D8E6
      Height          =   1455
      Left            =   7560
      TabIndex        =   20
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2566
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
      Caption         =   "PARAMETER MODEL TERBAIK"
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
            LCID            =   1057
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
            LCID            =   1057
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
      Height          =   330
      Left            =   7560
      Top             =   2160
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HP\OneDrive\Skripsi\Program\db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HP\OneDrive\Skripsi\Program\db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TLBO_parameter_terbaik"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Inisialisasi Parameter LM :"
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   3375
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   27
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   26
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000010&
         Caption         =   "PROSES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Faktor Beta [1.10]"
         BeginProperty Font 
            Name            =   "Ink Free"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Batas Error (0,1)"
         BeginProperty Font 
            Name            =   "Ink Free"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Maksimal Iterasi [1,500]"
         BeginProperty Font 
            Name            =   "Ink Free"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Parameter LM (0,1]"
         BeginProperty Font 
            Name            =   "Ink Free"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Parameter Model : "
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2895
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   5760
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   5160
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   5160
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   5760
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   240
         Picture         =   "inisialisasi_lm.frx":6D8FB
         ScaleHeight     =   6.376
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   1.72
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "delta"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         Top             =   840
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H008080FF&
         DrawMode        =   4  'Mask Not Pen
         FillColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         DataField       =   "no"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Parameter ke :"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         DataField       =   "omega"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         DataField       =   "mmre"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         DataField       =   "teta"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         DataField       =   "epsilon"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         DataField       =   "miu"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cari Parameter :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Pilih Parameter :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Error"
         BeginProperty Font 
            Name            =   "Cambria Math"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   14
         Top             =   4080
         Width           =   975
      End
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      DataField       =   "maxit"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   8400
      TabIndex        =   31
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      DataField       =   "jumlahpelajar"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7560
      TabIndex        =   30
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PROSES AWAL IDENTIFIKASI"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If (Text11.Text = "") Then
    MsgBox "Silahkan isi dengan nomor parameter!", vbOKOnly, "Peringatan"
    Text11.SetFocus
Else
    Text1.SetFocus
    Adodc1.Recordset.Find "no=" + Text11.Text, , adSearchForward, 1
    If Not Adodc1.Recordset.EOF Then
        Label16.Caption = Adodc1.Recordset!no
        Label14.Caption = Adodc1.Recordset!mmre
    Else
       MsgBox "Maaf, Data tidak ditemukan!", vbOKOnly, "Peringatan"
    End If
End If
End Sub

Private Sub Command2_Click()
Dim iterasi As Integer, unitinput As Double, unithidden As Double, unitoutput As Double, data As Integer
Dim myu(3) As Double, jmse(3, 5000) As Double, minimalmse(3) As Double, msej(3) As Double
Dim selisih(3, 60) As Double, ermse(3) As Double, sumv As Double, sumw As Double
Dim slopez(3, 60, 3) As Double, slopey(3, 60, 3) As Double, dw(3, 60, 2, 1) As Double, dv(3, 60, 2, 2) As Double
Dim jacobi(3, 60, 9) As Double, transpose(3, 9, 60) As Double, hessian(3, 9, 9) As Double, iden(3, 18, 18) As Double
Dim hess(3, 9, 9) As Double, gradien(3, 9, 1) As Double, elemen(3, 18, 18) As Double, invers(3, 18, 18) As Double
Dim deltabobot(3, 9, 1) As Double, errorval(60, 5) As Double, jumlah1 As Double, nomor As Integer

data = 59

mu = Val(Text1.Text)
beta = Val(Text2.Text)
epoch = Val(Text3.Text)
err = Val(Text4.Text)
betta = beta

nofix = Label16.Caption
nfix = Label17.Caption
maxitfix = Label18.Caption

'parameter model fix
deltafix = Label9.Caption
omegafix = Label10.Caption
miufix = Label11.Caption
epsilonfix = Label12.Caption
tetafix = Label13.Caption
mmrefix = Label14.Caption

'input parameter
If (Text1.Text = "") Then
    MsgBox "Silahkan isi parameter LM!", vbOKOnly, "Parameter tidak tepat"
    GoTo awal
Else
If (mu <= 0 Or mu > 1) Then
    MsgBox "Parameter LM harus pada interval (0,1)!", vbOKOnly, "Parameter tidak tepat"
    GoTo awal
Else
If (Text2.Text = "") Then
    MsgBox "silahkan isi faktor beta!", vbOKOnly, "Parameter tidak tepat"
    GoTo awal
Else
If (beta < 1 Or beta > 10) Then
    MsgBox "Faktor beta harus bilangan asli pada interval [1,10]!", vbOKOnly, "Parameter tidak tepat"
    GoTo awal
Else
If (Text3.Text = "") Then
    MsgBox "Silahkan isi maksimal iterasi!", vbOKOnly, "Parameter tidak tepat"
    GoTo awal
Else
If (epoch < 1) Then
    MsgBox "maksimal iterasi harus bilangan asli!", vbOKOnly, "Parameter tidak tepat"
    GoTo awal
Else
If (Text4.Text = "") Then
    GoTo awal
Else
If (err <= 0 Or err >= 1) Then
    MsgBox "batas error harus pada interval (0,1)", vbOKOnly, "Parameter tidak tepat"
    GoTo awal
Else
    GoTo proses
End If
End If
End If
End If
End If
End If
End If
End If

proses:
'menampilkan inisialisasi parameter LM
Form4.Text19.Text = mu
Form4.Text20.Text = beta
Form4.Text21.Text = epoch
Form4.Text22.Text = err

'input data dari vb ke datagrid
Form2.Adodc1.Recordset.MoveFirst
For i = 0 To 95
    dataasli(i, 0) = Form2.Adodc1.Recordset.Fields("Bulan Ke-").Value
    dataasli(i, 1) = Form2.Adodc1.Recordset.Fields("S").Value
    dataasli(i, 2) = Form2.Adodc1.Recordset.Fields("I").Value
    dataasli(i, 3) = Form2.Adodc1.Recordset.Fields("R").Value
    Form2.Adodc1.Recordset.MoveNext
Next i
For i = 0 To 95
    dataasli(i, 4) = dataasli(i, 1) * dataasli(i, 2) / (dataasli(i, 1) + dataasli(i, 2) + dataasli(i, 3))
Next i

'Normalisasi Data
For i = 1 To 4
    maks(i) = dataasli(0, i)
    mini(i) = dataasli(0, i)
Next i
For i = 1 To 4
    For j = 0 To 95
        If (maks(i) < dataasli(j, i)) Then
            maks(i) = dataasli(j, i)
        End If
        If (mini(i) > dataasli(j, i)) Then
            mini(i) = dataasli(j, i)
        End If
    Next j
Next i
For i = 1 To 4
    For j = 0 To 95
        normalisasi(j, i) = Round((-1 + ((2 * (dataasli(j, i) - mini(i))) / (maks(i) - mini(i)))), 6)
    Next j
Next i

'menampilkan hasil Normalisasi
For j = 0 To data
    Form4.Adodc1.Recordset.AddNew
    Form4.Adodc1.Recordset.Fields("Bulan Ke-") = j + 1
    Form4.Adodc1.Recordset.Fields("S") = normalisasi(j, 1)
    Form4.Adodc1.Recordset.Fields("I") = normalisasi(j, 2)
    Form4.Adodc1.Recordset.Fields("R") = normalisasi(j, 3)
    Form4.Adodc1.Recordset.Fields("M") = normalisasi(j, 4)
    Form4.Adodc1.Recordset.Update
    Form4.DataGrid1.Refresh
Next j
        
'Pola Data Pelatihan
For i = 1 To 3
    For j = 0 To data
        poladata(i, j, 0) = j + 1
    Next j
    If (i = 1) Then
        For j = 0 To data
            poladata(i, j, 1) = normalisasi(j, 1) 's
            poladata(i, j, 2) = normalisasi(j, 4) 'm
            target(i, j) = normalisasi(j, 1) ' t=s
        Next j
    Else
    If (i = 2) Then
        For j = 0 To data
            poladata(i, j, 1) = normalisasi(j, 2) 'i
            poladata(i, j, 2) = normalisasi(j, 4) 'm
            target(i, j) = normalisasi(j, 2) ' t=i
        Next j
    Else
        For j = 0 To data
            poladata(i, j, 1) = normalisasi(j, 2) 'i
            poladata(i, j, 2) = normalisasi(j, 3) 'r
            target(i, j) = normalisasi(j, 3) 't=r
        Next j
    End If
    End If
Next i

'menampilkan pola data pelatihan
'populasi S
For j = 0 To data
    Form4.Adodc2.Recordset.AddNew
    Form4.Adodc2.Recordset.Fields("Pola") = j + 1
    Form4.Adodc2.Recordset.Fields("S") = poladata(1, j, 1)
    Form4.Adodc2.Recordset.Fields("M") = poladata(1, j, 2)
    Form4.Adodc2.Recordset.Fields("Target") = target(1, j)
    Form4.Adodc2.Recordset.Update
    Form4.DataGrid3.Refresh
Next j
'populasi I
For j = 0 To data
    Form4.Adodc3.Recordset.AddNew
    Form4.Adodc3.Recordset.Fields("Pola") = j + 1
    Form4.Adodc3.Recordset.Fields("I") = poladata(2, j, 1)
    Form4.Adodc3.Recordset.Fields("M") = poladata(2, j, 2)
    Form4.Adodc3.Recordset.Fields("Target") = target(2, j)
    Form4.Adodc3.Recordset.Update
    Form4.DataGrid4.Refresh
Next j
'populasi R
For j = 0 To data
    Form4.Adodc4.Recordset.AddNew
    Form4.Adodc4.Recordset.Fields("Pola") = j + 1
    Form4.Adodc4.Recordset.Fields("I") = poladata(3, j, 1)
    Form4.Adodc4.Recordset.Fields("R") = poladata(3, j, 2)
    Form4.Adodc4.Recordset.Fields("Target") = target(3, j)
    Form4.Adodc4.Recordset.Update
    Form4.DataGrid5.Refresh
Next j
        
'Bobot dan bias awal
For pop = 1 To 3
    If (pop = 1) Then
        'populasi S (dari input ke hidden)
        'bobot
        bbv(pop, 1, 1) = 1 - miufix
        bbv(pop, 1, 2) = 1 - miufix
        bbv(pop, 2, 1) = -omegafix
        bbv(pop, 2, 2) = -omegafix
        'bias
        bbv(pop, 0, 1) = deltafix
        bbv(pop, 0, 2) = deltafix
        'bobot dan bias dari hidden ke output
        For j = 0 To 2
            bbw(pop, j, 1) = randomantara(-1, 1)
        Next j
        'bbw(pop, 1, 1) = -0.702653
        'bbw(pop, 2, 1) = 0.856822
        'bbw(pop, 0, 1) = -0.536934

    ElseIf (pop = 2) Then
        'populasi I (dari input ke hidden)
        'bobot
        bbv(pop, 1, 1) = 1 - (epsilonfix + miufix + tetafix)
        bbv(pop, 1, 2) = 1 - (epsilonfix + miufix + tetafix)
        For j = 1 To 2
            If (bbv(pop, 1, j) < -1 Or bbv(pop, 1, j) > 1) Then
                bbv(pop, 1, j) = sigmoidbipolar(bbv(pop, 1, j))
            End If
        Next j
        bbv(pop, 2, 1) = omegafix
        bbv(pop, 2, 2) = omegafix
        'bias
        bbv(pop, 0, 1) = 0
        bbv(pop, 0, 2) = 0
        'bobot dan bias dari hidden ke output
        For j = 0 To 2
            bbw(pop, j, 1) = randomantara(-1, 1)
        Next j
        'bbw(pop, 1, 1) = -0.724015
        'bbw(pop, 2, 1) = 0.92091
        'bbw(pop, 0, 1) = -0.35276
    ElseIf (pop = 3) Then
        'populasi R (dari input ke hidden)
        'bobot
        bbv(pop, 1, 1) = epsilonfix
        bbv(pop, 1, 2) = epsilonfix
        bbv(pop, 2, 1) = 1 - miufix
        bbv(pop, 2, 2) = 1 - miufix
        'bias
        bbv(pop, 0, 1) = 0
        bbv(pop, 0, 2) = 0
        'bobot dan bias dari hidden ke output
        For j = 0 To 2
            bbw(pop, j, 1) = randomantara(-1, 1)
        Next j
        'bbw(pop, 1, 1) = -0.147223
        'bbw(pop, 2, 1) = -0.809466
        'bbw(pop, 0, 1) = -0.545024
    End If
Next pop

'menampilkan bobot dan bias awal
'populasi S
Form4.v11s.Text = Round(bbv(1, 1, 1), 6)
Form4.v12s.Text = Round(bbv(1, 1, 2), 6)
Form4.v21s.Text = Round(bbv(1, 2, 1), 6)
Form4.v22s.Text = Round(bbv(1, 2, 2), 6)
Form4.v01s.Text = Round(bbv(1, 0, 1), 6)
Form4.v02s.Text = Round(bbv(1, 0, 2), 6)
Form4.w11s.Text = Round(bbw(1, 1, 1), 6)
Form4.w21s.Text = Round(bbw(1, 2, 1), 6)
Form4.w01s.Text = Round(bbw(1, 0, 1), 6)
'populasi I
Form4.v11i.Text = Round(bbv(2, 1, 1), 6)
Form4.v12i.Text = Round(bbv(2, 1, 2), 6)
Form4.v21i.Text = Round(bbv(2, 2, 1), 6)
Form4.v22i.Text = Round(bbv(2, 2, 2), 6)
Form4.v01i.Text = Round(bbv(2, 0, 1), 6)
Form4.v02i.Text = Round(bbv(2, 0, 2), 6)
Form4.w11i.Text = Round(bbw(2, 1, 1), 6)
Form4.w21i.Text = Round(bbw(2, 2, 1), 6)
Form4.w01i.Text = Round(bbw(2, 0, 1), 6)
'populasi R
Form4.v11r.Text = Round(bbv(3, 1, 1), 6)
Form4.v12r.Text = Round(bbv(3, 1, 2), 6)
Form4.v21r.Text = Round(bbv(3, 2, 1), 6)
Form4.v22r.Text = Round(bbv(3, 2, 2), 6)
Form4.v01r.Text = Round(bbv(3, 0, 1), 6)
Form4.v02r.Text = Round(bbv(3, 0, 2), 6)
Form4.w11r.Text = Round(bbw(3, 1, 1), 6)
Form4.w21r.Text = Round(bbw(3, 2, 1), 6)
Form4.w01r.Text = Round(bbw(3, 0, 1), 6)

unitinput = 2
unithidden = 2
unitoutput = 1
iterasi = 0

'Proses LM untuk identifikasi
Do
iterasi = iterasi + 1
If iterasi = 1 Then
    For pop = 1 To 3
        myu(pop) = mu
    Next pop
End If
For pop = 1 To 3
    m = 0
    Do
    'proses feedforward
    ermse(pop) = 0
    For pola = 0 To data
        For i = 1 To unithidden
            sumv = 0
            For j = 1 To unitinput
                sumv = sumv + poladata(pop, pola, j) * bbv(pop, j, i)
            Next j
            zin(pop, pola, i) = bbv(pop, 0, i) + sumv
            z(pop, pola, i) = sigmoidbipolar(zin(pop, pola, i))
            slopez(pop, pola, i) = slope(z(pop, pola, i))
        Next i
        For i = 1 To unitoutput
            sumw = 0
            For j = 1 To unithidden
                sumw = sumw + z(pop, pola, j) * bbw(pop, j, i)
            Next j
            yin(pop, pola, i) = bbw(pop, 0, 1) + sumw
            Y(pop, pola, i) = sigmoidbipolar(yin(pop, pola, i))
            slopey(pop, pola, i) = slope(Y(pop, pola, i))
            outputs(pop, pola) = Y(pop, pola, i)
            'selisih (t-y)
            selisih(pop, pola) = target(pop, pola) - outputs(pop, pola)
            ermse(pop) = ermse(pop) + (selisih(pop, pola)) ^ 2
        Next i
    'proses backward
        For i = 1 To unithidden
            For j = 1 To unitoutput
                dw(pop, pola, i, j) = slopey(pop, pola, j) * z(pop, pola, i)
                dw(pop, pola, 0, j) = slopey(pop, pola, j)
            Next j
        Next i
        For i = 1 To unitinput
            For j = 1 To unithidden
                dv(pop, pola, i, j) = slopey(pop, pola, 1) * slopez(pop, pola, j) * poladata(pop, pola, i)
                dv(pop, pola, 0, j) = slopey(pop, pola, 1) * slopez(pop, pola, j)
            Next j
        Next i
        'membentuk matriks jacobi
        jacobi(pop, pola, 1) = -dv(pop, pola, 1, 1)
        jacobi(pop, pola, 2) = -dv(pop, pola, 1, 2)
        jacobi(pop, pola, 3) = -dv(pop, pola, 2, 1)
        jacobi(pop, pola, 4) = -dv(pop, pola, 2, 2)
        jacobi(pop, pola, 5) = -dv(pop, pola, 0, 1)
        jacobi(pop, pola, 6) = -dv(pop, pola, 0, 2)
        jacobi(pop, pola, 7) = -dw(pop, pola, 1, 1)
        jacobi(pop, pola, 8) = -dw(pop, pola, 2, 1)
        jacobi(pop, pola, 9) = -dw(pop, pola, 0, 1)
    Next pola
    
    'menghitung MSE
    jmse(pop, iterasi) = ermse(pop) / (data + 1) 'nilai mse populasi ke-pop iterasi ke-iterasi
    
    'Proses pembaruan bobot dan bias
    If iterasi > 1 Then
        'transpose jacobi
        For pola = 0 To data
            For j = 1 To 9
                transpose(pop, j, pola) = jacobi(pop, pola, j)
            Next j
        Next pola
        'perkalian matriks jacobi (J^T*J)
        For i = 1 To 9
            For j = 1 To 9
                hess(pop, i, j) = 0
                For pola = 0 To data
                    hess(pop, i, j) = hess(pop, i, j) + transpose(pop, i, pola) * jacobi(pop, pola, j)
                Next pola
            Next j
        Next i
        'matriks identitas 9x9
        For i = 1 To 9
            For j = 1 To 9
                If (i = j) Then
                    iden(pop, i, j) = 1
                Else
                    iden(pop, i, j) = 0
                End If
            Next j
        Next i
        'menghitung (H+mu*I)
        For i = 1 To 9
            For j = 1 To 9
                hessian(pop, i, j) = hess(pop, i, j) + (myu(pop) * iden(pop, i, j))
            Next j
        Next i
        'menghitung invers hessian (A|I)
        For i = 1 To 9
            For j = 1 To 18
                If j <= 9 Then
                    elemen(pop, i, j) = hessian(pop, i, j)
                Else
                    elemen(pop, i, j) = iden(pop, i, j - 9)
                End If
            Next j
        Next i
        For i = 1 To 9
            For j = 1 To 18
                If (i <> j) Then
                    elemen(pop, i, j) = elemen(pop, i, j) / elemen(pop, i, i)
                End If
                Next j
                For j = 1 To 18
                    If (i = j) Then
                        elemen(pop, i, j) = 1
                    End If
            Next j
            For l = 1 To 9
                If (i <> l) Then
                    For j = i + 1 To 18
                        elemen(pop, l, j) = elemen(pop, l, j) - (elemen(pop, i, j) * elemen(pop, l, i))
                    Next j
                End If
            Next l
            For k = 1 To 9
                If (i <> k) Then
                    elemen(pop, k, i) = 0
                End If
            Next k
        Next i
        'hasil invers
        For i = 1 To 9
            For j = 1 To 9
                invers(pop, i, j) = elemen(pop, i, j + 9)
            Next j
        Next i
        'menghitung gradien g
        For i = 1 To 9
            For j = 1 To 1
                gradien(pop, i, j) = 0
                For pola = 0 To data
                    gradien(pop, i, j) = gradien(pop, i, j) + (transpose(pop, i, pola) * selisih(pop, pola))
                Next pola
            Next j
        Next i
        'menghitung delta bobot dan bias
        For i = 1 To 9
            For j = 1 To 1
                deltabobot(pop, i, j) = 0
                For k = 1 To 9
                    deltabobot(pop, i, j) = deltabobot(pop, i, j) + (invers(pop, i, k) * gradien(pop, k, j))
                Next k
            Next j
        Next i
        For i = 1 To 9
            If deltabobot(pop, i, 1) < -1 Or deltabobot(pop, i, 1) > 1 Then
                deltabobot(pop, i, 1) = sigmoidbipolar(deltabobot(pop, i, 1))
            End If
        Next i
        
       'update bobot dan bias
       bbv(pop, 1, 1) = bbv(pop, 1, 1) - deltabobot(pop, 1, 1)
       bbv(pop, 1, 2) = bbv(pop, 1, 2) - deltabobot(pop, 2, 1)
       bbv(pop, 2, 1) = bbv(pop, 2, 1) - deltabobot(pop, 3, 1)
       bbv(pop, 2, 2) = bbv(pop, 2, 2) - deltabobot(pop, 4, 1)
       bbv(pop, 0, 1) = bbv(pop, 0, 1) - deltabobot(pop, 5, 1)
       bbv(pop, 0, 2) = bbv(pop, 0, 2) - deltabobot(pop, 6, 1)
       bbw(pop, 1, 1) = bbw(pop, 1, 1) - deltabobot(pop, 7, 1)
       bbw(pop, 2, 1) = bbw(pop, 2, 1) - deltabobot(pop, 8, 1)
       bbw(pop, 0, 1) = bbw(pop, 0, 1) - deltabobot(pop, 9, 1)
        'For i = 0 To 2
        '    For j = 1 To 2
        '        If bbv(pop, i, j) <= -1 Or bbv(pop, i, j) >= 1 Then
        '         bbv(pop, i, j) = sigmoidbipolar(bbv(pop, i, j))
        '      End If
        '    Next j
        '      For j = 1 To 1
        '        If bbw(pop, i, j) <= -1 Or bbw(pop, i, j) >= 1 Then
        '            bbw(pop, i, j) = sigmoidbipolar(bbw(pop, i, j))
        '        End If
        '    Next j
        'Next i
    End If 'iterasi>1
    If (msej(pop) < jmse(pop, iterasi)) Then 'perbandingan antara mse E(k) < mse E(k+1)
        myu(pop) = myu(pop) * beta
    End If
    m = m + 1
    Loop While msej(pop) < jmse(pop, iterasi) And m <= 5
    If (iterasi > 0) Then
        myu(pop) = myu(pop) / beta
    End If
Next pop

'minimal MSE
If (iterasi = 1) Then
    For pop = 1 To 3
        minimalmse(pop) = jmse(pop, iterasi) ' jmse adalah nilai mse populasi ke-pop tiap iterasi
    Next pop
Else
    For pop = 1 To 3
        If jmse(pop, iterasi) < minimalmse(pop) Then
            For i = 0 To 2
                For j = 1 To 2
                    If bbv(pop, i, j) < -1 Or bbv(pop, i, j) > 1 Then
                     bbv(pop, i, j) = sigmoidbipolar(bbv(pop, i, j))
                  End If
                Next j
                  For j = 1 To 1
                    If bbw(pop, i, j) < -1 Or bbw(pop, i, j) > 1 Then
                        bbw(pop, i, j) = sigmoidbipolar(bbw(pop, i, j))
                    End If
                Next j
            Next i
            'bobot dan bias optimal
            minimalmse(pop) = jmse(pop, iterasi)
            bbvn(pop, 1, 1) = bbv(pop, 1, 1)
            bbvn(pop, 1, 2) = bbv(pop, 1, 2)
            bbvn(pop, 2, 1) = bbv(pop, 2, 1)
            bbvn(pop, 2, 2) = bbv(pop, 2, 2)
            bbvn(pop, 0, 1) = bbv(pop, 0, 1)
            bbvn(pop, 0, 2) = bbv(pop, 0, 2)
            bbwn(pop, 1, 1) = bbw(pop, 1, 1)
            bbwn(pop, 2, 1) = bbw(pop, 2, 1)
            bbwn(pop, 0, 1) = bbw(pop, 0, 1)
            'keluaran jaringan
            For j = 0 To data
                outputfix(pop, j) = outputs(pop, j)
            Next j
        End If
    Next pop
End If

For pop = 1 To 3
    msej(pop) = jmse(pop, iterasi) 'msej adalah nilai mse E(k)
Next pop

'menghitung MSE akhir
mse = (minimalmse(1) + minimalmse(2) + minimalmse(3)) / 3
Form4.Adodc5.Recordset.AddNew
Form4.Adodc5.Recordset.Fields("Iterasi") = iterasi
Form4.Adodc5.Recordset.Fields("MSE S") = Round(minimalmse(1), 6)
Form4.Adodc5.Recordset.Fields("MSE I") = Round(minimalmse(2), 6)
Form4.Adodc5.Recordset.Fields("MSE R") = Round(minimalmse(3), 6)
Form4.Adodc5.Recordset.Fields("MSE Akhir") = Round(mse, 6)
Form4.Adodc5.Recordset.Update
Form4.DataGrid6.Refresh

Loop While iterasi < epoch And mse > err

epochakhir = iterasi
mseakhir = Round(mse, 6)
Form5.Label3.Caption = Round(mse, 6)

Form4.v11bs.Text = Round(bbvn(1, 1, 1), 6)
Form4.v12bs.Text = Round(bbvn(1, 1, 2), 6)
Form4.v21bs.Text = Round(bbvn(1, 2, 1), 6)
Form4.v22bs.Text = Round(bbvn(1, 2, 2), 6)
Form4.v01bs.Text = Round(bbvn(1, 0, 1), 6)
Form4.v02bs.Text = Round(bbvn(1, 0, 2), 6)
Form4.w11bs.Text = Round(bbwn(1, 1, 1), 6)
Form4.w21bs.Text = Round(bbwn(1, 2, 1), 6)
Form4.w01bs.Text = Round(bbwn(1, 0, 1), 6)

Form4.v11bi.Text = Round(bbvn(2, 1, 1), 6)
Form4.v12bi.Text = Round(bbvn(2, 1, 2), 6)
Form4.v21bi.Text = Round(bbvn(2, 2, 1), 6)
Form4.v22bi.Text = Round(bbvn(2, 2, 2), 6)
Form4.v01bi.Text = Round(bbvn(2, 0, 1), 6)
Form4.v02bi.Text = Round(bbvn(2, 0, 2), 6)
Form4.w11bi.Text = Round(bbwn(2, 1, 1), 6)
Form4.w21bi.Text = Round(bbwn(2, 2, 1), 6)
Form4.w01bi.Text = Round(bbwn(2, 0, 1), 6)

Form4.v11br.Text = Round(bbvn(3, 1, 1), 6)
Form4.v12br.Text = Round(bbvn(3, 1, 2), 6)
Form4.v21br.Text = Round(bbvn(3, 2, 1), 6)
Form4.v22br.Text = Round(bbvn(3, 2, 2), 6)
Form4.v01br.Text = Round(bbvn(3, 0, 1), 6)
Form4.v02br.Text = Round(bbvn(3, 0, 2), 6)
Form4.w11br.Text = Round(bbwn(3, 1, 1), 6)
Form4.w21br.Text = Round(bbwn(3, 2, 1), 6)
Form4.w01br.Text = Round(bbwn(3, 0, 1), 6)

'denormalisasi data
For pop = 1 To 3
    For pola = 0 To data
        denormalisasi(pola, pop) = ((outputfix(pop, pola) + 1) * (maks(pop) - mini(pop)) / 2) + mini(pop)
    Next pola
Next pop

'MMRE Akhir untuk identifikasi
For pola = 0 To data
    jumlah1 = 0
    For pop = 1 To 3
        errorval(pola, pop) = mre(dataasli(pola, pop), denormalisasi(pola, pop))
        jumlah1 = jumlah1 + errorval(pola, pop)
    Next pop
    errorval(pola, 4) = jumlah1 / 3
Next pola
jumlah1 = 0
For pola = 0 To data
    jumlah1 = jumlah1 + errorval(pola, 4)
Next pola
mmreakhir = Round(jumlah1 / (data + 1), 6)

'menampilkan nilai MMRE
Form6.Label6.Caption = mmreakhir

For i = 0 To data
    Form6.Adodc1.Recordset.AddNew
    Form6.Adodc1.Recordset.Fields("Bulan Ke-") = 1 + i
    Form6.Adodc1.Recordset.Fields("S") = Round(denormalisasi(i, 1), 0)
    Form6.Adodc1.Recordset.Fields("I") = Round(denormalisasi(i, 2), 0)
    Form6.Adodc1.Recordset.Fields("R") = Round(denormalisasi(i, 3), 0)
    Form6.Adodc1.Recordset.Update
    Form6.DataGrid1.Refresh
Next i

'Save to database
Form4.Adodc6.Recordset.MoveLast
no = Form4.Adodc6.Recordset.Fields("no").Value
no = no + 1
Form4.Adodc6.Recordset.AddNew
Form4.Adodc6.Recordset.Fields("no") = no
Form4.Adodc6.Recordset.Fields("parameter ke") = nofix
Form4.Adodc6.Recordset.Fields("jumlah pelajar") = nfix
Form4.Adodc6.Recordset.Fields("maxit") = maxitfix
Form4.Adodc6.Recordset.Fields("delta") = deltafix
Form4.Adodc6.Recordset.Fields("omega") = omegafix
Form4.Adodc6.Recordset.Fields("miu") = miufix
Form4.Adodc6.Recordset.Fields("epsilon") = epsilonfix
Form4.Adodc6.Recordset.Fields("teta") = deltafix
Form4.Adodc6.Recordset.Fields("error") = mmrefix
Form4.Adodc6.Recordset.Fields("parameter lm") = mu
Form4.Adodc6.Recordset.Fields("faktor beta") = beta
Form4.Adodc6.Recordset.Fields("maks epoch") = epoch
Form4.Adodc6.Recordset.Fields("epoch") = epochakhir
Form4.Adodc6.Recordset.Fields("batas error") = err
Form4.Adodc6.Recordset.Fields("v11s") = Round(bbvn(1, 1, 1), 6)
Form4.Adodc6.Recordset.Fields("v12s") = Round(bbvn(1, 1, 2), 6)
Form4.Adodc6.Recordset.Fields("v21s") = Round(bbvn(1, 2, 1), 6)
Form4.Adodc6.Recordset.Fields("v22s") = Round(bbvn(1, 2, 2), 6)
Form4.Adodc6.Recordset.Fields("v01s") = Round(bbvn(1, 0, 1), 6)
Form4.Adodc6.Recordset.Fields("v02s") = Round(bbvn(1, 0, 2), 6)
Form4.Adodc6.Recordset.Fields("w11s") = Round(bbwn(1, 1, 1), 6)
Form4.Adodc6.Recordset.Fields("w21s") = Round(bbwn(1, 2, 1), 6)
Form4.Adodc6.Recordset.Fields("w01s") = Round(bbwn(1, 0, 1), 6)
Form4.Adodc6.Recordset.Fields("v11i") = Round(bbvn(2, 1, 1), 6)
Form4.Adodc6.Recordset.Fields("v12i") = Round(bbvn(2, 1, 2), 6)
Form4.Adodc6.Recordset.Fields("v21i") = Round(bbvn(2, 2, 1), 6)
Form4.Adodc6.Recordset.Fields("v22i") = Round(bbvn(2, 2, 2), 6)
Form4.Adodc6.Recordset.Fields("v01i") = Round(bbvn(2, 0, 1), 6)
Form4.Adodc6.Recordset.Fields("v02i") = Round(bbvn(2, 0, 2), 6)
Form4.Adodc6.Recordset.Fields("w11i") = Round(bbwn(2, 1, 1), 6)
Form4.Adodc6.Recordset.Fields("w21i") = Round(bbwn(2, 2, 1), 6)
Form4.Adodc6.Recordset.Fields("w01i") = Round(bbwn(2, 0, 1), 6)
Form4.Adodc6.Recordset.Fields("v11r") = Round(bbvn(3, 1, 1), 6)
Form4.Adodc6.Recordset.Fields("v12r") = Round(bbvn(3, 1, 2), 6)
Form4.Adodc6.Recordset.Fields("v21r") = Round(bbvn(3, 2, 1), 6)
Form4.Adodc6.Recordset.Fields("v22r") = Round(bbvn(3, 2, 2), 6)
Form4.Adodc6.Recordset.Fields("v01r") = Round(bbvn(3, 0, 1), 6)
Form4.Adodc6.Recordset.Fields("v02r") = Round(bbvn(3, 0, 2), 6)
Form4.Adodc6.Recordset.Fields("w11r") = Round(bbwn(3, 1, 1), 6)
Form4.Adodc6.Recordset.Fields("w21r") = Round(bbwn(3, 2, 1), 6)
Form4.Adodc6.Recordset.Fields("w01r") = Round(bbwn(3, 0, 1), 6)
Form4.Adodc6.Recordset.Fields("mse") = mseakhir
Form4.Adodc6.Recordset.Fields("MMRE Validasi") = mmreakhir
Form4.Adodc6.Recordset.Update
Form4.DataGrid7.Refresh

For i = 0 To data
Form4.Adodc8.Recordset.AddNew
Form4.Adodc8.Recordset.Fields("no") = no
Form4.Adodc8.Recordset.Fields("parameter") = nofix
Form4.Adodc8.Recordset.Fields("Bulan Ke-") = 1 + i
Form4.Adodc8.Recordset.Fields("S") = Round(denormalisasi(i, 1), 0)
Form4.Adodc8.Recordset.Fields("I") = Round(denormalisasi(i, 2), 0)
Form4.Adodc8.Recordset.Fields("R") = Round(denormalisasi(i, 3), 0)
Form4.Adodc8.Recordset.Update
Form4.DataGrid9.Refresh
Next i

Form4.v11s.Visible = False
Form4.v12s.Visible = False
Form4.v21s.Visible = False
Form4.v22s.Visible = False
Form4.v01s.Visible = False
Form4.v02s.Visible = False
Form4.w11s.Visible = False
Form4.w21s.Visible = False
Form4.w01s.Visible = False
Form4.v11i.Visible = False
Form4.v12i.Visible = False
Form4.v21i.Visible = False
Form4.v22i.Visible = False
Form4.v01i.Visible = False
Form4.v02i.Visible = False
Form4.w11i.Visible = False
Form4.w21i.Visible = False
Form4.w01i.Visible = False
Form4.v11r.Visible = False
Form4.v12r.Visible = False
Form4.v21r.Visible = False
Form4.v22r.Visible = False
Form4.v01r.Visible = False
Form4.v02r.Visible = False
Form4.w11r.Visible = False
Form4.w21r.Visible = False
Form4.w01r.Visible = False

Form4.v11bs.Visible = False
Form4.v12bs.Visible = False
Form4.v21bs.Visible = False
Form4.v22bs.Visible = False
Form4.v01bs.Visible = False
Form4.v02bs.Visible = False
Form4.w11bs.Visible = False
Form4.w21bs.Visible = False
Form4.w01bs.Visible = False
Form4.v11bi.Visible = False
Form4.v12bi.Visible = False
Form4.v21bi.Visible = False
Form4.v22bi.Visible = False
Form4.v01bi.Visible = False
Form4.v02bi.Visible = False
Form4.w11bi.Visible = False
Form4.w21bi.Visible = False
Form4.w01bi.Visible = False
Form4.v11br.Visible = False
Form4.v12br.Visible = False
Form4.v21br.Visible = False
Form4.v22br.Visible = False
Form4.v01br.Visible = False
Form4.v02br.Visible = False
Form4.w11br.Visible = False
Form4.w21br.Visible = False
Form4.w01br.Visible = False

Form4.DataGrid3.Visible = False
Form4.DataGrid4.Visible = False
Form4.DataGrid5.Visible = False

Form4.Show
Form3.Hide

awal:
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

End Sub

Private Sub Command3_Click()
    If Not (Adodc1.Recordset.EOF) Then
        Adodc1.Recordset.MoveNext
        If Label16.Caption = "" Then
            Adodc1.Recordset.MovePrevious
        End If
    End If
End Sub

Private Sub Command4_Click()
    If Not (Adodc1.Recordset.BOF) Then
        Adodc1.Recordset.MovePrevious
        If Label16.Caption = "" Or Label16.Caption = 0 Or Label16.Caption = "0" Then
            Adodc1.Recordset.MoveNext
        End If
    End If
End Sub

Private Sub Command7_Click()
Form2.Show
Form2.Text8.SetFocus
Form2.Command2.Enabled = False
Form3.Hide
End Sub
