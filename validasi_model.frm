VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Hasil Validasi Model"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12105
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "validasi_model.frx":0000
   ScaleHeight     =   6375
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4455
      Left            =   3960
      OleObjectBlob   =   "validasi_model.frx":207B2
      TabIndex        =   13
      Top             =   1080
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
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
      Left            =   10680
      TabIndex        =   10
      Top             =   5880
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   6720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      RecordSource    =   "Denormalisasi"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Populasi "
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3735
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "validasi_model.frx":22C4E
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
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
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   5520
      Width           =   6015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   5520
      Width           =   2505
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "ERROR :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   600
      Width           =   8055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Hasil Denormalisasi Data"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HASIL VALIDASI MODEL"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   4245
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Show
Form6.Hide
End Sub

Private Sub Option1_Click()
Dim xmin As Double, xmax As Double, xstep As Single, num_x As Integer, X As Double
Dim jumlah1 As Double

'grafik
xmin = 0
xmax = 60
xstep = 1
num_x = (xmax - xmin) / xstep
ReDim Values(1 To num_x, 1)

'menghitung nilai data
X = xmin
Values(1, 0) = dataasli(0, 1)
Values(1, 1) = denormalisasi(0, 1)
For i = 2 To num_x
    Values(i, 0) = dataasli(i - 1, 1)
    Values(i, 1) = denormalisasi(i - 1, 1)
    X = X + xstep
Next i
Form6.MSChart1.RowCount = 2
Form6.MSChart1.ColumnCount = num_x
Form6.MSChart1.ChartData = Values
With Form6.MSChart1.Legend
    .Location.Visible = True
    .Location.LocationType = VtChLocationTypeBottomRight
    .TextLayout.HorzAlignment = VtHorizontalAlignmentCenter
End With
Form6.MSChart1.Plot.SeriesCollection(1).LegendText = "Data Asli"
Form6.MSChart1.Plot.SeriesCollection(2).LegendText = "Keluaran Jaringan"

'MMRE Populasi S
jumlah1 = 0
For i = 0 To 59
    jumlah1 = jumlah1 + Abs(dataasli(i, 1) - denormalisasi(i, 1)) / dataasli(i, 1)
Next i
mmres = jumlah1 / 60
Label3.Caption = "MMRE POPULASI S =  "
Label4.Caption = Round(mmres, 6)
Label5.Caption = "POPULASI S(SUSCEPTIBLE)"

Option2.Value = False
Option3.Value = False
End Sub

Private Sub Option2_Click()
Dim xmin As Double, xmax As Double, xstep As Single, num_x As Integer, X As Double
Dim jumlah1 As Double

'grafik
xmin = 0
xmax = 60
xstep = 1
num_x = (xmax - xmin) / xstep
ReDim Values(1 To num_x, 1)

'menghitung nilai data
X = xmin
Values(1, 0) = dataasli(0, 2)
Values(1, 1) = denormalisasi(0, 2)
For i = 2 To num_x
    Values(i, 0) = dataasli(i - 1, 2)
    Values(i, 1) = denormalisasi(i - 1, 2)
    X = X + xstep
Next i
Form6.MSChart1.RowCount = 2
Form6.MSChart1.ColumnCount = num_x
Form6.MSChart1.ChartData = Values
With Form6.MSChart1.Legend
    .Location.Visible = True
    .Location.LocationType = VtChLocationTypeBottomRight
    .TextLayout.HorzAlignment = VtHorizontalAlignmentCenter
End With
Form6.MSChart1.Plot.SeriesCollection(1).LegendText = "Data Asli"
Form6.MSChart1.Plot.SeriesCollection(2).LegendText = "Keluaran Jaringan"

'MMRE Populasi I
jumlah1 = 0
For i = 0 To 59
    jumlah1 = jumlah1 + Abs(dataasli(i, 2) - denormalisasi(i, 2)) / dataasli(i, 2)
Next i
mmrei = jumlah1 / 60
Label3.Caption = "MMRE POPULASI I =  "
Label4.Caption = Round(mmrei, 6)
Label5.Caption = "POPULASI I (INFECTED)"

Option1.Value = False
Option3.Value = False
End Sub

Private Sub Option3_Click()
Dim xmin As Double, xmax As Double, xstep As Single, num_x As Integer, X As Double
Dim jumlah1 As Double
'grafik
xmin = 0
xmax = 60
xstep = 1
num_x = (xmax - xmin) / xstep
ReDim Values(1 To num_x, 1)

'menghitung nilai data
X = xmin
Values(1, 0) = dataasli(0, 3)
Values(1, 1) = denormalisasi(0, 3)
For i = 2 To num_x
    Values(i, 0) = dataasli(i - 1, 3)
    Values(i, 1) = denormalisasi(i - 1, 3)
    X = X + xstep
Next i
Form6.MSChart1.RowCount = 2
Form6.MSChart1.ColumnCount = num_x
Form6.MSChart1.ChartData = Values
With Form6.MSChart1.Legend
    .Location.Visible = True
    .Location.LocationType = VtChLocationTypeBottomRight
    .TextLayout.HorzAlignment = VtHorizontalAlignmentCenter
End With
Form6.MSChart1.Plot.SeriesCollection(1).LegendText = "Data Asli"
Form6.MSChart1.Plot.SeriesCollection(2).LegendText = "Keluaran Jaringan"

'MMRE Populasi R
jumlah1 = 0
For i = 0 To 59
    jumlah1 = jumlah1 + Abs(dataasli(i, 3) - denormalisasi(i, 3)) / dataasli(i, 3)
Next i
mmrer = jumlah1 / 60
Label3.Caption = "MMRE POPULASI R =  "
Label4.Caption = Round(mmrer, 6)
Label5.Caption = "POPULASI R(RECOVERY)"

Option1.Value = False
Option2.Value = False
End Sub
