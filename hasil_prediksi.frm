VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "hasil prediksi"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10635
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "hasil_prediksi.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "KEMBALI"
      Height          =   375
      Left            =   9000
      TabIndex        =   13
      Top             =   5880
      Width           =   1335
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4095
      Left            =   3840
      OleObjectBlob   =   "hasil_prediksi.frx":288D3
      TabIndex        =   9
      Top             =   1440
      Width           =   6615
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "hasil_prediksi.frx":2AC27
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   4680
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "hasil_prediksi.frx":2AC3C
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2990
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Populasi : "
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3375
      Begin VB.OptionButton Option3 
         Caption         =   "R"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "I"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "S"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   7920
      TabIndex        =   14
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   615
      Left            =   5880
      TabIndex        =   12
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Error"
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label Label3 
      Caption         =   "Data Hasil Prediksi"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Data Uji"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HASIL PREDIKSI"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   2865
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form8.Show
Form9.Hide
End Sub

Private Sub Option1_Click()
Dim xmin As Double, xmax As Double, xstep As Single, num_x As Integer, X As Double
Dim jumlah1 As Double
'grafik
xmin = 0
xmax = 36
xstep = 1
num_x = (xmax - xmin) / xstep
ReDim Values(1 To num_x, 1)

'menghitung nilai data
X = xmin
Values(1, 0) = dataaslipre(1, 1)
Values(1, 1) = denormalisasipre(1, 1)
For i = 2 To num_x
    Values(i, 0) = dataaslipre(i, 1)
    Values(i, 1) = denormalisasipre(i, 1)
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
Form6.MSChart1.Plot.SeriesCollection(2).LegendText = "Hasil Prediksi"

'MMRE Populasi S
jumlah1 = 0
For i = 1 To 36
    jumlah1 = jumlah1 + Abs(dataaslipre(i, 1) - denormalisasipre(i, 1)) / dataaslipre(i, 1)
Next i
mmres = jumlah1 / 36
'Label8.Caption = "MMRE POPULASI S = "
Label7.Caption = Round(mmres, 6)
Label4.Caption = "POPULASI S(SUCEPTIBLE)"

Option2.Value = False
Option3.Value = False

End Sub
