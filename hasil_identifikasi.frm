VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hasil Identifikasi Model dengan JST-LM"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12735
   FillColor       =   &H80000017&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "hasil_identifikasi.frx":0000
   ScaleHeight     =   6135
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid9 
      Bindings        =   "hasil_identifikasi.frx":124BE
      Height          =   135
      Left            =   4080
      TabIndex        =   104
      Top             =   7680
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   238
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   375
      Left            =   2040
      Top             =   8160
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
      RecordSource    =   "Denormalisasi_Identifikasi"
      Caption         =   "Adodc8"
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
   Begin MSDataGridLib.DataGrid DataGrid8 
      Bindings        =   "hasil_identifikasi.frx":124D3
      Height          =   135
      Left            =   4080
      TabIndex        =   103
      Top             =   7560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   238
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   375
      Left            =   2040
      Top             =   7800
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
      RecordSource    =   "LM_Hasil_Prediksi"
      Caption         =   "Adodc7"
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
   Begin MSDataGridLib.DataGrid DataGrid7 
      Bindings        =   "hasil_identifikasi.frx":124E8
      Height          =   135
      Left            =   4080
      TabIndex        =   102
      Top             =   7440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   238
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   2040
      Top             =   7440
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "LM_Hasil"
      Caption         =   "Adodc6"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   120
      Top             =   8880
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "LM_MSE"
      Caption         =   "Adodc5"
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
   Begin MSDataGridLib.DataGrid DataGrid5 
      Bindings        =   "hasil_identifikasi.frx":124FD
      Height          =   2295
      Left            =   3240
      TabIndex        =   55
      Top             =   3720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Populasi R (Recovery)"
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
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "hasil_identifikasi.frx":12512
      Height          =   2295
      Left            =   3240
      TabIndex        =   54
      Top             =   3720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Populasi I (Infected)"
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "hasil_identifikasi.frx":12527
      Height          =   2295
      Left            =   3240
      TabIndex        =   1
      Top             =   3720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Populasi S (Suceptible)"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   120
      Top             =   8520
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "LM_Poladata_R"
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   120
      Top             =   8160
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "LM_Poladata_I"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   120
      Top             =   7800
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "LM_Poladata_S"
      Caption         =   "Adodc2"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "hasil_identifikasi.frx":1253C
      Height          =   2175
      Left            =   3240
      TabIndex        =   52
      Top             =   1080
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   7440
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "Normalisasi"
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
   Begin VB.CommandButton Command4 
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
      Left            =   1200
      TabIndex        =   28
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PREDIKSI"
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
      Left            =   840
      TabIndex        =   27
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "VALIDASI"
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
      Left            =   840
      TabIndex        =   26
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GRAFIK MSE"
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
      Left            =   840
      TabIndex        =   25
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1080
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bobot dan Bias Akhir : "
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   7440
      TabIndex        =   4
      Top             =   3360
      Width           =   5175
      Begin VB.TextBox w01br 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox w21br 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox w11br 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v02br 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox v22br 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox v12br 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v01br 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox v21br 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox v11br 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox w01bi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox w21bi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox w11bi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v02bi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox v22bi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox v12bi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v01bi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox v21bi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox v11bi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox w01bs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox w21bs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox w11bs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v02bs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox v22bs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox v12bs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v01bs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox v21bs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox v11bs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   600
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   1620
         Left            =   3360
         Picture         =   "hasil_identifikasi.frx":12551
         Top             =   600
         Width           =   660
      End
      Begin VB.Image Image5 
         Height          =   1590
         Left            =   1800
         Picture         =   "hasil_identifikasi.frx":12D08
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Image4 
         Height          =   1650
         Left            =   120
         Picture         =   "hasil_identifikasi.frx":13465
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bobot dan Bias Awal :"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   7440
      TabIndex        =   3
      Top             =   600
      Width           =   5175
      Begin VB.TextBox v11r 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v21r 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox v11i 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v21i 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox v11s 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v21s 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v01r 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox v01i 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox v01s 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox v12r 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v12i 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v12s 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox v22r 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox v02r 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox v02i 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox v22i 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox v22s 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox v02s 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox w11r 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox w11i 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox w11s 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox w21r 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox w21i 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox w21s 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox w01r 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox w01i 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox w01s 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   1800
         Width           =   855
      End
      Begin VB.Image Image3 
         Height          =   1620
         Left            =   3360
         Picture         =   "hasil_identifikasi.frx":13C17
         Top             =   600
         Width           =   660
      End
      Begin VB.Image Image2 
         Height          =   1590
         Left            =   1800
         Picture         =   "hasil_identifikasi.frx":143CE
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   1650
         Left            =   120
         Picture         =   "hasil_identifikasi.frx":14B2B
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "Populasi"
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   30
      Top             =   3120
      Width           =   2895
      Begin VB.OptionButton Option3 
         BackColor       =   &H00400000&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Cambria Math"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2160
         TabIndex        =   33
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400000&
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Cambria Math"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   32
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00400000&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Cambria Math"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2295
      Left            =   3240
      TabIndex        =   23
      Top             =   3720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4048
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
   Begin MSDataGridLib.DataGrid DataGrid6 
      Bindings        =   "hasil_identifikasi.frx":152DD
      Height          =   615
      Left            =   4680
      TabIndex        =   65
      Top             =   7680
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "DATA PENYEBARAN PENYAKIT DBD"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   51
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Faktor Beta"
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "MSE UNTUK SETIAP ITERASI"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   24
      Top             =   7440
      Width           =   3495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Inputan Parameter LM"
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Batas Error"
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maksimal Iterasi"
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Parameter LM "
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HASIL IDENTIFIKASI MODEL"
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
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Setelah di Normalisasi ke Interval [-1,1]"
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3525
      TabIndex        =   2
      Top             =   840
      Width           =   3525
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "POLA DATA PELATIHAN"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   3360
      Width           =   4095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim xmin As Double, xmax As Double, xstep As Single, num_x As Integer, X As Double
xmin = 0
xmax = Adodc5.Recordset.RecordCount
xstep = 1
num_x = (xmax - xmin) / xstep
ReDim Values(1 To num_x, 0)
X = xmin
Adodc5.Recordset.MoveFirst
For i = 1 To num_x
    Values(i, 0) = Adodc5.Recordset.Fields("MSE Akhir").Value
    X = X + xstep
    Adodc5.Recordset.MoveNext
Next i

Form5.MSChart1.RowCount = 1
Form5.MSChart1.ColumnCount = num_x
Form5.MSChart1.ChartData = Values

'keterangan
With Form5.MSChart1.Legend
.Location.Visible = False
.Location.LocationType = VtChLocationTypeRight
.TextLayout.HorzAlignment = VtHorizontalAlignmentRight
End With

Form5.MSChart1.Plot.SeriesCollection(1).LegendText = "Grafik MSE"
Form5.Show
Form4.Hide
End Sub

Private Sub Command2_Click()
Form6.Show
Form4.Hide
End Sub

Private Sub Command4_Click()
Form3.Show
Form4.Hide
End Sub

Private Sub Command3_Click()
Dim iterasi As Integer, unitinput As Double, unithidden As Double, unitoutput As Double, data As Integer, mp As Integer
Dim myup(3) As Double, jmsep(3, 1000) As Double, minimalmsep(3) As Double, msejp(3) As Double
Dim selisihp(3, 100) As Double, ermsep(3) As Double, sumvp As Double, sumwp As Double
Dim selisihpu(3, 100) As Double, ermsepu(3) As Double, sumvpu As Double, sumwpu As Double
Dim slopezp(3, 100, 3) As Double, slopeyp(3, 100, 3) As Double, slopezpu(3, 100, 3) As Double, slopeypu(3, 100, 3) As Double
Dim dwp(3, 100, 2, 1) As Double, dvp(3, 100, 2, 2) As Double
Dim jacobip(3, 100, 9) As Double, transposep(3, 9, 100) As Double, hessianp(3, 9, 9) As Double, idenp(3, 18, 18) As Double
Dim hessp(3, 9, 9) As Double, gradienp(3, 9, 1) As Double, elemenp(3, 18, 18) As Double, inversp(3, 18, 18) As Double
Dim deltabobotp(3, 9, 1) As Double, errorval(60, 5) As Double, jumlah1 As Double, nomor As Integer, datauji As Integer

data = 57

'menampilkan inputan parameter LM
Form8.Label4.Caption = mu
Form8.Label6.Caption = beta
Form8.Label8.Caption = epoch
Form8.Label10.Caption = err

'pola ke 0 to 57 training
'pola ke 58 to 93 testing

'pola data prediksi
For i = 1 To 3
    If (i = 1) Then
        For j = 0 To 93
            poladataprediksi(i, j, 1) = normalisasi(j, 1)
            poladataprediksi(i, j, 2) = normalisasi(j + 1, 1)
            targetprediksi(i, j) = normalisasi(j + 2, 1)
            If (j > 57) Then
                poladatauji(i, j - 58, 1) = normalisasi(j, 1)
                poladatauji(i, j - 58, 2) = normalisasi(j + 1, 1)
                targetuji(i, j - 58) = normalisasi(j + 2, 1)
            End If
        Next j
    ElseIf (i = 2) Then
        For j = 0 To 93
            poladataprediksi(i, j, 1) = normalisasi(j, 2)
            poladataprediksi(i, j, 2) = normalisasi(j + 1, 2)
            targetprediksi(i, j) = normalisasi(j + 2, 2)
            If (j > 57) Then
                poladatauji(i, j - 58, 1) = normalisasi(j, 2)
                poladatauji(i, j - 58, 2) = normalisasi(j + 1, 2)
                targetuji(i, j - 58) = normalisasi(j + 2, 2)
            End If
        Next j
    ElseIf (i = 3) Then
        For j = 0 To 93
            poladataprediksi(i, j, 1) = normalisasi(j, 3)
            poladataprediksi(i, j, 2) = normalisasi(j + 1, 3)
            targetprediksi(i, j) = normalisasi(j + 2, 3)
            If (j > 57) Then
                poladatauji(i, j - 58, 1) = normalisasi(j, 3)
                poladatauji(i, j - 58, 2) = normalisasi(j + 1, 3)
                targetuji(i, j - 58) = normalisasi(j + 2, 3)
            End If
        Next j
    End If
Next i

'menampilkan pola data pelatihan
'populasi S
For j = 0 To data
    Form8.Adodc3.Recordset.AddNew
    Form8.Adodc3.Recordset.Fields("Pola") = j + 1
    Form8.Adodc3.Recordset.Fields("x1") = poladataprediksi(1, j, 1)
    Form8.Adodc3.Recordset.Fields("x2") = poladataprediksi(1, j, 2)
    Form8.Adodc3.Recordset.Fields("x3") = targetprediksi(1, j)
    Form8.Adodc3.Recordset.Update
    Form8.DataGrid3.Refresh
Next j
'populasi I
For j = 0 To data
    Form8.Adodc4.Recordset.AddNew
    Form8.Adodc4.Recordset.Fields("Pola") = j + 1
    Form8.Adodc4.Recordset.Fields("x1") = poladataprediksi(2, j, 1)
    Form8.Adodc4.Recordset.Fields("x2") = poladataprediksi(2, j, 2)
    Form8.Adodc4.Recordset.Fields("x3") = targetprediksi(2, j)
    Form8.Adodc4.Recordset.Update
    Form8.DataGrid4.Refresh
Next j
'populasi R
For j = 0 To data
    Form8.Adodc5.Recordset.AddNew
    Form8.Adodc5.Recordset.Fields("Pola") = j + 1
    Form8.Adodc5.Recordset.Fields("x1") = poladataprediksi(3, j, 1)
    Form8.Adodc5.Recordset.Fields("x2") = poladataprediksi(3, j, 2)
    Form8.Adodc5.Recordset.Fields("x3") = targetprediksi(3, j)
    Form8.Adodc5.Recordset.Update
    Form8.DataGrid5.Refresh
Next j

'bobot dan bias awal untuk prediksi
For pop = 1 To 3
    For i = 1 To 2
        For j = 0 To 2
            bbvp(pop, j, i) = bbvn(pop, j, i)
            bbwp(pop, j, i) = bbwn(pop, j, 1)
        Next j
    Next i
Next pop

'menampilkan bobot dan bias awal
'S
Form8.v11s.Text = Round(bbvp(1, 1, 1), 6)
Form8.v12s.Text = Round(bbvp(1, 1, 2), 6)
Form8.v21s.Text = Round(bbvp(1, 2, 1), 6)
Form8.v22s.Text = Round(bbvp(1, 2, 2), 6)
Form8.v01s.Text = Round(bbvp(1, 0, 1), 6)
Form8.v02s.Text = Round(bbvp(1, 0, 2), 6)
Form8.w11s.Text = Round(bbwp(1, 1, 1), 6)
Form8.w21s.Text = Round(bbwp(1, 2, 1), 6)
Form8.w01s.Text = Round(bbwp(1, 0, 1), 6)
'I
Form8.v11i.Text = Round(bbvp(2, 1, 1), 6)
Form8.v12i.Text = Round(bbvp(2, 1, 2), 6)
Form8.v21i.Text = Round(bbvp(2, 2, 1), 6)
Form8.v22i.Text = Round(bbvp(2, 2, 2), 6)
Form8.v01i.Text = Round(bbvp(2, 0, 1), 6)
Form8.v02i.Text = Round(bbvp(2, 0, 2), 6)
Form8.w11i.Text = Round(bbwp(2, 1, 1), 6)
Form8.w21i.Text = Round(bbwp(2, 2, 1), 6)
Form8.w01i.Text = Round(bbwp(2, 0, 1), 6)
'R
Form8.v11r.Text = Round(bbvp(3, 1, 1), 6)
Form8.v12r.Text = Round(bbvp(3, 1, 2), 6)
Form8.v21r.Text = Round(bbvp(3, 2, 1), 6)
Form8.v22r.Text = Round(bbvp(3, 2, 2), 6)
Form8.v01r.Text = Round(bbvp(3, 0, 1), 6)
Form8.v02r.Text = Round(bbvp(3, 0, 2), 6)
Form8.w11r.Text = Round(bbwp(3, 1, 1), 6)
Form8.w21r.Text = Round(bbwp(3, 2, 1), 6)
Form8.w01r.Text = Round(bbwp(3, 0, 1), 6)

unitinput = 2
unithidden = 2
unitoutput = 1
iterasi = 0

'Proses LM untuk prediksi
Do
iterasi = iterasi + 1
If iterasi = 1 Then
    For pop = 1 To 3
        myup(pop) = mu
    Next pop
End If
For pop = 1 To 3
    mp = 0
    Do
    'proses feedforward
    'ermsep(pop) = 0
    For pola = 0 To data
        For i = 1 To unithidden
            sumvp = 0
            For j = 1 To unitinput
                sumvp = sumvp + poladataprediksi(pop, pola, j) * bbvp(pop, j, i)
            Next j
            zinp(pop, pola, i) = bbvp(pop, 0, i) + sumvp
            zp(pop, pola, i) = sigmoidbipolar(zinp(pop, pola, i))
            slopezp(pop, pola, i) = slope(zp(pop, pola, i))
        Next i
        For i = 1 To unitoutput
            sumwp = 0
            For j = 1 To unithidden
                sumwp = sumwp + zp(pop, pola, j) * bbwp(pop, j, i)
            Next j
            yinp(pop, pola, i) = bbwp(pop, 0, 1) + sumwp
            yp(pop, pola, i) = sigmoidbipolar(yinp(pop, pola, i))
            slopeyp(pop, pola, i) = slope(yp(pop, pola, i))
            outputsp(pop, pola) = yp(pop, pola, i)
            'selisih (t - y)
            'selisihp(pop, pola) = targetprediksi(pop, pola) - outputsp(pop, pola)
            'ermsep(pop) = ermsep(pop) + (selisihp(pop, pola)) ^ 2
        Next i
    Next pola
    
    'feedforward data uji
    ermsep(pop) = 0
    For pola = 0 To 35
       For i = 1 To unithidden
            sumvpu = 0
            For j = 1 To unitinput
               sumvpu = sumvpu + poladatauji(pop, pola, j) * bbvp(pop, j, i)
            Next j
            zinpu(pop, pola, i) = bbvp(pop, 0, i) + sumvpu
            zpu(pop, pola, i) = sigmoidbipolar(zinpu(pop, pola, i))
        Next i
        For i = 1 To unitoutput
            sumwpu = 0
            For j = 1 To unithidden
                sumwpu = sumwpu + zpu(pop, pola, j) * bbwp(pop, j, i)
            Next j
            yinpu(pop, pola, i) = bbwp(pop, 0, 1) + sumwpu
            ypu(pop, pola, i) = sigmoidbipolar(yinpu(pop, pola, i))
            outputspu(pop, pola) = ypu(pop, pola, i)
            'selisih (t-y)
            selisihp(pop, pola) = targetuji(pop, pola) - outputspu(pop, pola)
            ermsep(pop) = ermsep(pop) + (selisihp(pop, pola)) ^ 2
        Next i
    Next pola
    
    'menghitung MSE
    jmsep(pop, iterasi) = ermsep(pop) / (36) 'nilai mse populasi ke-pop iterasi ke-iterasi
    
    'proses backward
    For pola = 0 To data
        For i = 1 To unithidden
            For j = 1 To unitoutput
                dwp(pop, pola, i, j) = slopeyp(pop, pola, j) * zp(pop, pola, i)
                dwp(pop, pola, 0, j) = slopeyp(pop, pola, j)
            Next j
        Next i
        For i = 1 To unitinput
            For j = 1 To unithidden
                dvp(pop, pola, i, j) = slopeyp(pop, pola, 1) * slopezp(pop, pola, j) * poladataprediksi(pop, pola, i)
                dvp(pop, pola, 0, j) = slopeyp(pop, pola, 1) * slopezp(pop, pola, j)
            Next j
        Next i
        'membentuk matriks jacobi
        jacobip(pop, pola, 1) = -dvp(pop, pola, 1, 1)
        jacobip(pop, pola, 2) = -dvp(pop, pola, 1, 2)
        jacobip(pop, pola, 3) = -dvp(pop, pola, 2, 1)
        jacobip(pop, pola, 4) = -dvp(pop, pola, 2, 2)
        jacobip(pop, pola, 5) = -dvp(pop, pola, 0, 1)
        jacobip(pop, pola, 6) = -dvp(pop, pola, 0, 2)
        jacobip(pop, pola, 7) = -dwp(pop, pola, 1, 1)
        jacobip(pop, pola, 8) = -dwp(pop, pola, 2, 1)
        jacobip(pop, pola, 9) = -dwp(pop, pola, 0, 1)
    Next pola

    'Proses pembaruan bobot dan bias
    If iterasi > 1 Then
        'transpose jacobi
        For pola = 0 To data
            For j = 1 To 9
                transposep(pop, j, pola) = jacobip(pop, pola, j)
            Next j
        Next pola
        'perkalian matriks jacobi (J^T*J)
        For i = 1 To 9
            For j = 1 To 9
                hessp(pop, i, j) = 0
                For pola = 0 To data
                    hessp(pop, i, j) = hessp(pop, i, j) + transposep(pop, i, pola) * jacobip(pop, pola, j)
                Next pola
            Next j
        Next i
        'matriks identitas 9x9
        For i = 1 To 9
            For j = 1 To 9
                If (i = j) Then
                    idenp(pop, i, j) = 1
                Else
                    idenp(pop, i, j) = 0
                End If
            Next j
        Next i
        'menghitung (H+mu*I)
        For i = 1 To 9
            For j = 1 To 9
                hessianp(pop, i, j) = hessp(pop, i, j) + (myup(pop) * idenp(pop, i, j))
            Next j
        Next i
        'menghitung invers hessian A|I
        For i = 1 To 9
            For j = 1 To 18
                If j <= 9 Then
                    elemenp(pop, i, j) = hessianp(pop, i, j)
                Else
                    elemenp(pop, i, j) = idenp(pop, i, j - 9)
                End If
            Next j
        Next i
        For i = 1 To 9
            For j = 1 To 18
                If (i <> j) Then
                    elemenp(pop, i, j) = elemenp(pop, i, j) / elemenp(pop, i, i)
                End If
            Next j
            For j = 1 To 18
                If (i = j) Then
                    elemenp(pop, i, j) = 1
                End If
            Next j
            For l = 1 To 9
                If (i <> l) Then
                    For j = i + 1 To 18
                        elemenp(pop, l, j) = elemenp(pop, l, j) - (elemenp(pop, i, j) * elemenp(pop, l, i))
                    Next j
                End If
            Next l
            For k = 1 To 9
                If (i <> k) Then
                    elemenp(pop, k, i) = 0
                End If
            Next k
        Next i
        'hasil invers
        For i = 1 To 9
            For j = 1 To 9
                inversp(pop, i, j) = elemenp(pop, i, j + 9)
            Next j
        Next i
        'menghitung gradien g
        For i = 1 To 9
            For j = 1 To 1
                gradienp(pop, i, j) = 0
                For pola = 0 To data
                    gradienp(pop, i, j) = gradienp(pop, i, j) + transposep(pop, i, pola) * selisihp(pop, pola)
                Next pola
            Next j
        Next i
        'menghitung delta bobot dan bias
        For i = 1 To 9
            For j = 1 To 1
                deltabobotp(pop, i, j) = 0
                For k = 1 To 9
                    deltabobotp(pop, i, j) = deltabobotp(pop, i, j) + inversp(pop, i, k) * gradienp(pop, k, j)
                Next k
            Next j
        Next i
        For i = 1 To 9
            If (deltabobotp(pop, i, 1) < -1 Or deltabobotp(pop, i, 1) > 1) Then
                deltabobotp(pop, i, 1) = sigmoidbipolar(deltabobotp(pop, i, 1))
            End If
        Next i
        'update bobot dan bias
        bbvp(pop, 1, 1) = bbvp(pop, 1, 1) - deltabobotp(pop, 1, 1)
        bbvp(pop, 1, 2) = bbvp(pop, 1, 2) - deltabobotp(pop, 2, 1)
        bbvp(pop, 2, 1) = bbvp(pop, 2, 1) - deltabobotp(pop, 3, 1)
        bbvp(pop, 2, 2) = bbvp(pop, 2, 2) - deltabobotp(pop, 4, 1)
        bbvp(pop, 0, 1) = bbvp(pop, 0, 1) - deltabobotp(pop, 5, 1)
        bbvp(pop, 0, 2) = bbvp(pop, 0, 2) - deltabobotp(pop, 6, 1)
        bbwp(pop, 1, 1) = bbwp(pop, 1, 1) - deltabobotp(pop, 7, 1)
        bbwp(pop, 2, 1) = bbwp(pop, 2, 1) - deltabobotp(pop, 8, 1)
        bbwp(pop, 0, 1) = bbwp(pop, 0, 1) - deltabobotp(pop, 9, 1)
        For i = 0 To 2
            For j = 1 To 2
                If bbvp(pop, i, j) < -1 Or bbvp(pop, i, j) > 1 Then
                  bbvp(pop, i, j) = sigmoidbipolar(bbvp(pop, i, j))
               End If
            Next j
           For j = 1 To 1
                If bbwp(pop, i, j) < -1 Or bbwp(pop, i, j) > 1 Then
                    bbwp(pop, i, j) = sigmoidbipolar(bbwp(pop, i, j))
                End If
            Next j
        Next i
    End If 'iterasi>1
    If (msejp(pop) < jmsep(pop, iterasi)) Then 'perbandingan antara mse E(k) < mse E(k+1)
        myup(pop) = myup(pop) * beta
    End If
    mp = mp + 1
    
    Loop While msejp(pop) < jmsep(pop, iterasi) And mp <= 5
    
    If (iterasi > 0) Then
        myup(pop) = myup(pop) / beta
    End If
Next pop

'minimal MSE
If (iterasi = 1) Then
    For pop = 1 To 3
        minimalmsep(pop) = jmsep(pop, iterasi) ' jmse adalah nilai mse populasi ke-pop tiap iterasi
    Next pop
Else
    For pop = 1 To 3
        If jmsep(pop, iterasi) < minimalmsep(pop) Then
           'For i = 0 To 2
           '     For j = 1 To 2
           '         If bbvp(pop, i, j) < -1 Or bbvp(pop, i, j) > 1 Then
           '             bbvp(pop, i, j) = sigmoidbipolar(bbvp(pop, i, j))
           '         End If
           '     Next j
           '     For j = 1 To 1
           '         If bbwp(pop, i, j) < -1 Or bbwp(pop, i, j) > 1 Then
           '             bbwp(pop, i, j) = sigmoidbipolar(bbwp(pop, i, j))
           '         End If
           '     Next j
           ' Next i
            minimalmsep(pop) = jmsep(pop, iterasi)
            bbvnp(pop, 1, 1) = bbvp(pop, 1, 1)
            bbvnp(pop, 1, 2) = bbvp(pop, 1, 2)
            bbvnp(pop, 2, 1) = bbvp(pop, 2, 1)
            bbvnp(pop, 2, 2) = bbvp(pop, 2, 2)
            bbvnp(pop, 0, 1) = bbvp(pop, 0, 1)
            bbvnp(pop, 0, 2) = bbvp(pop, 0, 2)
            bbwnp(pop, 1, 1) = bbwp(pop, 1, 1)
            bbwnp(pop, 2, 1) = bbwp(pop, 2, 1)
            bbwnp(pop, 0, 1) = bbwp(pop, 0, 1)
            'For j = 0 To data
            '    outputfixp(pop, j) = outputsp(pop, j)
            'Next j
        End If
    Next pop
End If

For pop = 1 To 3
    msejp(pop) = jmsep(pop, iterasi) 'msej adalah nilai mse E(k)
Next pop

'menghitung MSE akhir
msep = (minimalmsep(1) + minimalmsep(2) + minimalmsep(3)) / 3
Form8.Adodc1.Recordset.AddNew
Form8.Adodc1.Recordset.Fields("Iterasi") = iterasi
Form8.Adodc1.Recordset.Fields("MSE S") = Round(minimalmsep(1), 6)
Form8.Adodc1.Recordset.Fields("MSE I") = Round(minimalmsep(2), 6)
Form8.Adodc1.Recordset.Fields("MSE R") = Round(minimalmsep(3), 6)
Form8.Adodc1.Recordset.Fields("MSE Akhir") = Round(msep, 6)
Form8.Adodc1.Recordset.Update
Form8.DataGrid1.Refresh
Loop While iterasi < epoch And msep > err

Form8.Label13.Caption = Round(msep, 6)

Form8.v11bs.Text = Round(bbvnp(1, 1, 1), 6)
Form8.v12bs.Text = Round(bbvnp(1, 1, 2), 6)
Form8.v21bs.Text = Round(bbvnp(1, 2, 1), 6)
Form8.v22bs.Text = Round(bbvnp(1, 2, 2), 6)
Form8.v01bs.Text = Round(bbvnp(1, 0, 1), 6)
Form8.v02bs.Text = Round(bbvnp(1, 0, 2), 6)
Form8.w11bs.Text = Round(bbwnp(1, 1, 1), 6)
Form8.w21bs.Text = Round(bbwnp(1, 2, 1), 6)
Form8.w01bs.Text = Round(bbwnp(1, 0, 1), 6)

Form8.v11bi.Text = Round(bbvnp(2, 1, 1), 6)
Form8.v12bi.Text = Round(bbvnp(2, 1, 2), 6)
Form8.v21bi.Text = Round(bbvnp(2, 2, 1), 6)
Form8.v22bi.Text = Round(bbvnp(2, 2, 2), 6)
Form8.v01bi.Text = Round(bbvnp(2, 0, 1), 6)
Form8.v02bi.Text = Round(bbvnp(2, 0, 2), 6)
Form8.w11bi.Text = Round(bbwnp(2, 1, 1), 6)
Form8.w21bi.Text = Round(bbwnp(2, 2, 1), 6)
Form8.w01bi.Text = Round(bbwnp(2, 0, 1), 6)

Form8.v11br.Text = Round(bbvnp(3, 1, 1), 6)
Form8.v12br.Text = Round(bbvnp(3, 1, 2), 6)
Form8.v21br.Text = Round(bbvnp(3, 2, 1), 6)
Form8.v22br.Text = Round(bbvnp(3, 2, 2), 6)
Form8.v01br.Text = Round(bbvnp(3, 0, 1), 6)
Form8.v02br.Text = Round(bbvnp(3, 0, 2), 6)
Form8.w11br.Text = Round(bbwnp(3, 1, 1), 6)
Form8.w21br.Text = Round(bbwnp(3, 2, 1), 6)
Form8.w01br.Text = Round(bbwnp(3, 0, 1), 6)

'Proses Testing
'proses feedforward data validasi
For pop = 1 To 3
    For pola = 0 To 35
        For i = 1 To unithidden
            sumvpu = 0
            For j = 1 To unitinput
                sumvpu = sumvpu + poladatauji(pop, pola, j) * bbvnp(pop, j, i)
            Next j
            zinpu(pop, pola, i) = bbvnp(pop, 0, i) + sumvpu
            zpu(pop, pola, i) = sigmoidbipolar(zinpu(pop, pola, i))
        Next i
        For i = 1 To unitoutput
            sumwpu = 0
            For j = 1 To unithidden
                sumwpu = sumwpu + zpu(pop, pola, j) * bbwnp(pop, j, i)
            Next j
            yinpu(pop, pola, i) = bbwnp(pop, 0, 1) + sumwpu
            ypu(pop, pola, i) = sigmoidbipolar(yinpu(pop, pola, i))
            outputsp(pop, pola) = ypu(pop, pola, i)
        Next i
    Next pola
    For pola = 0 To 35
        outputfixp(pop, pola) = outputsp(pop, pola)
   Next pola
Next pop

datauji = 35

'denormalisasi data
For pop = 1 To 3
    For pola = 0 To datauji
        denormalisasiprediksi(pola, pop) = ((outputfixp(pop, pola) + 1) * (maks(pop) - mini(pop)) / 2) + mini(pop)
    Next pola
Next pop

'MMRE Akhir
For pola = 0 To datauji
    jumlah1 = 0
    For pop = 1 To 3
        errorval(pola, pop) = mre(dataasli(pola + 60, pop), denormalisasiprediksi(pola, pop))
        jumlah1 = jumlah1 + errorval(pola, pop)
    Next pop
    errorval(pola, 4) = jumlah1 / 3
Next pola
jumlah1 = 0
For pola = 0 To datauji
    jumlah1 = jumlah1 + errorval(pola, 4)
Next pola
mmreakhirp = Round(jumlah1 / (datauji + 1), 6)

'menampilkan nilai MMRE
Form7.Label3.Caption = mmreakhirp
Form7.Label8.Caption = Round(msepu, 6)
For i = 0 To datauji
    Form7.Adodc2.Recordset.AddNew
    Form7.Adodc2.Recordset.Fields("Bulan Ke") = 61 + i
    Form7.Adodc2.Recordset.Fields("S") = Round(denormalisasiprediksi(i, 1), 0)
    Form7.Adodc2.Recordset.Fields("I") = Round(denormalisasiprediksi(i, 2), 0)
    Form7.Adodc2.Recordset.Fields("R") = Round(denormalisasiprediksi(i, 3), 0)
    Form7.Adodc2.Recordset.Update
    Form7.DataGrid2.Refresh
Next i

'Save to database
Form4.Adodc7.Recordset.MoveLast
no = Form4.Adodc7.Recordset.Fields("no").Value
no = no + 1
Form4.Adodc7.Recordset.AddNew
Form4.Adodc7.Recordset.Fields("no") = no
Form4.Adodc7.Recordset.Fields("parameter ke") = nofix
Form4.Adodc7.Recordset.Fields("jumlah pelajar") = nfix
Form4.Adodc7.Recordset.Fields("maxit") = maxitfix
Form4.Adodc7.Recordset.Fields("delta") = deltafix
Form4.Adodc7.Recordset.Fields("omega") = omegafix
Form4.Adodc7.Recordset.Fields("miu") = miufix
Form4.Adodc7.Recordset.Fields("epsilon") = epsilonfix
Form4.Adodc7.Recordset.Fields("teta") = deltafix
Form4.Adodc7.Recordset.Fields("error") = mmrefix
Form4.Adodc7.Recordset.Fields("parameter lm") = mu
Form4.Adodc7.Recordset.Fields("faktor beta") = beta
Form4.Adodc7.Recordset.Fields("maks epoch") = epoch
Form4.Adodc7.Recordset.Fields("epoch") = epochakhir
Form4.Adodc7.Recordset.Fields("batas error") = err
Form4.Adodc7.Recordset.Fields("v11s") = Round(bbvnp(1, 1, 1), 6)
Form4.Adodc7.Recordset.Fields("v12s") = Round(bbvnp(1, 1, 2), 6)
Form4.Adodc7.Recordset.Fields("v21s") = Round(bbvnp(1, 2, 1), 6)
Form4.Adodc7.Recordset.Fields("v22s") = Round(bbvnp(1, 2, 2), 6)
Form4.Adodc7.Recordset.Fields("v01s") = Round(bbvnp(1, 0, 1), 6)
Form4.Adodc7.Recordset.Fields("v02s") = Round(bbvnp(1, 0, 2), 6)
Form4.Adodc7.Recordset.Fields("w11s") = Round(bbwnp(1, 1, 1), 6)
Form4.Adodc7.Recordset.Fields("w21s") = Round(bbwnp(1, 2, 1), 6)
Form4.Adodc7.Recordset.Fields("w01s") = Round(bbwnp(1, 0, 1), 6)
Form4.Adodc7.Recordset.Fields("v11i") = Round(bbvnp(2, 1, 1), 6)
Form4.Adodc7.Recordset.Fields("v12i") = Round(bbvnp(2, 1, 2), 6)
Form4.Adodc7.Recordset.Fields("v21i") = Round(bbvnp(2, 2, 1), 6)
Form4.Adodc7.Recordset.Fields("v22i") = Round(bbvnp(2, 2, 2), 6)
Form4.Adodc7.Recordset.Fields("v01i") = Round(bbvnp(2, 0, 1), 6)
Form4.Adodc7.Recordset.Fields("v02i") = Round(bbvnp(2, 0, 2), 6)
Form4.Adodc7.Recordset.Fields("w11i") = Round(bbwnp(2, 1, 1), 6)
Form4.Adodc7.Recordset.Fields("w21i") = Round(bbwnp(2, 2, 1), 6)
Form4.Adodc7.Recordset.Fields("w01i") = Round(bbwnp(2, 0, 1), 6)
Form4.Adodc7.Recordset.Fields("v11r") = Round(bbvnp(3, 1, 1), 6)
Form4.Adodc7.Recordset.Fields("v12r") = Round(bbvnp(3, 1, 2), 6)
Form4.Adodc7.Recordset.Fields("v21r") = Round(bbvnp(3, 2, 1), 6)
Form4.Adodc7.Recordset.Fields("v22r") = Round(bbvnp(3, 2, 2), 6)
Form4.Adodc7.Recordset.Fields("v01r") = Round(bbvnp(3, 0, 1), 6)
Form4.Adodc7.Recordset.Fields("v02r") = Round(bbvnp(3, 0, 2), 6)
Form4.Adodc7.Recordset.Fields("w11r") = Round(bbwnp(3, 1, 1), 6)
Form4.Adodc7.Recordset.Fields("w21r") = Round(bbwnp(3, 2, 1), 6)
Form4.Adodc7.Recordset.Fields("w01r") = Round(bbwnp(3, 0, 1), 6)
Form4.Adodc7.Recordset.Fields("mse") = mseakhir
Form4.Adodc7.Recordset.Fields("MMRE Validasi") = mmreakhir
Form4.Adodc7.Recordset.Fields("mse prediksi") = Round(msep, 6)
Form4.Adodc7.Recordset.Fields("MMRE Prediksi") = mmreakhirp
Form4.Adodc7.Recordset.Update
Form4.DataGrid8.Refresh

For i = 0 To datauji
Form7.Adodc3.Recordset.AddNew
Form7.Adodc3.Recordset.Fields("no") = no
Form7.Adodc3.Recordset.Fields("parameter") = nofix
Form7.Adodc3.Recordset.Fields("Bulan Ke-") = 61 + i
Form7.Adodc3.Recordset.Fields("S") = Round(denormalisasiprediksi(i, 1), 0)
Form7.Adodc3.Recordset.Fields("I") = Round(denormalisasiprediksi(i, 2), 0)
Form7.Adodc3.Recordset.Fields("R") = Round(denormalisasiprediksi(i, 3), 0)
Form7.Adodc3.Recordset.Update
Form7.DataGrid3.Refresh
Next i

Form8.v11s.Visible = False
Form8.v12s.Visible = False
Form8.v21s.Visible = False
Form8.v22s.Visible = False
Form8.v01s.Visible = False
Form8.v02s.Visible = False
Form8.w11s.Visible = False
Form8.w21s.Visible = False
Form8.w01s.Visible = False
Form8.v11i.Visible = False
Form8.v12i.Visible = False
Form8.v21i.Visible = False
Form8.v22i.Visible = False
Form8.v01i.Visible = False
Form8.v02i.Visible = False
Form8.w11i.Visible = False
Form8.w21i.Visible = False
Form8.w01i.Visible = False
Form8.v11r.Visible = False
Form8.v12r.Visible = False
Form8.v21r.Visible = False
Form8.v22r.Visible = False
Form8.v01r.Visible = False
Form8.v02r.Visible = False
Form8.w11r.Visible = False
Form8.w21r.Visible = False
Form8.w01r.Visible = False

Form8.v11bs.Visible = False
Form8.v12bs.Visible = False
Form8.v21bs.Visible = False
Form8.v22bs.Visible = False
Form8.v01bs.Visible = False
Form8.v02bs.Visible = False
Form8.w11bs.Visible = False
Form8.w21bs.Visible = False
Form8.w01bs.Visible = False
Form8.v11bi.Visible = False
Form8.v12bi.Visible = False
Form8.v21bi.Visible = False
Form8.v22bi.Visible = False
Form8.v01bi.Visible = False
Form8.v02bi.Visible = False
Form8.w11bi.Visible = False
Form8.w21bi.Visible = False
Form8.w01bi.Visible = False
Form8.v11br.Visible = False
Form8.v12br.Visible = False
Form8.v21br.Visible = False
Form8.v22br.Visible = False
Form8.v01br.Visible = False
Form8.v02br.Visible = False
Form8.w11br.Visible = False
Form8.w21br.Visible = False
Form8.w01br.Visible = False

Form8.DataGrid3.Visible = False
Form8.DataGrid4.Visible = False
Form8.DataGrid5.Visible = False

Form8.Option1.Value = False
Form8.Option2.Value = False
Form8.Option3.Value = False

Form8.Show
Form4.Hide

End Sub

Private Sub Option1_Click()
Option1.Value = True
Option2.Value = False
Option3.Value = False

DataGrid3.Visible = True
DataGrid4.Visible = False
DataGrid5.Visible = False

v11s.Visible = True
v12s.Visible = True
v21s.Visible = True
v22s.Visible = True
v01s.Visible = True
v02s.Visible = True
w11s.Visible = True
w21s.Visible = True
w01s.Visible = True
v11i.Visible = False
v12i.Visible = False
v21i.Visible = False
v22i.Visible = False
v01i.Visible = False
v02i.Visible = False
w11i.Visible = False
w21i.Visible = False
w01i.Visible = False
v11r.Visible = False
v12r.Visible = False
v21r.Visible = False
v22r.Visible = False
v01r.Visible = False
v02r.Visible = False
w11r.Visible = False
w21r.Visible = False
w01r.Visible = False

v11bs.Visible = True
v12bs.Visible = True
v21bs.Visible = True
v22bs.Visible = True
v01bs.Visible = True
v02bs.Visible = True
w11bs.Visible = True
w21bs.Visible = True
w01bs.Visible = True
v11bi.Visible = False
v12bi.Visible = False
v21bi.Visible = False
v22bi.Visible = False
v01bi.Visible = False
v02bi.Visible = False
w11bi.Visible = False
w21bi.Visible = False
w01bi.Visible = False
v11br.Visible = False
v12br.Visible = False
v21br.Visible = False
v22br.Visible = False
v01br.Visible = False
v02br.Visible = False
w11br.Visible = False
w21br.Visible = False
w01br.Visible = False


End Sub

Private Sub Option2_Click()
Option1.Value = False
Option2.Value = True
Option3.Value = False

DataGrid3.Visible = False
DataGrid4.Visible = True
DataGrid5.Visible = False

v11s.Visible = False
v12s.Visible = False
v21s.Visible = False
v22s.Visible = False
v01s.Visible = False
v02s.Visible = False
w11s.Visible = False
w21s.Visible = False
w01s.Visible = False
v11i.Visible = True
v12i.Visible = True
v21i.Visible = True
v22i.Visible = True
v01i.Visible = True
v02i.Visible = True
w11i.Visible = True
w21i.Visible = True
w01i.Visible = True
v11r.Visible = False
v12r.Visible = False
v21r.Visible = False
v22r.Visible = False
v01r.Visible = False
v02r.Visible = False
w11r.Visible = False
w21r.Visible = False
w01r.Visible = False

v11bs.Visible = False
v12bs.Visible = False
v21bs.Visible = False
v22bs.Visible = False
v01bs.Visible = False
v02bs.Visible = False
w11bs.Visible = False
w21bs.Visible = False
w01bs.Visible = False
v11bi.Visible = True
v12bi.Visible = True
v21bi.Visible = True
v22bi.Visible = True
v01bi.Visible = True
v02bi.Visible = True
w11bi.Visible = True
w21bi.Visible = True
w01bi.Visible = True
v11br.Visible = False
v12br.Visible = False
v21br.Visible = False
v22br.Visible = False
v01br.Visible = False
v02br.Visible = False
w11br.Visible = False
w21br.Visible = False
w01br.Visible = False

End Sub

Private Sub Option3_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = True

DataGrid3.Visible = False
DataGrid4.Visible = False
DataGrid5.Visible = True

v11s.Visible = False
v12s.Visible = False
v21s.Visible = False
v22s.Visible = False
v01s.Visible = False
v02s.Visible = False
w11s.Visible = False
w21s.Visible = False
w01s.Visible = False
v11i.Visible = False
v12i.Visible = False
v21i.Visible = False
v22i.Visible = False
v01i.Visible = False
v02i.Visible = False
w11i.Visible = False
w21i.Visible = False
w01i.Visible = False
v11r.Visible = True
v12r.Visible = True
v21r.Visible = True
v22r.Visible = True
v01r.Visible = True
v02r.Visible = True
w11r.Visible = True
w21r.Visible = True
w01r.Visible = True

v11bs.Visible = False
v12bs.Visible = False
v21bs.Visible = False
v22bs.Visible = False
v01bs.Visible = False
v02bs.Visible = False
w11bs.Visible = False
w21bs.Visible = False
w01bs.Visible = False
v11bi.Visible = False
v12bi.Visible = False
v21bi.Visible = False
v22bi.Visible = False
v01bi.Visible = False
v02bi.Visible = False
w11bi.Visible = False
w21bi.Visible = False
w01bi.Visible = False
v11br.Visible = True
v12br.Visible = True
v21br.Visible = True
v22br.Visible = True
v01br.Visible = True
v02br.Visible = True
w11br.Visible = True
w21br.Visible = True
w01br.Visible = True

End Sub

