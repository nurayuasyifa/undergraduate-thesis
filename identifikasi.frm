VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "proses levenberg-marquardt"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6465
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Left            =   4080
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Inisialisasi Parameter LM :"
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   3255
      Begin VB.TextBox err 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Text            =   "Text8"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox epoch 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Text            =   "Text7"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox lambda 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Text            =   "Text6"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Batas Error"
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
         Left            =   480
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Maksimal Iterasi"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Parameter LM"
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
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Parameter Model Terbaik : "
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2895
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Text            =   "Text5"
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Text            =   "Text4"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   360
         Picture         =   "identifikasi.frx":0000
         ScaleHeight     =   6.165
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   1.72
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   480
      TabIndex        =   15
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
Form4.Show
Form3.Hide
End Sub
