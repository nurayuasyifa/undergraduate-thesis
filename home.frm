VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TLBO-LM"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "home.frx":0000
   ScaleHeight     =   5115
   ScaleMode       =   0  'User
   ScaleWidth      =   7606.529
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Gabriola"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "home.frx":23E6B
      Left            =   2400
      List            =   "home.frx":23E6D
      TabIndex        =   9
      Text            =   "Pilih Proses"
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Universitas Airlangga"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fakultas Sains dan Teknologi"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Departemen Matematika"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Asri Bekti Pratiwi, S.Si., M.Si"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Auli Damayanti, S.Si., M.Si"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dosen Pembimbing :"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "081711233092"
      BeginProperty Font 
         Name            =   "Ink Free"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NUR AYU ASYIFA"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"home.frx":23E6F
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_click()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HP\OneDrive\Skripsi\Program\db.mdb;Persist Security Info=False")
con.Execute ("DELETE FROM TLBO_parameter_model;")
con.Execute ("DELETE FROM TLBO_parameter_model_faseguru;")
con.Execute ("DELETE FROM TLBO_parameter_model_fasepelajar;")
con.Execute ("DELETE FROM TLBO_Interaksi_Pelajar;")
con.Execute ("DELETE FROM TLBO_Interaksi_Perbandingan;")
con.Execute ("DELETE FROM TLBO_Interaksi_Perbandingan_Hasil;")
con.Execute ("DELETE FROM TLBO_parameter_terbaik_iterasi;")
con.Execute ("DELETE FROM Normalisasi;")
con.Execute ("DELETE FROM Denormalisasi;")
con.Execute ("DELETE FROM LM_MSE;")
con.Execute ("DELETE FROM LM_Poladata_S;")
con.Execute ("DELETE FROM LM_Poladata_I;")
con.Execute ("DELETE FROM LM_Poladata_R;")
con.Execute ("DELETE FROM Prediksi_Hasil;")
con.Execute ("DELETE FROM Prediksi_Normalisasi;")
con.Execute ("DELETE FROM Prediksi_Poladatalatih_S;")
con.Execute ("DELETE FROM Prediksi_Poladatalatih_I;")
con.Execute ("DELETE FROM Prediksi_Poladatalatih_R;")
con.Execute ("DELETE FROM Prediksi_MSE;")
con.Execute ("DELETE FROM Hasil_RK;")
con.Close
Form3.Adodc1.Recordset.Sort = Form3.DataGrid1.Columns(0).DataField
Form3.Adodc1.Refresh
If Combo1 = "Estimasi Parameter" Then
    Form2.Show
    Form2.Text8.SetFocus
    Form2.Command2.Enabled = False
    Form1.Hide
Else
    If Combo1 = "Identifikasi Model dan Prediksi" Then
        If Form3.Adodc1.Recordset.RecordCount > 1 Then
            Form3.Show
            Form3.Adodc1.Recordset.MoveFirst
            Form3.Adodc1.Recordset.MoveNext
            Form3.Text11.SetFocus
            Form1.Hide
        Else
            MsgBox "Tidak ada riwayat estimasi parameter", vbOKOnly, "Peringatan"
            Form2.Show
            Form1.Hide
        End If
    End If
End If
End Sub

Private Sub Form_Load()
    Combo1.AddItem "Estimasi Parameter"
    Combo1.AddItem "Identifikasi Model dan Prediksi"
End Sub
