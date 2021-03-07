VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form F13_ArAp 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Hutang Piutang"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   ClipControls    =   0   'False
   Icon            =   "F13_ArAp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BATAL"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   2160
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51314689
      CurrentDate     =   39753
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51314689
      CurrentDate     =   39753
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Text            =   "Combo3"
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   480
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S/D TANGGAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TANGGAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PERIODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS LAPORAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JENIS LAPORAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   21
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2025
   End
End
Attribute VB_Name = "F13_ArAp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_Click()
If Combo3.ListIndex = 0 Then
    Label1(3).Visible = False
    DTPicker2.Visible = False
Else
    Label1(3).Visible = True
    DTPicker2.Visible = True
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Combo1.ListIndex = 0 Then
        CrystalReport1.ReportFileName = App.Path & "\Laporan Piutang (AR).rpt"
    Else
        CrystalReport1.ReportFileName = App.Path & "\Laporan Hutang (AP).rpt"
    End If
    
    'lap piutang
    If Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{penjualan.tgl_jual}= date(" & Format(DTPicker1, "yyyy,m,d") & ")"
    ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{penjualan.tgl_jual}= date(" & Format(DTPicker1, "yyyy,m,d") & ")and {penjualan.status_bayar}= false"
    ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{penjualan.tgl_jual}= date(" & Format(DTPicker1, "yyyy,m,d") & ")and {penjualan.status_bayar}= true"
    ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
        CrystalReport1.SelectionFormula = "{penjualan.tgl_jual}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {penjualan.tgl_jual}<= date(" & Format(DTPicker2, "yyyy,m,d") & ")"
    ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
        CrystalReport1.SelectionFormula = "{penjualan.tgl_jual}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {penjualan.tgl_jual}<= date(" & Format(DTPicker2, "yyyy,m,d") & ") and {penjualan.status_bayar}= false"
    ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
        CrystalReport1.SelectionFormula = "{penjualan.tgl_jual}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {penjualan.tgl_jual}<= date(" & Format(DTPicker2, "yyyy,m,d") & ") and {penjualan.status_bayar}= true"

    'lap hutang
    ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{pembelian.tgl_beli}= date(" & Format(DTPicker1, "yyyy,m,d") & ")"
    ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{pembelian.tgl_beli}= date(" & Format(DTPicker1, "yyyy,m,d") & ")and {pembelian.status_bayar}= false"
    ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{pembelian.tgl_beli}= date(" & Format(DTPicker1, "yyyy,m,d") & ")and {pembelian.status_bayar}= true"
    ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 And Combo3.ListIndex = 1 Then
        CrystalReport1.SelectionFormula = "{pembelian.tgl_beli}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {pembelian.tgl_beli}<= date(" & Format(DTPicker2, "yyyy,m,d") & ")"
    ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 And Combo3.ListIndex = 1 Then
        CrystalReport1.SelectionFormula = "{pembelian.tgl_beli}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {pembelian.tgl_beli}<= date(" & Format(DTPicker2, "yyyy,m,d") & ") and {pembelian.status_bayar}= false"
    ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 2 And Combo3.ListIndex = 1 Then
        CrystalReport1.SelectionFormula = "{pembelian.tgl_beli}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {pembelian.tgl_beli}<= date(" & Format(DTPicker2, "yyyy,m,d") & ") and {pembelian.status_bayar}= true"
    
    End If
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
Case 1
    Unload Me
End Select
End Sub

Private Sub Form_Load()
isi_cmb1
isi_cmb2
isi_cmb3
DTPicker1 = Date
DTPicker2 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
F01_Main.Enabled = True
F01_Main.Show
End Sub

Sub isi_cmb1()
Combo1.Clear
Combo1.AddItem "AR/Piutang"
Combo1.AddItem "AP/Hutang"
Combo1.ListIndex = 0
End Sub

Sub isi_cmb2()
Combo2.Clear
Combo2.AddItem "All/Semua"
Combo2.AddItem "Belum Lunas"
Combo2.AddItem "Lunas"
Combo2.ListIndex = 0
End Sub

Sub isi_cmb3()
Combo3.Clear
Combo3.AddItem "Harian"
Combo3.AddItem "Periodik"
Combo3.ListIndex = 0
End Sub
