VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form F14_Lkas 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Arus Kas"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ClipControls    =   0   'False
   Icon            =   "F14_Lkas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   840
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BATAL"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51314689
      CurrentDate     =   39752
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51314689
      CurrentDate     =   39752
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
      TabIndex        =   7
      Top             =   120
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
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2025
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
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2025
   End
End
Attribute VB_Name = "F14_Lkas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
    Label1(1).Visible = False
    DTPicker2.Visible = False
Else
    Label1(1).Visible = True
    DTPicker2.Visible = True
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Combo1.ListIndex = 0 Then
        CrystalReport1.ReportFileName = App.Path & "\Laporan Arus Kas.rpt"
        CrystalReport1.SelectionFormula = "{kas.tgl_kas}= date(" & Format(DTPicker1, "yyyy,m,d") & ")"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    Else
        CrystalReport1.ReportFileName = App.Path & "\Laporan Arus Kas.rpt"
        CrystalReport1.SelectionFormula = "{kas.tgl_kas}>= date(" & Format(DTPicker1, "yyyy,m,d") & ") and {kas.tgl_kas}<= date(" & Format(DTPicker2, "yyyy,m,d") & ") "
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
    End If
Case 1
    Unload Me
End Select
End Sub

Private Sub Form_Load()
F01_Main.Enabled = False
Combo1.Clear
Combo1.AddItem "Harian"
Combo1.AddItem "Periodik"
Combo1.ListIndex = 0
DTPicker1 = Date
DTPicker2 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
F01_Main.Enabled = True
F01_Main.Show
End Sub

