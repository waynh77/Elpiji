VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form F24_ByrJual 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BAYAR PENJUALAN"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ClipControls    =   0   'False
   Icon            =   "F24_ByrJual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "F24_ByrJual.frx":3482
      Height          =   4335
      Left            =   120
      OleObjectBlob   =   "F24_ByrJual.frx":3496
      TabIndex        =   8
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BATAL"
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "F24_ByrJual.frx":3E69
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51773441
      CurrentDate     =   39715
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KETERANGAN"
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
      TabIndex        =   7
      Top             =   840
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH BAYAR"
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
      Index           =   6
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2025
   End
End
Attribute VB_Name = "F24_ByrJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Dim tgl_jual As Date
    Dim noMember As String
    Dim kodeproduk As String
    Dim stsproduk As String
    Dim qty As Single
    Dim jmljual As Double
    Dim namaspl As String
    Dim namaprd As String
    If Text1 <> "" Then
        With F11_TBayar
            tgl_jual = .DTPicker1
            noMember = .ComboBox1
            kodeproduk = .ComboBox2
            namaspl = .Text1
            namaprd = .Text2
            qty = .Text3
            stsproduk = .Text10
            jmljual = Format(.Text5, "###")
            .Data4.Recordset.Edit
            .Data4.Recordset!frek_bayar = .Data4.Recordset!frek_bayar + 1
            .Data4.Recordset!jumlah_bayar = .Data4.Recordset!jumlah_bayar + Text1
            If .Data4.Recordset!jumlah_bayar >= .Data4.Recordset!qty_jual * .Data4.Recordset!harga_satuan Then
                .Data4.Recordset!status_bayar = True
            End If
            .Data4.Recordset.Update
        End With
        With Data1.Recordset
            .AddNew
            !tgl_byr = DTPicker1
            !jml_byr = Text1
            !keterangan = Text2
            !tgl_jual = tgl_jual
            !no_member = noMember
            !nama_member = nama_spl
            !kode_produk = kodeproduk
            !qty = qty
            !nama_produk = namaprd
            !jml_jual = jmljual
            .Update
        End With
        Unload Me
    Else
        MsgBox "Data belum lengkap", vbCritical, "Validasi Input"
    End If
Case 1
    Unload Me
End Select
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Call db_byrJUAL
DTPicker1 = Date
Text1 = ""
Text2 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
F11_TBayar.Enabled = True
F11_TBayar.Show
F11_TBayar.isi1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub
