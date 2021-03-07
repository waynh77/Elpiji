VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form F01_Main 
   BackColor       =   &H00000000&
   Caption         =   "TOKO EDISON"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   ClipControls    =   0   'False
   Icon            =   "F01_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer5 
      Interval        =   500
      Left            =   12600
      Top             =   1200
   End
   Begin VB.Timer Timer4 
      Interval        =   3000
      Left            =   2760
      Top             =   2400
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   2280
      Top             =   2400
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1080
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":2444
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":2BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":3338
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":3AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":422C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":49A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":5120
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":589A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":6014
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":678E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":6F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":7682
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":7DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":8576
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":8CF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2760
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   2280
      Top             =   2880
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":946A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":A4FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":B58E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":C620
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":D6B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":E744
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":F7D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":10868
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":118FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":1298C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":13A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":14AB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":15B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":16BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":17C66
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F01_Main.frx":18CF8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1217
      ButtonWidth     =   1217
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MsData"
            Object.ToolTipText     =   "Master Database"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Produk"
            Object.ToolTipText     =   "Database Produk"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Supplier"
            Object.ToolTipText     =   "Database Supplier"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Petugas"
            Object.ToolTipText     =   "Database Petugas"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Member"
            Object.ToolTipText     =   "Database Anggota"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Harga"
            Object.ToolTipText     =   "Setting Harga Produk"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stok"
            Object.ToolTipText     =   "Stok Barang Keluar/Masuk"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nota"
            Object.ToolTipText     =   "Buat Nota Penjualan"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "KasKecil"
            Object.ToolTipText     =   "Input Kas Keluar/Masuk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bayar"
            Object.ToolTipText     =   "Pembayaran Cash/Credit"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "L.Jual"
            Object.ToolTipText     =   "Laporan Penjualan"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "AR/AP"
            Object.ToolTipText     =   "Laporan Hutang Piutang Penjualan"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "L.Kas"
            Object.ToolTipText     =   "Laporan Arus Kas"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Calk"
            Object.ToolTipText     =   "Kalkulator"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remain"
            Object.ToolTipText     =   "Pengingat"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Object.ToolTipText     =   "Keluar dari Program....."
            ImageIndex      =   15
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "F01_Main.frx":19D8A
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10335
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21220
            Text            =   "PT Arjuna Wahana Putra - Sub Dealer Gas Elpiji Pertamina / FIFO SYSTEM"
            TextSave        =   "PT Arjuna Wahana Putra - Sub Dealer Gas Elpiji Pertamina / FIFO SYSTEM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "3/1/2009"
            Key             =   "Tanggal Sekarang "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "10:18 PM"
            Key             =   "Jam Sekarang"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remainder"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   13080
      MouseIcon       =   "F01_Main.frx":1A0A4
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tanggal,jan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   13920
      TabIndex        =   8
      Top             =   840
      Width           =   1275
   End
   Begin VB.Image Image2 
      Height          =   960
      Index           =   1
      Left            =   14160
      MouseIcon       =   "F01_Main.frx":1A3AE
      MousePointer    =   99  'Custom
      Picture         =   "F01_Main.frx":1A6B8
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1065
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   $"F01_Main.frx":1C382
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   8520
      Width           =   14040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PT. Arjuna Wahana Putra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   630
      Left            =   150
      TabIndex        =   6
      Top             =   720
      Width           =   6645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOKO ""EDISON"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   2490
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Sub Dealer Gas Elpiji Pertamina"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   3600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WaynhSoft"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   12960
      TabIndex        =   2
      Top             =   9360
      Width           =   2160
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Design by:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13200
      TabIndex        =   1
      Top             =   9120
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   960
      Index           =   0
      Left            =   14160
      MouseIcon       =   "F01_Main.frx":1C413
      MousePointer    =   99  'Custom
      Picture         =   "F01_Main.frx":1C71D
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   6720
      Index           =   0
      Left            =   0
      Picture         =   "F01_Main.frx":1E3E7
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   15480
   End
   Begin VB.Menu db_mnu 
      Caption         =   "&Database"
      Begin VB.Menu msdb_mnu 
         Caption         =   "Master &Database"
      End
      Begin VB.Menu produk_mnu 
         Caption         =   "&Produk"
      End
      Begin VB.Menu spl_mnu 
         Caption         =   "&Supplier"
      End
      Begin VB.Menu petugas_mnu 
         Caption         =   "Pe&tugas"
      End
      Begin VB.Menu member_mnu 
         Caption         =   "&Member"
      End
      Begin VB.Menu harga_mnu 
         Caption         =   "&Harga"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu trans_mnu 
      Caption         =   "&Transaksi"
      Begin VB.Menu Stok_mnu 
         Caption         =   "&Stok"
      End
      Begin VB.Menu nota_mnu 
         Caption         =   "&Nota Penjualan"
      End
      Begin VB.Menu ctkUlang_mnu 
         Caption         =   "&Cetak Ulang Nota"
      End
      Begin VB.Menu kas_mnu 
         Caption         =   "&Kas Kecil"
      End
      Begin VB.Menu bayar_mnu 
         Caption         =   "Pem&bayaran"
      End
   End
   Begin VB.Menu lap_mnu 
      Caption         =   "&Laporan"
      Begin VB.Menu Ljual_mnu 
         Caption         =   "Laporan &Penjualan"
      End
      Begin VB.Menu ArAp_mnu 
         Caption         =   "Laporan &Hutang/Piutang"
      End
      Begin VB.Menu Lkas_mnu 
         Caption         =   "Laporan &Kas"
      End
   End
   Begin VB.Menu tool_mnu 
      Caption         =   "T&ools"
      Begin VB.Menu kalk_mnu 
         Caption         =   "&Kalkulator"
      End
      Begin VB.Menu remain_mnu 
         Caption         =   "&Remainder"
      End
   End
   Begin VB.Menu exit_mnu 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "F01_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lbl As String
Dim lbl2 As String
Dim b, c, d As Byte
Dim e, f As Boolean
Dim a As Boolean

Private Sub ArAp_mnu_Click()
Me.Enabled = False
F13_ArAp.Show
End Sub

Private Sub bayar_mnu_Click()
F11_TBayar.Show
Me.Enabled = False
End Sub

Private Sub ctkUlang_mnu_Click()
Me.Enabled = False
F21_UlangNota.Show
End Sub

Private Sub exit_mnu_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call db_main
cek_remain
End Sub

Private Sub Form_Load()
e = True
f = True
Label2(3).Caption = Format(Date, "d mmmm yyyy") & ", " & Format(Time, "hh:mm:ss")
End Sub

Private Sub Form_Unload(Cancel As Integer)
x = MsgBox("Apakah anda yakin ingin keluar dari program...???", vbYesNo, "EXIT")
If x = vbNo Then
    Cancel = True
Else
    End
End If
End Sub

Private Sub harga_mnu_Click()
Me.Enabled = False
F07_Harga.Show
End Sub

Private Sub Image2_Click(Index As Integer)
Form2.Show
End Sub

Private Sub kalk_mnu_Click()
    AppActivate Shell("calc.exe", 1)
End Sub

Private Sub kas_mnu_Click()
Me.Enabled = False
F10_TKas.Show
End Sub

Private Sub Label3_Click()
REmain_frm.Show
End Sub

Private Sub Ljual_mnu_Click()
    Me.Enabled = False
    F12_LJual.Show
End Sub

Private Sub Lkas_mnu_Click()
Me.Enabled = False
F14_Lkas.Show
End Sub

Private Sub member_mnu_Click()
Me.Enabled = False
F06_Member.Show
End Sub

Private Sub msdb_mnu_Click()
Me.Enabled = False
F02_MsDb.Show
End Sub

Private Sub nota_mnu_Click()
x = MsgBox("Apakah Pembeli adalah Member...???", vbYesNo, "MEMBER/TIDAK")
With F09_Nota
If x = vbNo Then
        .Command3.Visible = True
        .Label1(1).Caption = "NAMA"
        .kosong1
        .Combo4.Visible = True
        .Text1.Enabled = True
        .Text2.Enabled = True
        .Text3.Enabled = True
        .Text5.Enabled = False
        .isi_cmb4
        .isi_cmb2
        .dt_member = ""
Else
    .Command3.Visible = False
    .kosong1
    .kosong2
    .tutup1
    .tutup2
    .Combo4.Visible = False
    .isi_cmb1
    .isi_cmb2
End If
Me.Enabled = False
F09_Nota.Show
End With
End Sub

Private Sub petugas_mnu_Click()
Me.Enabled = False
F05_Petugas.Show
End Sub

Private Sub produk_mnu_Click()
Me.Enabled = False
F03_Produk.Show
End Sub

Private Sub remain_mnu_Click()
Me.Enabled = False
F15_Remain.Show
End Sub

Private Sub spl_mnu_Click()
Me.Enabled = False
F04_Supplier.Show
End Sub

Private Sub Stok_mnu_Click()
Me.Enabled = False
F08_TStok.Show
End Sub

Private Sub Timer1_Timer()
lbl = "Sub Dealer Gas Elpiji Pertamina"
If b < Len(lbl) Then
    b = b + 1
    Label2(0).Caption = Left(lbl, b)
Else
    b = 0
End If
End Sub

Private Sub Timer2_Timer()
lbl2 = "Jalan Lambung Mangkurat No.10 Banjarmasin Indonesia, Telp : (0511)3351050 - 9012828 - 9012288                                                "
x = Len(lbl2)

If c < x + 100 Then
    c = c + 1
    If c >= x Then
        Label2(2).Alignment = 0
        d = d + 1
        Label2(2).Caption = Mid(lbl2, d, 100)
    Else
        Label2(2).Alignment = 1
        Label2(2).Caption = Mid(lbl2, 1, c)
    End If
Else
    c = 1
    d = 1
End If
End Sub

Private Sub Timer3_Timer()
If e = True Then
    Image2(0).Visible = True
    Image2(1).Visible = False
    e = False
Else
    Image2(0).Visible = False
    Image2(1).Visible = True
    e = True
End If
Label2(3).Caption = Format(Date, "d mmmm yyyy") & ", " & Format(Time, "hh:mm:ss")
End Sub

Private Sub Timer4_Timer()
If f = True Then
    Me.Caption = "PT ARJUNA WAHANA PUTRA"
    f = False
Else
    Me.Caption = "TOKO EDISON"
    f = True
End If
End Sub

Private Sub Timer5_Timer()
If a = False Then
'    Label1.Visible = True
    Label3.Caption = "REMAINDER"
    a = True
Else
'    Label1.Visible = False
    Label3.Caption = "Klik disini"
    a = False
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    msdb_mnu_Click
Case 2
    produk_mnu_Click
Case 3
    spl_mnu_Click
Case 4
    petugas_mnu_Click
Case 5
    member_mnu_Click
Case 6
    harga_mnu_Click
Case 7
    Stok_mnu_Click
Case 8
    nota_mnu_Click
Case 9
    kas_mnu_Click
Case 10
    bayar_mnu_Click
Case 11
    Ljual_mnu_Click
Case 12
    ArAp_mnu_Click
Case 13
    Lkas_mnu_Click
Case 14
    kalk_mnu_Click
Case 15
    remain_mnu_Click
Case 16
    Unload Me
End Select
End Sub

Sub cek_remain()
Dim cek As Boolean
cek = False
Call db_main
Data1.RecordSource = "select * from remainder where status=true"
Data1.Refresh
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If !tgl = Date Then
            If Val(Format(Time, "hh.mm")) >= Val(Format(!waktu, "hh.mm")) Then
                cek = True
                .MoveLast
            End If
        ElseIf !tgl < Date Then
            cek = True
            .MoveLast
        End If
        .MoveNext
    Loop
Else
    Timer5.Enabled = False
    Label3.Visible = False
End If
End With
If cek = True Then
    Timer5.Enabled = True
    Label3.Visible = True
Else
    Timer5.Enabled = False
    Label3.Visible = False
End If
Data1.Refresh

End Sub
