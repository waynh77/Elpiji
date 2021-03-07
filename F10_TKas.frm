VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form F10_TKas 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Kas"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   ClipControls    =   0   'False
   Icon            =   "F10_TKas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5640
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51576833
      CurrentDate     =   39707
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "KELUAR"
      DownPicture     =   "F10_TKas.frx":3482
      Height          =   1575
      Index           =   6
      Left            =   9120
      MouseIcon       =   "F10_TKas.frx":3BEC
      MousePointer    =   99  'Custom
      Picture         =   "F10_TKas.frx":3EF6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "TAMBAH"
      DownPicture     =   "F10_TKas.frx":4F78
      Height          =   735
      Index           =   0
      Left            =   6240
      MouseIcon       =   "F10_TKas.frx":56E2
      MousePointer    =   99  'Custom
      Picture         =   "F10_TKas.frx":59EC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "EDIT"
      DownPicture     =   "F10_TKas.frx":6A6E
      Height          =   735
      Index           =   1
      Left            =   7200
      MouseIcon       =   "F10_TKas.frx":71D8
      MousePointer    =   99  'Custom
      Picture         =   "F10_TKas.frx":74E2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "HAPUS"
      DownPicture     =   "F10_TKas.frx":8564
      Height          =   735
      Index           =   2
      Left            =   6240
      MouseIcon       =   "F10_TKas.frx":8CCE
      MousePointer    =   99  'Custom
      Picture         =   "F10_TKas.frx":8FD8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CARI"
      DownPicture     =   "F10_TKas.frx":A05A
      Height          =   735
      Index           =   3
      Left            =   7200
      MouseIcon       =   "F10_TKas.frx":A7C4
      MousePointer    =   99  'Custom
      Picture         =   "F10_TKas.frx":AACE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CETAK"
      DownPicture     =   "F10_TKas.frx":BB50
      Height          =   735
      Index           =   4
      Left            =   8160
      MouseIcon       =   "F10_TKas.frx":C2BA
      MousePointer    =   99  'Custom
      Picture         =   "F10_TKas.frx":C5C4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "REFRESH"
      DownPicture     =   "F10_TKas.frx":D646
      Height          =   735
      Index           =   5
      Left            =   8160
      MouseIcon       =   "F10_TKas.frx":DDB0
      MousePointer    =   99  'Custom
      Picture         =   "F10_TKas.frx":E0BA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   735
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "F10_TKas.frx":F13C
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   480
      Width           =   3855
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   8250
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12806
            Text            =   "PT Arjuna Wahana Putra - Sub Dealer Gas Elpiji Pertamina / FIFO SYSTEM"
            TextSave        =   "PT Arjuna Wahana Putra - Sub Dealer Gas Elpiji Pertamina / FIFO SYSTEM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "11/2/2008"
            Key             =   "Tanggal Sekarang "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "12:10 PM"
            Key             =   "Jam Sekarang"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F10_TKas.frx":F142
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F10_TKas.frx":101D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F10_TKas.frx":11266
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F10_TKas.frx":122F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Index           =   1
      Left            =   5160
      TabIndex        =   25
      Top             =   2880
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8281
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648384
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   2880
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8281
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648384
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "TTL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   315
      Index           =   10
      Left            =   8160
      TabIndex        =   23
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "TOTAL PENGELUARAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   315
      Index           =   9
      Left            =   5160
      TabIndex        =   22
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "TTL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   315
      Index           =   8
      Left            =   3120
      TabIndex        =   21
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "TOTAL PENERIMAAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   315
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   19
      Top             =   120
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PENGELUARAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   5
      Left            =   5160
      TabIndex        =   18
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "PENERIMAAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   6120
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMINAL"
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
      TabIndex        =   15
      Top             =   2040
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA TRANSAKSI"
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
      TabIndex        =   13
      Top             =   840
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JENIS TRANSAKSI"
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
      TabIndex        =   12
      Top             =   480
      Width           =   2025
   End
End
Attribute VB_Name = "F10_TKas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim cek As Boolean

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        tambah = True
        cmd_simpan
        buka
        kosong
        isi_cmb1
        Text1.SetFocus
    Else
        simpan
    End If
Case 1
    If Command1(1).Caption = "EDIT" Then
        If Not Data1.Recordset.BOF And Text1 <> "" Then
            tambah = False
            cmd_simpan
            buka
            Text1.SetFocus
            Text3 = Format(Text3, "###")
        Else
            MsgBox "Data Kosong/Belum Dipilih...", vbInformation, "Validasi Data"
        End If
    Else
        tutup
        Data1.Refresh
        isi
        isi_list
        cmd_awal
    End If
Case 2
    With Data1.Recordset
        If Not .BOF And Text1 <> "" Then
            x = MsgBox("Apakah anda yakin ingin menghapus transaksi : " & Text1 & " - " & Format(Text3, "###,###.00"), vbYesNo, "Hapus Data transaksi " & Text1)
            If x = vbYes Then
                .Delete
                .MovePrevious
                If .BOF Then
                    Data1.Refresh
                    kosong
                End If
                isi_list
                isi
            End If
        Else
            MsgBox "Data kosong/belum dipilih...", vbCritical, "Hapus Data"
        End If
    End With
Case 3
Case 4
    Unload Me
    F14_Lkas.Show
Case 5
Case 6
    Unload Me
End Select
End Sub

Sub simpan()
If Text1 = "" Or Text3 = "" Then
    MsgBox "Data belum lengkap...", vbCritical, "Validasi Input"
    If Text1 = "" Then
        Text1.SetFocus
    Else
        Text3.SetFocus
    End If
Else
    If tambah = True Then
        cek = False
        cek_data
        If cek = True Then
            MsgBox "Data Sudah Ada", vbInformation, "Validasi Data"
        Else
            With Data1.Recordset
                .AddNew
                transfer
                .Update
                .MovePrevious
                tutup
                cmd_awal
                isi_list
                isi
            End With
        End If
    Else
        Data1.Recordset.Edit
        transfer
        Data1.Recordset.Update
        tutup
        cmd_awal
        isi_list
        isi
    End If
End If
End Sub

Sub cek_data()
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Text1 = !nama_kas Then
            cek = True
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Sub transfer()
With Data1.Recordset
    !tgl_kas = DTPicker1
    !jenis_kas = Combo1
    !nama_kas = Text1
    !keterangan = Text2
    !nominal = Val(Text3)
    !jam_kas = Time
End With
End Sub

Sub isi()
With Data1.Recordset
    If Not .BOF Then
        Text1 = !nama_kas
        Text2 = !keterangan
        Text3 = Format(!nominal, "###,###.00")
        Combo1 = !jenis_kas
    End If
End With
End Sub

Sub isi_list()
Dim head As ColumnHeader
Dim dtl As ListItem
Dim ttl_masuk As Double
Dim ttl_keluar As Double
ttl_masuk = 0
ttl_keluar = 0
ListView1(0).ColumnHeaders.Clear
ListView1(0).ListItems.Clear
Set head = ListView1(0).ColumnHeaders.Add(, , "NAMA KAS", ListView1(0).Width / 3)
Set head = ListView1(0).ColumnHeaders.Add(, , "KETERANGAN", ListView1(0).Width / 3)
Set head = ListView1(0).ColumnHeaders.Add(, , "NOMINAL", ListView1(0).Width / 3 - 100, 1)
ListView1(0).View = lvwReport
ListView1(1).ColumnHeaders.Clear
ListView1(1).ListItems.Clear
Set head = ListView1(1).ColumnHeaders.Add(, , "NAMA KAS", ListView1(0).Width / 3)
Set head = ListView1(1).ColumnHeaders.Add(, , "KETERANGAN", ListView1(0).Width / 3)
Set head = ListView1(1).ColumnHeaders.Add(, , "NOMINAL", ListView1(0).Width / 3 - 100, 1)
ListView1(1).View = lvwReport
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If !jenis_kas = "PENERIMAAN" Then
            Set dtl = ListView1(0).ListItems.Add(, , !nama_kas)
            dtl.SubItems(1) = !keterangan
            dtl.SubItems(2) = Format(!nominal, "###,###.00")
            ttl_masuk = ttl_masuk + !nominal
        Else
            Set dtl = ListView1(1).ListItems.Add(, , !nama_kas)
            dtl.SubItems(1) = !keterangan
            dtl.SubItems(2) = Format(!nominal, "###,###.00")
            ttl_keluar = ttl_keluar + !nominal
        End If
        .MoveNext
    Loop
End If
End With
Label1(8).Caption = Format(ttl_masuk, "###,###.00")
Label1(10).Caption = Format(ttl_keluar, "###,###.00")
Data1.Refresh
End Sub

Private Sub DTPicker1_Change()
kosong
Data1.RecordSource = "select * from kas where cdate(tgl_kas)='" & DTPicker1 & "'"
Data1.Refresh
isi_list
End Sub

Private Sub DTPicker1_Click()
kosong
Data1.RecordSource = "select * from kas where cdate(tgl_kas)='" & DTPicker1 & "'"
Data1.Refresh
isi_list
End Sub

Private Sub Form_Activate()
Data1.RecordSource = "select * from kas where cdate(tgl_kas)='" & DTPicker1 & "'"
Data1.Refresh
isi_list
End Sub

Private Sub Form_Load()
Call db_TKas
kosong
isi_cmb1
tutup
DTPicker1 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
F01_Main.Enabled = True
F01_Main.Show
End Sub

Sub isi_cmb1()
Combo1.Clear
Combo1.AddItem "PENERIMAAN"
Combo1.AddItem "PENGELUARAN"
Combo1.ListIndex = 0
End Sub

Sub kosong()
Text1 = ""
Text2 = ""
Text3 = 0
End Sub

Sub buka()
Combo1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
DTPicker1.Enabled = True
End Sub

Sub tutup()
'DTPicker1.Enabled = False
Combo1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
End Sub

Sub cmd_awal()
Command1(0).Picture = ImageList1.ListImages(1).Picture
Command1(1).Picture = ImageList1.ListImages(3).Picture
Command1(0).Caption = "TAMBAH"
Command1(1).Caption = "EDIT"
Command1(2).Visible = True
Command1(3).Visible = True
Command1(4).Visible = True
Command1(5).Visible = True
Command1(6).Visible = True
ListView1(0).Enabled = True
ListView1(1).Enabled = True
End Sub

Sub cmd_simpan()
Command1(0).Picture = ImageList1.ListImages(2).Picture
Command1(1).Picture = ImageList1.ListImages(4).Picture
Command1(0).Caption = "SIMPAN"
Command1(1).Caption = "BATAL"
Command1(2).Visible = False
Command1(3).Visible = False
Command1(4).Visible = False
Command1(5).Visible = False
Command1(6).Visible = False
ListView1(0).Enabled = False
ListView1(1).Enabled = False
End Sub

Sub cari_data1()
cek = False
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until !nama_kas = ListView1(0).SelectedItem.Text Or .EOF
            If !nama_kas = ListView1(0).SelectedItem.Text Then
                cek = True
            End If
            .MoveNext
        Loop
        isi
    End If
End With
End Sub

Sub cari_data2()
cek = False
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until !nama_kas = ListView1(1).SelectedItem.Text Or .EOF
            If !nama_kas = ListView1(1).SelectedItem.Text Then
                cek = True
            End If
            .MoveNext
        Loop
        isi
    End If
End With
End Sub

Private Sub ListView1_Click(Index As Integer)
Select Case Index
Case 0
    cari_data1
Case 1
    cari_data2
End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Sub
