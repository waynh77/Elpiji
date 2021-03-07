VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form F08_TStok 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stok Persediaan"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   ClipControls    =   0   'False
   Icon            =   "F08_TStok.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6600
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton m 
      BackColor       =   &H0080FF80&
      Caption         =   "HAPUS"
      DownPicture     =   "F08_TStok.frx":3482
      Height          =   975
      Index           =   2
      Left            =   2760
      MouseIcon       =   "F08_TStok.frx":3BEC
      MousePointer    =   99  'Custom
      Picture         =   "F08_TStok.frx":3EF6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "REFRESH"
      DownPicture     =   "F08_TStok.frx":4F78
      Height          =   735
      Index           =   5
      Left            =   3840
      MouseIcon       =   "F08_TStok.frx":56E2
      MousePointer    =   99  'Custom
      Picture         =   "F08_TStok.frx":59EC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CETAK"
      DownPicture     =   "F08_TStok.frx":6A6E
      Height          =   735
      Index           =   4
      Left            =   8160
      MouseIcon       =   "F08_TStok.frx":71D8
      MousePointer    =   99  'Custom
      Picture         =   "F08_TStok.frx":74E2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CARI"
      DownPicture     =   "F08_TStok.frx":8564
      Height          =   735
      Index           =   3
      Left            =   3840
      MouseIcon       =   "F08_TStok.frx":8CCE
      MousePointer    =   99  'Custom
      Picture         =   "F08_TStok.frx":8FD8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "ISI ULANG"
      DownPicture     =   "F08_TStok.frx":A05A
      Height          =   735
      Index           =   1
      Left            =   6720
      MouseIcon       =   "F08_TStok.frx":A7C4
      MousePointer    =   99  'Custom
      Picture         =   "F08_TStok.frx":AACE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "TAMBAH/EDIT"
      DownPicture     =   "F08_TStok.frx":BB50
      Height          =   735
      Index           =   0
      Left            =   5280
      MouseIcon       =   "F08_TStok.frx":C2BA
      MousePointer    =   99  'Custom
      Picture         =   "F08_TStok.frx":C5C4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "KELUAR"
      DownPicture     =   "F08_TStok.frx":D646
      Height          =   735
      Index           =   6
      Left            =   9600
      MouseIcon       =   "F08_TStok.frx":DDB0
      MousePointer    =   99  'Custom
      Picture         =   "F08_TStok.frx":E0BA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   4080
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
            Picture         =   "F08_TStok.frx":F13C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F08_TStok.frx":101CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F08_TStok.frx":11260
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F08_TStok.frx":122F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1335
      Index           =   1
      Left            =   5640
      TabIndex        =   1
      Top             =   7080
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2355
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   5880
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5655
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9975
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jumlah Stok Kosong"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   5880
      Width           =   3450
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jumlah Stok Isi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   6480
      Width           =   3450
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Stok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   9
      Top             =   5880
      Width           =   4050
   End
End
Attribute VB_Name = "F08_TStok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Me.Enabled = False
    F16_isiStok.Show
Case 1
    Me.Enabled = False
    F18_isiUlang.Show
Case 2
Case 3
Case 4
        CrystalReport1.ReportFileName = App.Path & "\Laporan Stok.rpt"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
Case 5
Case 6
    Unload Me
End Select
End Sub

Private Sub Form_Activate()
isi_list1
isi_total
End Sub

Private Sub Form_Load()
Call db_stok
kosong
End Sub

Sub kosong()
ListView1(0).ListItems.Clear
ListView1(1).ListItems.Clear
End Sub

Sub isi_total()
Dim ttl As Single
Data3.Refresh
ttl = 0
If Not Data3.Recordset.BOF Then
    With Data3.Recordset
        .MoveFirst
        Do While Not .EOF
            ttl = ttl + !jumlah
            .MoveNext
        Loop
    End With
End If
Text1 = ttl
End Sub

Sub isi_list1()
Dim head As ColumnHeader
Dim dtl As ListItem
Dim gbr As ListImage
Dim ttl As Single
Dim tkosong As Single
Dim tisi As Single
Dim kode As String
Dim kosong As Single
Dim isi As Single
ListView1(0).ColumnHeaders.Clear
ListView1(0).ListItems.Clear
'Data3.RecordSource = "select * from stok,produk where stok.kode_produk=produk.kode_produk order by kode_produk,status asc"
'Data3.Refresh
Set head = ListView1(0).ColumnHeaders.Add(, , "Kode Produk", ListView1(0).Width / 6)
Set head = ListView1(0).ColumnHeaders.Add(, , "Nama Produk", ListView1(0).Width / ((6 / 2)) - 100)
Set head = ListView1(0).ColumnHeaders.Add(, , "Tabung Kosong", ListView1(0).Width / 6, 2)
Set head = ListView1(0).ColumnHeaders.Add(, , "Tabung Isi", ListView1(0).Width / 6, 2)
Set head = ListView1(0).ColumnHeaders.Add(, , "Jumlah Stok", ListView1(0).Width / 6, 2)
ListView1(0).View = lvwReport
Data2.Refresh
With Data2.Recordset
tkosong = 0
tisi = 0
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set dtl = ListView1(0).ListItems.Add(, , !kode_produk)
        dtl.SubItems(1) = !nama_produk
        kode = !kode_produk
        ttl = 0
        Data3.Refresh
        With Data3.Recordset
        If Not .BOF Then
            .MoveFirst
            isi = 0
            kosong = 0
            Do While Not .EOF
                If kode = !kode_produk Then
                    isi = isi + !qty_isi
                    kosong = kosong + !qty_kosong
                    ttl = ttl + !jumlah
                    tisi = tisi + !qty_isi
                    tkosong = tkosong + !qty_kosong
                End If
                .MoveNext
            Loop
        End If
        End With
        dtl.SubItems(2) = kosong
        dtl.SubItems(3) = isi
        dtl.SubItems(4) = ttl
'        dtl.SubItems(3) = Rkanan(!harga_kosong, "###,###.00")
'        dtl.SubItems(4) = Rkanan(!harga_isi, "###,###.00")
'        dtl.SubItems(5) = Rkanan(!harga_kosong + !harga_isi, "###,###.00")
        .MoveNext
    Loop
End If
Text2 = tisi
Text3 = tkosong
End With
Data1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
F01_Main.Enabled = True
F01_Main.Show
End Sub

Sub isi_list2()
Dim head As ColumnHeader
Dim dtl As ListItem
Dim gbr As ListImage
Dim kosong As Single
Dim isi As Single
Data4.RecordSource = "select * from stok,produk where stok.kode_produk=produk.kode_produk and kode_produk ='" & ListView1(0).SelectedItem.Text & "' order by status desc"
Data4.Refresh
ListView1(1).ColumnHeaders.Clear
ListView1(1).ListItems.Clear
Set head = ListView1(1).ColumnHeaders.Add(, , "Kode Produk", ListView1(1).Width / 6, 2)
Set head = ListView1(1).ColumnHeaders.Add(, , "Nama Supplier", ListView1(1).Width / ((6 / 2)) - 100)
Set head = ListView1(1).ColumnHeaders.Add(, , "Jumlah ", ListView1(1).Width / 6, 2)
Set head = ListView1(1).ColumnHeaders.Add(, , "Harga Beli ", ListView1(1).Width / 6, 1)
ListView1(1).View = lvwReport
With Data4.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Set dtl = ListView1(1).ListItems.Add(, , !tanggal_masuk)
            dtl.SubItems(1) = !kode_supplier
            dtl.SubItems(2) = !nama_supplier
            dtl.SubItems(3) = !jumlah
            dtl.SubItems(4) = Format(!harga_beli, "###,###.00")
            .MoveNext
        Loop
    End If
End With
End Sub

Private Sub ListView1_GotFocus(Index As Integer)
Select Case Index
Case 0
'isi_list2
End Select
End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
Select Case Index
Case 0
'isi_list2
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
