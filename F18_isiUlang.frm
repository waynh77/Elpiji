VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form F18_isiUlang 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Isi Ulang Tabung"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14175
   ClipControls    =   0   'False
   Icon            =   "F18_isiUlang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   14175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   2880
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
            Picture         =   "F18_isiUlang.frx":3482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F18_isiUlang.frx":4514
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F18_isiUlang.frx":55A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F18_isiUlang.frx":6638
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "KELUAR"
      DownPicture     =   "F18_isiUlang.frx":76CA
      Height          =   735
      Index           =   2
      Left            =   12240
      MouseIcon       =   "F18_isiUlang.frx":7E34
      MousePointer    =   99  'Custom
      Picture         =   "F18_isiUlang.frx":813E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "SUPPLIER BARU"
      DownPicture     =   "F18_isiUlang.frx":91C0
      Height          =   735
      Index           =   1
      Left            =   10200
      MouseIcon       =   "F18_isiUlang.frx":992A
      MousePointer    =   99  'Custom
      Picture         =   "F18_isiUlang.frx":9C34
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "ISI ULANG"
      DownPicture     =   "F18_isiUlang.frx":ACB6
      Height          =   735
      Index           =   0
      Left            =   8160
      MouseIcon       =   "F18_isiUlang.frx":B420
      MousePointer    =   99  'Custom
      Picture         =   "F18_isiUlang.frx":B72A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   11160
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Format          =   57212929
      CurrentDate     =   39689
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   12240
      TabIndex        =   6
      Text            =   "Text7"
      Top             =   2400
      Width           =   1665
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   10200
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   2400
      Width           =   1665
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   8160
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   2400
      Width           =   1665
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   9840
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   9840
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   9840
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080FFFF&
      Height          =   285
      Left            =   9840
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   360
      Width           =   4095
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
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3960
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6165
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8160
      X2              =   13920
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HRG TABUNG @"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   8
      Left            =   8160
      TabIndex        =   21
      Top             =   2040
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Due Date"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   6
      Left            =   9360
      TabIndex        =   20
      Top             =   3000
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QTY"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   9
      Left            =   12240
      TabIndex        =   19
      Top             =   2040
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HARGA ISI @"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   7
      Left            =   10200
      TabIndex        =   18
      Top             =   2040
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Stok Kosong"
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
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   6090
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA SUPPLIER"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   4
      Left            =   8160
      TabIndex        =   15
      Top             =   1080
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE SUPPLIER"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   3
      Left            =   8160
      TabIndex        =   14
      Top             =   1440
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA PRODUK"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   2
      Left            =   8160
      TabIndex        =   13
      Top             =   360
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE PRODUK"
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   1
      Left            =   8160
      TabIndex        =   12
      Top             =   720
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PILIH PRODUK YANG AKAN DI ISI ULANG :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3885
   End
End
Attribute VB_Name = "F18_isiUlang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kode_prd As String
Dim kode_spl As String

Private Sub Combo1_Change()
If Command1(0).Caption = "SIMPAN" Then
    kode_sup2
    Text5 = kode_spl
End If
End Sub

Private Sub Combo1_Click()
If Command1(0).Caption = "SIMPAN" Then
    kode_sup2
    Text5 = kode_spl
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "SIMPAN" Then
        simpan
    Else
        If Text2 <> "" Then
            cmd_simpan
            buka
            isi_cmb1
            nama_supplier
        Else
            MsgBox "Silahkan pilih tabung yang akan diisi terlebih dahulu...", vbInformation, "Validasi Data"
        End If
    End If
Case 1
    If Command1(1).Caption = "BATAL" Then
        cmd_awal
        tutup
        isi
    Else
        Me.Enabled = False
        ctl3 = True
        F04_Supplier.Show
    End If
Case 2
    Unload Me
End Select
End Sub

Sub simpan()
Dim jml As Single
If Text4 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "Data belum lengkap...", vbInformation, "Validasi Input"
    If Text4 = "" Then
        Text4.SetFocus
    ElseIf Text6 = "" Then
        Text6.SetFocus
    Else
        Text7.SetFocus
    End If
Else
    With Data1.Recordset
        jml = !jumlah - Val(Text7)
        If Val(Text7) <= !qty_kosong Then
            If jml <> 0 Then
                .Edit
                    !qty_kosong = !qty_kosong - Val(Text7)
                    !jumlah = !jumlah - Val(Text7)
                .Update
            Else
                .Delete
                Data1.Refresh
            End If
            update_beli
            .AddNew
                !kode_produk = Text3
                !tanggal_masuk = Date
                !kode_supplier = Text5
                !jumlah = Val(Text7)
                !harga_tabung = Val(Text4)
                !due_date = DTPicker1
                !Status = True
                !qty_isi = Val(Text7)
                !qty_kosong = 0
                !jam = Time
                !harga_isi = Val(Text6)
            .Update
            cmd_awal
            tutup
            isi_list
            isi
        Else
            MsgBox "Input qty terlalu banyak...", vbCritical, "validsi Input"
            Text7.SetFocus
        End If
    End With
End If
End Sub

Sub isi_cmb1()
Data3.Refresh
Combo1.Clear
With Data3.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Combo1.AddItem !nama_supplier
        .MoveNext
    Loop
End If
End With
End Sub

Private Sub Form_Activate()
Data4.RecordSource = "select tanggal_masuk,nama_Produk,nama_supplier,Qty_kosong from stok,produk,supplier where stok.kode_produk=produk.kode_produk and stok.kode_supplier=supplier.kode_supplier and qty_kosong>0 order by stok.kode_produk asc , tanggal_masuk desc"
Data4.Refresh
isi_list
Command1(0).SetFocus
End Sub

Private Sub Form_Load()
Call db_isiUlang
kosong
tutup
cmd_awal
DTPicker1 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
F08_TStok.Enabled = True
F08_TStok.Show
End Sub

Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Combo1.Clear
End Sub

Sub isi_list()
Dim ttl As Single
Dim head As ColumnHeader
Dim dtl As ListItem
ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear
Data4.Refresh
Set head = ListView1.ColumnHeaders.Add(, , "Tgl Masuk", ListView1.Width / 6)
Set head = ListView1.ColumnHeaders.Add(, , "Nama Produk", ListView1.Width / (6 / 2))
Set head = ListView1.ColumnHeaders.Add(, , "Nama Supplier", ListView1.Width / (6 / 2))
Set head = ListView1.ColumnHeaders.Add(, , "Jumlah Kosong", ListView1.Width / 6, 2)
ListView1.View = lvwReport
With Data4.Recordset
ttl = 0
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set dtl = ListView1.ListItems.Add(, , !tanggal_masuk)
        dtl.SubItems(1) = !nama_produk
        dtl.SubItems(2) = !nama_supplier
        dtl.SubItems(3) = !qty_kosong
        ttl = ttl + !qty_kosong
        .MoveNext
    Loop
End If
Text1 = ttl
End With
Data4.Refresh
End Sub

Sub tutup()
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Combo1.Enabled = False
DTPicker1.Enabled = False
End Sub

Sub buka()
Combo1.Enabled = True
Text4.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text4 = Format(Text4, "###")
Text6 = Format(Text6, "###")
DTPicker1.Enabled = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
kode_pro
kode_sup
cek = False
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until cek = True Or .EOF
            If !kode_produk = kode_prd And !kode_supplier = kode_spl And !tanggal_masuk = ListView1.SelectedItem.Text And !qty_kosong = ListView1.SelectedItem.ListSubItems(3).Text Then
                cek = True
                .MovePrevious
            End If
            .MoveNext
        Loop
        isi
    End If
End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Sub kode_pro()
Data2.Refresh
With Data2.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !nama_produk = ListView1.SelectedItem.ListSubItems(1).Text Then
                kode_prd = !kode_produk
                .MoveLast
            End If
            .MoveNext
        Loop
    End If
End With
End Sub

Sub kode_sup()
Data3.Refresh
With Data3.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !nama_supplier = ListView1.SelectedItem.ListSubItems(2).Text Then
                kode_spl = !kode_supplier
                .MoveLast
            End If
            .MoveNext
        Loop
    End If
End With
End Sub

Sub kode_sup2()
Data3.Refresh
With Data3.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !nama_supplier = Combo1 Then
                kode_spl = !kode_supplier
                .MoveLast
            End If
            .MoveNext
        Loop
    End If
End With
End Sub

Sub isi()
Dim cari As Boolean
With Data1.Recordset
If Not .BOF Then
    Text3 = kode_prd
    nama_produk
    Text5 = kode_spl
    nama_supplier
    Text4 = Format(!harga_tabung, "###,###.00")
    Text6 = Format(!harga_isi, "###,###.00")
    Text7 = !qty_kosong
End If
End With
End Sub

Sub nama_produk()
Data2.Refresh
With Data2.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !kode_produk = Text3 Then
                Text2 = !nama_produk
                .MoveLast
            End If
            .MoveNext
        Loop
    End If
End With
End Sub

Sub nama_supplier()
Data3.Refresh
With Data3.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !kode_supplier = Text5 Then
                Combo1 = !nama_supplier
                .MoveLast
            End If
            .MoveNext
        Loop
    End If
End With
End Sub

Sub cmd_awal()
Command1(0).Picture = ImageList1.ListImages(1).Picture
Command1(1).Picture = ImageList1.ListImages(2).Picture
Command1(0).Caption = "ISI ULANG"
Command1(1).Caption = "SUPPLIER BARU"
Command1(2).Visible = True
ListView1.Enabled = True
End Sub

Sub cmd_simpan()
Command1(0).Picture = ImageList1.ListImages(3).Picture
Command1(1).Picture = ImageList1.ListImages(4).Picture
Command1(0).Caption = "SIMPAN"
Command1(1).Caption = "BATAL"
Command1(2).Visible = False
ListView1.Enabled = False
End Sub

Sub update_beli()
Data5.Refresh
With Data5.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
    If !tgl_beli = Data1.Recordset!tanggal_masuk And !kode_supplier = Data1.Recordset!kode_supplier And !kode_produk = Data1.Recordset!kode_produk And !qty_beli = Data1.Recordset!jumlah And !due_date = Data1.Recordset!due_date Then
            .AddNew
            !tgl_beli = Date
            !kode_supplier = Text5
            !nama_supplier = Combo1
            !qty_beli = Text7
            !harga_satuan = Val(Text4) + Val(Text6)
            !kode_produk = Text3
            !nama_produk = Text2
            !status_produk = "TABUNG+ISI"
            !due_date = DTPicker1
            If DTPicker1 = Date Then
                !jumlah_bayar = !qty_beli * !harga_satuan
                !status_bayar = True
                !frek_bayar = 1
                With Data6.Recordset
                    .AddNew
                    !tgl_byr = Date
                    !jml_byr = Data5.Recordset!jumlah_bayar
                    !keterangan = ""
                    !tgl_beli = Date
                    !kode_supplier = Text5
                    !kode_produk = Text3
                    !status_produk = Data5.Recordset!status_produk
                    !jml_beli = Data5.Recordset!jumlah_bayar
                    !qty = Text7
                    !nama_supplier = Combo1
                    !nama_produk = Text2
                    .Update
                End With
                
                'update kas
                With Data7.Recordset
                    .AddNew
                    !jenis_kas = "PENGELUARAN"
                    !tgl_kas = Date
                    !jam_kas = Time
                    !nama_kas = "Isi Ulang " & Text2
                    !keterangan = Data6.Recordset!status_produk & " qty = " & Text7
                    !nominal = Data6.Recordset!jml_byr
                    .Update
                End With

            Else
                !jumlah_bayar = 0
                !status_bayar = False
                !frek_bayar = 0
                'isi remainder
                With Data8.Recordset
                    .AddNew
                    Call remain_auto
                    !nomor = urut_remain
                    !tgl = DTPicker1
                    !waktu = "10:00"
                    !Status = True
                    !keterangan = "Hutang kepada " & Combo1 & " untuk Pembelian " & Text2 & " " & Data5.Recordset!status_produk & " qty = " & Text7
                    .Update
                End With
            End If
            .Update
        End If
    .MoveNext
    Loop
End If
End With
End Sub

