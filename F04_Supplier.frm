VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form F04_Supplier 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Supplier"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   ClipControls    =   0   'False
   Icon            =   "F04_Supplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2280
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "DATABASE SUPPLIER"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "KELUAR"
      DownPicture     =   "F04_Supplier.frx":3482
      Height          =   975
      Index           =   6
      Left            =   9840
      MouseIcon       =   "F04_Supplier.frx":3BEC
      MousePointer    =   99  'Custom
      Picture         =   "F04_Supplier.frx":3EF6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "TAMBAH"
      DownPicture     =   "F04_Supplier.frx":4F78
      Height          =   975
      Index           =   0
      Left            =   5520
      MouseIcon       =   "F04_Supplier.frx":56E2
      MousePointer    =   99  'Custom
      Picture         =   "F04_Supplier.frx":59EC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "EDIT"
      DownPicture     =   "F04_Supplier.frx":6A6E
      Height          =   975
      Index           =   1
      Left            =   7440
      MouseIcon       =   "F04_Supplier.frx":71D8
      MousePointer    =   99  'Custom
      Picture         =   "F04_Supplier.frx":74E2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "HAPUS"
      DownPicture     =   "F04_Supplier.frx":8564
      Height          =   975
      Index           =   2
      Left            =   9360
      MouseIcon       =   "F04_Supplier.frx":8CCE
      MousePointer    =   99  'Custom
      Picture         =   "F04_Supplier.frx":8FD8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CARI"
      DownPicture     =   "F04_Supplier.frx":A05A
      Height          =   975
      Index           =   3
      Left            =   5520
      MouseIcon       =   "F04_Supplier.frx":A7C4
      MousePointer    =   99  'Custom
      Picture         =   "F04_Supplier.frx":AACE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CETAK"
      DownPicture     =   "F04_Supplier.frx":BB50
      Height          =   975
      Index           =   4
      Left            =   6960
      MouseIcon       =   "F04_Supplier.frx":C2BA
      MousePointer    =   99  'Custom
      Picture         =   "F04_Supplier.frx":C5C4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "REFRESH"
      DownPicture     =   "F04_Supplier.frx":D646
      Height          =   975
      Index           =   5
      Left            =   8400
      MouseIcon       =   "F04_Supplier.frx":DDB0
      MousePointer    =   99  'Custom
      Picture         =   "F04_Supplier.frx":E0BA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      Height          =   1005
      Index           =   3
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "F04_Supplier.frx":F13C
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   1680
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
            Picture         =   "F04_Supplier.frx":F142
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F04_Supplier.frx":101D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F04_Supplier.frx":11266
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F04_Supplier.frx":122F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6015
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10610
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
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ALAMAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TELEPON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA SUPPLIER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE SUPPLIER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "F04_Supplier"
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
        supplier_auto
        Text1(1).SetFocus
    Else
        simpan
    End If
Case 1
    If Command1(1).Caption = "EDIT" Then
        If Not Data1.Recordset.BOF And Text1(1) <> "" Then
            tambah = False
            cmd_simpan
            buka
            Text1(1).SetFocus
        Else
            MsgBox "Data Kosong/Belum Dipilih...", vbInformation, "Validasi Data"
        End If
    Else
        If ctl2 = True Or ctl3 = True Then
            Unload Me
        Else
            tutup
            Data1.Refresh
            isi_list
            isi
            cmd_awal
        End If
    End If
Case 2
        With Data1.Recordset
        If Not .BOF And Text1(1) <> "" Then
            x = MsgBox("Apakah anda yakin ingin menghapus Supplier : " & Text1(0) & " - " & Text1(1), vbYesNo, "Hapus Data Supplier " & Text2)
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
    cek = False
    cari = InputBox("Masukan Nama Supplier", "Cari Data", "Nama Supplier")
    If cari <> "Nama Supplier" And cari <> "" Then
    With Data1.Recordset
        If Not .BOF Then
           Data1.Refresh
            .MoveFirst
            Do Until .EOF Or cek = True
                If !nama_supplier Like ("*" & cari & "*") Then
                    cek = True
                End If
                .MoveNext
            Loop
            If cek = False Then
                MsgBox "Data tidak ditemukan", vbOKOnly, "Cari Data"
            Else
                MsgBox "Data ditemukan...", vbOKOnly, "Cari Data"
                .MovePrevious
                Text1(0) = !kode_supplier
                Text1(1) = !nama_supplier
                Text1(2) = !telp_supplier
                Text1(3) = !alamat_supplier
            End If
        End If
    End With
    End If
Case 4
        CrystalReport1.ReportFileName = App.Path & "\Data Supplier.rpt"
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1
Case 5
    Data1.Refresh
    isi
Case 6
    Unload Me
End Select
End Sub

Sub simpan()
If Text1(1) = "" Then
    MsgBox "Data belum lengkap...", vbCritical, "Validasi Input"
    Text1(1).SetFocus
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
            End With
        If ctl2 = True Or ctl3 = True Then
            Unload Me
        End If
        End If
    Else
        Data1.Recordset.Edit
        transfer
        Data1.Recordset.Update
        tutup
        cmd_awal
        isi_list
    End If
End If
End Sub

Sub cek_data()
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Text1(1) = !nama_supplier Then
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
    !kode_supplier = Text1(0)
    !nama_supplier = Text1(1)
    !telp_supplier = Text1(2)
    !alamat_supplier = Text1(3)
End With
End Sub

Private Sub Form_Activate()
Data1.Refresh
Text1(1).MaxLength = 50
Text1(2).MaxLength = 20
Text1(3).MaxLength = 100
isi_list
isi
If ctl2 = True Or ctl3 = True Then
    Command1_Click (0)
End If
End Sub

Private Sub Form_Load()
Call db_supplier
kosong
tutup
cmd_awal
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ctl2 = False And ctl3 = False Then
    F01_Main.Enabled = True
    F01_Main.Show
ElseIf ctl2 = True Then
    F16_isiStok.Enabled = True
    F16_isiStok.Show
    ctl2 = False
Else
    F18_isiUlang.Enabled = True
    F18_isiUlang.Show
    ctl3 = False
End If
End Sub

Sub isi_list()
Dim head As ColumnHeader
Dim dtl As ListItem
ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear
Set head = ListView1.ColumnHeaders.Add(, , "Kode", ListView1.Width / 6)
Set head = ListView1.ColumnHeaders.Add(, , "Nama", ListView1.Width / 6)
Set head = ListView1.ColumnHeaders.Add(, , "Telp", ListView1.Width / 6)
Set head = ListView1.ColumnHeaders.Add(, , "Alamat", ListView1.Width / (6 / 4))
ListView1.View = lvwReport
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set dtl = ListView1.ListItems.Add(, , !kode_supplier)
        dtl.SubItems(1) = !nama_supplier
        dtl.SubItems(2) = !telp_supplier
        dtl.SubItems(3) = !alamat_supplier
        .MoveNext
    Loop
End If
End With
Data1.Refresh
End Sub

Sub kosong()
Text1(0) = ""
Text1(1) = ""
Text1(2) = ""
Text1(3) = ""
End Sub

Sub isi()
'If Command1(0).Caption <> "SIMPAN" Then
With Data1.Recordset
    If Not .BOF Then
        Text1(0) = !kode_supplier
        Text1(1) = !nama_supplier
        Text1(2) = !telp_supplier
        Text1(3) = !alamat_supplier
    End If
End With
'End If
End Sub

Sub tutup()
Text1(0).Enabled = False
Text1(1).Enabled = False
Text1(2).Enabled = False
Text1(3).Enabled = False
End Sub

Sub buka()
Text1(1).Enabled = True
Text1(2).Enabled = True
Text1(3).Enabled = True
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
ListView1.Enabled = True
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
ListView1.Enabled = False
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
cari_data
End Sub

Sub cari_data()
cek = False
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until !kode_supplier = ListView1.SelectedItem.Text Or .EOF
            If !kode_supplier = ListView1.SelectedItem.Text Then
                cek = True
            End If
            .MoveNext
        Loop
        isi
    End If
End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 1
    If KeyAscii = 13 Then
        Text1(2).SetFocus
    End If
Case 2
    If KeyAscii = 13 Then
        Text1(3).SetFocus
    End If
End Select
End Sub
