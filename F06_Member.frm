VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form F06_Member 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Member"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   ClipControls    =   0   'False
   Icon            =   "F06_Member.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1800
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "LOKASI BARU"
      DownPicture     =   "F06_Member.frx":3482
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      MouseIcon       =   "F06_Member.frx":414C
      MousePointer    =   99  'Custom
      Picture         =   "F06_Member.frx":4456
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Height          =   285
      Index           =   5
      Left            =   7200
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      Height          =   1365
      Index           =   4
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "F06_Member.frx":5120
      Top             =   480
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   7200
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   3735
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
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      Height          =   1005
      Index           =   3
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "F06_Member.frx":5126
      Top             =   1200
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "REFRESH"
      DownPicture     =   "F06_Member.frx":512C
      Height          =   735
      Index           =   5
      Left            =   7920
      MouseIcon       =   "F06_Member.frx":5896
      MousePointer    =   99  'Custom
      Picture         =   "F06_Member.frx":5BA0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CETAK"
      DownPicture     =   "F06_Member.frx":6C22
      Height          =   735
      Index           =   4
      Left            =   6360
      MouseIcon       =   "F06_Member.frx":738C
      MousePointer    =   99  'Custom
      Picture         =   "F06_Member.frx":7696
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CARI"
      DownPicture     =   "F06_Member.frx":8718
      Height          =   735
      Index           =   3
      Left            =   4800
      MouseIcon       =   "F06_Member.frx":8E82
      MousePointer    =   99  'Custom
      Picture         =   "F06_Member.frx":918C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "HAPUS"
      DownPicture     =   "F06_Member.frx":A20E
      Height          =   735
      Index           =   2
      Left            =   3240
      MouseIcon       =   "F06_Member.frx":A978
      MousePointer    =   99  'Custom
      Picture         =   "F06_Member.frx":AC82
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "EDIT"
      DownPicture     =   "F06_Member.frx":BD04
      Height          =   735
      Index           =   1
      Left            =   1680
      MouseIcon       =   "F06_Member.frx":C46E
      MousePointer    =   99  'Custom
      Picture         =   "F06_Member.frx":C778
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "TAMBAH"
      DownPicture     =   "F06_Member.frx":D7FA
      Height          =   735
      Index           =   0
      Left            =   120
      MouseIcon       =   "F06_Member.frx":DF64
      MousePointer    =   99  'Custom
      Picture         =   "F06_Member.frx":E26E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "KELUAR"
      DownPicture     =   "F06_Member.frx":F2F0
      Height          =   735
      Index           =   6
      Left            =   9480
      MouseIcon       =   "F06_Member.frx":FA5A
      MousePointer    =   99  'Custom
      Picture         =   "F06_Member.frx":FD64
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
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
            Picture         =   "F06_Member.frx":10DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F06_Member.frx":11E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F06_Member.frx":12F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F06_Member.frx":13F9C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5175
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9128
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
      Caption         =   "BIAYA KIRIM"
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
      Index           =   6
      Left            =   5520
      TabIndex        =   22
      Top             =   1920
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA LOKASI"
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
      Index           =   5
      Left            =   5520
      TabIndex        =   21
      Top             =   480
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE LOKASI"
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
      Index           =   4
      Left            =   5520
      TabIndex        =   20
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NO MEMBER"
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
      TabIndex        =   19
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA MEMBER"
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
      TabIndex        =   18
      Top             =   480
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
      TabIndex        =   17
      Top             =   840
      Width           =   1665
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
      TabIndex        =   16
      Top             =   1200
      Width           =   1665
   End
End
Attribute VB_Name = "F06_Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim cek As Boolean

Private Sub Combo1_Change()
isi_lokasi
End Sub

Private Sub Combo1_Click()
isi_lokasi
Command1(0).SetFocus
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
    If KeyAscii = 13 Then
        Command1(0).SetFocus
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        tambah = True
        cmd_simpan
        buka
        kosong
        member_auto
        isi_cmb1
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
        tutup
        Data1.Refresh
        isi
        isi_list
        cmd_awal
        If ctl4 = True Then
            Unload Me
            F09_Nota.isi_cmb1
        End If
    End If
Case 2
    With Data1.Recordset
        If Not .BOF And Text1(1) <> "" Then
            x = MsgBox("Apakah anda yakin ingin menghapus Member : " & Text1(0) & " - " & Text1(1), vbYesNo, "Hapus Data Member " & Text2)
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
    cari = InputBox("Masukan Nama Member", "Cari Data", "Nama Member")
    If cari <> "Nama Member" And cari <> "" Then
    With Data1.Recordset
        If Not .BOF Then
           Data1.Refresh
            .MoveFirst
            Do Until .EOF Or cek = True
                If !nama_member Like ("*" & cari & "*") Then
                    cek = True
                End If
                .MoveNext
            Loop
            If cek = False Then
                MsgBox "Data tidak ditemukan", vbOKOnly, "Cari Data"
            Else
                MsgBox "Data ditemukan...", vbOKOnly, "Cari Data"
                .MovePrevious
                Text1(0) = !no_member
                Text1(1) = !nama_member
                Text1(2) = !telp_member
                Text1(3) = !alamat_member
                Combo1 = !kode_lokasi
            End If
        End If
    End With
    End If
Case 4
        CrystalReport1.ReportFileName = App.Path & "\Data Member (Anggota).rpt"
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
If Text1(1) = "" Or Combo1 = "" Then
    MsgBox "Data belum lengkap...", vbCritical, "Validasi Input"
    If Text1(1) = "" Then
        Text1(1).SetFocus
    Else
        Combo1.SetFocus
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
            End With
        End If
        If ctl4 = True Then
            Unload Me
            F09_Nota.isi_cmb1
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
        If Text1(1) = !nama_member Then
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
    !no_member = Text1(0)
    !nama_member = Text1(1)
    !telp_member = Text1(2)
    !alamat_member = Text1(3)
    !kode_lokasi = Combo1
End With
End Sub

Private Sub Command2_Click()
Me.Enabled = False
F17_Lokasi.Show
End Sub

Private Sub Form_Activate()
Data1.Refresh
Text1(1).MaxLength = 50
Text1(2).MaxLength = 20
Text1(3).MaxLength = 100
isi_cmb1
isi_list
isi
If ctl4 = True Then
    Command1_Click (0)
End If
End Sub

Private Sub Form_Load()
Call db_member
kosong
tutup
cmd_awal
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ctl4 = True Then
    F09_Nota.Enabled = True
    F09_Nota.Show
    ctl4 = False
Else
    F01_Main.Enabled = True
    F01_Main.Show
End If
End Sub

Sub isi_list()
Dim head As ColumnHeader
Dim dtl As ListItem
Dim gbr As ListImage
ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear
Set head = ListView1.ColumnHeaders.Add(, , "NO", ListView1.Width / 9)
Set head = ListView1.ColumnHeaders.Add(, , "Nama", ListView1.Width / 9)
Set head = ListView1.ColumnHeaders.Add(, , "Telp", ListView1.Width / 9)
Set head = ListView1.ColumnHeaders.Add(, , "Alamat", ListView1.Width / (9 / 2))
Set head = ListView1.ColumnHeaders.Add(, , "Kode Lokasi", ListView1.Width / 9)
Set head = ListView1.ColumnHeaders.Add(, , "Nama Lokasi", ListView1.Width / (9 / 2))
Set head = ListView1.ColumnHeaders.Add(, , "Biaya Kirim", ListView1.Width / 9)
ListView1.View = lvwReport
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set dtl = ListView1.ListItems.Add(, , !no_member)
        dtl.SubItems(1) = !nama_member
        dtl.SubItems(2) = !telp_member
        dtl.SubItems(3) = !alamat_member
        dtl.SubItems(4) = !kode_lokasi
        Data2.Refresh
        With Data2.Recordset
            If Not .BOF Then
                .MoveFirst
                Do While Not .EOF
                    If !kode_lokasi = Data1.Recordset!kode_lokasi Then
                        dtl.SubItems(5) = !nama_lokasi
                        dtl.SubItems(6) = Format(!biaya_kirim, "###,###.00")
                        .MoveLast
                    End If
                    .MoveNext
                Loop
            End If
        End With
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
Text1(4) = ""
Text1(5) = ""
End Sub

Sub isi()
'If Command1(0).Caption <> "SIMPAN" Then
With Data1.Recordset
    If Not .BOF Then
        Text1(0) = !no_member
        Text1(1) = !nama_member
        Text1(2) = !telp_member
        Text1(3) = !alamat_member
        Combo1 = !kode_lokasi
    End If
End With
'End If
End Sub

Sub tutup()
Text1(0).Enabled = False
Text1(1).Enabled = False
Text1(2).Enabled = False
Text1(3).Enabled = False
Text1(4).Enabled = False
Text1(5).Enabled = False
Combo1.Enabled = False
End Sub

Sub buka()
Text1(1).Enabled = True
Text1(2).Enabled = True
Text1(3).Enabled = True
Combo1.Enabled = True
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
Command2.Visible = True
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
Command2.Visible = False
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
        Do Until !no_member = ListView1.SelectedItem.Text Or .EOF
            If !no_member = ListView1.SelectedItem.Text Then
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
Case 3
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
End Select
End Sub

Sub isi_cmb1()
Combo1.Clear
Data2.Refresh
With Data2.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo1.AddItem !kode_lokasi
            .MoveNext
        Loop
        Combo1.ListIndex = 0
    End If
End With
End Sub

Sub isi_lokasi()
Data2.Refresh
If Not Data2.Recordset.BOF Then
    With Data2.Recordset
        .MoveFirst
        Do While Not .EOF
            If !kode_lokasi = Combo1 Then
                Text1(4) = !nama_lokasi
                Text1(5) = Format(!biaya_kirim, "###,###.00")
                .MoveLast
            End If
            .MoveNext
        Loop
    End With
End If
End Sub
