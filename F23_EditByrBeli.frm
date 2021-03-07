VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form F23_EditByrBeli 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDIT PEMBAYARAN PERSEDIAAN"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   ClipControls    =   0   'False
   Icon            =   "F23_EditByrBeli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F23_EditByrBeli.frx":3482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F23_EditByrBeli.frx":4E14
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "F23_EditByrBeli.frx":67A6
      Top             =   3360
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4680
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3000
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51904513
      CurrentDate     =   39715
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "KELUAR"
      DownPicture     =   "F23_EditByrBeli.frx":67AC
      Height          =   855
      Index           =   2
      Left            =   5760
      Picture         =   "F23_EditByrBeli.frx":7476
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "HAPUS"
      DownPicture     =   "F23_EditByrBeli.frx":8140
      Height          =   855
      Index           =   1
      Left            =   5760
      Picture         =   "F23_EditByrBeli.frx":8E0A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "EDIT"
      DownPicture     =   "F23_EditByrBeli.frx":9AD4
      Height          =   855
      Index           =   0
      Left            =   5760
      Picture         =   "F23_EditByrBeli.frx":A79E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "F23_EditByrBeli.frx":C120
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "F23_EditByrBeli.frx":C134
      TabIndex        =   1
      Top             =   4080
      Width           =   6735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4895
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   8454143
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
      Top             =   3360
      Width           =   1425
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
      Left            =   3000
      TabIndex        =   6
      Top             =   3000
      Width           =   1665
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
      TabIndex        =   5
      Top             =   3000
      Width           =   1425
   End
End
Attribute VB_Name = "F23_EditByrBeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "EDIT" Then
        If Not Data1.Recordset.BOF Then
            cmd_simpan
            buka
            Text1 = Format(Text1, "###")
            Text1.SetFocus
        End If
    Else
        simpan
    End If
Case 1
    Dim jml As Double
    If Command1(1).Caption = "HAPUS" Then
        If Not Data1.Recordset.BOF Then
            x = MsgBox("Apakah anda yakin menghapus data?", vbYesNo, "Hapus Data")
            If x = vbYes Then
                jml = Data1.Recordset!jml_byr
                With F11_TBayar.Data1.Recordset
                    .Edit
                    !jumlah_bayar = !jumlah_bayar - jml
                    !frek_bayar = !frek_bayar - 1
                    If !jumlah_bayar < !qty_beli * !harga_satuan Then
                        !status_bayar = False
                    Else
                        !status_bayar = True
                    End If
                    .Update
                End With
                Data1.Recordset.Delete
                Data1.Refresh
                isi_list
                isi
            End If
        End If
    Else
        cmd_awal
        tutup
        kosong
        isi
    End If
Case 2
    Unload Me
End Select
End Sub

Sub simpan()
Dim jmlAwal As Double
Dim selisih As Double
Dim jmlEdit As Double
If Text1 = "" Then
    MsgBox "Data belum lengkap...", vbInformation, "Validasi Input"
    Text1.SetFocus
Else
    With Data1.Recordset
        jmlAwal = !jml_byr
        jmlEdit = Text1
        selisih = jmlEdit - jmlAwal
        .Edit
        !tgl_byr = DTPicker1
        !jml_byr = jmlEdit
        !keterangan = Text2
        .Update
    End With
    With F11_TBayar.Data1.Recordset
        .Edit
        !jumlah_bayar = !jumlah_bayar + selisih
        If !jumlah_bayar < !qty_beli * !harga_satuan Then
            !status_bayar = False
        Else
            !status_bayar = True
        End If
        .Update
    End With
    cmd_awal
    tutup
    isi_list
    isi
End If
End Sub

Private Sub Form_Load()
tutup
End Sub

Private Sub Form_Unload(Cancel As Integer)
F11_TBayar.Enabled = True
F11_TBayar.Show
End Sub

Sub isi_list()
Dim head As ColumnHeader
Dim dtl As ListItem
ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear
Set head = ListView1.ColumnHeaders.Add(, , "Tanggal", ListView1.Width / 5)
Set head = ListView1.ColumnHeaders.Add(, , "Jumlah Bayar", ListView1.Width / 5, 1)
Set head = ListView1.ColumnHeaders.Add(, , "Keterangan", ListView1.Width / (5 / 3) - 100)
ListView1.View = lvwReport
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set dtl = ListView1.ListItems.Add(, , !tgl_byr)
        dtl.SubItems(1) = Format(!jml_byr, "###,###.00")
        dtl.SubItems(2) = !keterangan
        .MoveNext
    Loop
End If
End With
Data1.Refresh
End Sub

Sub gerak()
Dim cek As Boolean
Dim tgl As Date
Dim ket As String
Dim jml As Double
tgl = ListView1.SelectedItem.Text
ket = ListView1.SelectedItem.SubItems(2)
jml = ListView1.SelectedItem.SubItems(1)
cek = False
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until cek = True Or .EOF
            If !tgl_byr = tgl And !keterangan = ket And !jml_byr = jml Then
                cek = True
                isi
                .MovePrevious
            End If
            .MoveNext
        Loop
    End If
End With
End Sub

Sub isi()
With Data1.Recordset
If Not .BOF Then
DTPicker1 = !tgl_byr
Text1 = Format(!jml_byr, "###,###.00")
Text2 = !keterangan
End If
End With
End Sub

Private Sub ListView1_Click()
gerak
End Sub

Sub kosong()
DTPicker1 = Date
Text1 = ""
Text2 = ""
End Sub

Sub tutup()
DTPicker1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
End Sub

Sub buka()
DTPicker1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
End Sub

Sub cmd_awal()
Command1(0).Picture = ImageList1.ListImages(1).Picture
Command1(0).Caption = "EDIT"
Command1(1).Caption = "HAPUS"
Command1(2).Visible = True
End Sub

Sub cmd_simpan()
Command1(0).Picture = ImageList1.ListImages(2).Picture
Command1(0).Caption = "SIMPAN"
Command1(1).Caption = "BATAL"
Command1(2).Visible = False
End Sub
