VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form F15_Remain 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remainder/Pengingat"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   ClipControls    =   0   'False
   Icon            =   "F15_Remain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "DATA REMAINDER"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "TAMBAH"
      Height          =   735
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "EDIT"
      Height          =   735
      Index           =   1
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "HAPUS"
      Height          =   735
      Index           =   2
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808000&
      Caption         =   "AKTIF"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "F15_Remain.frx":3482
      Top             =   1560
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "F15_Remain.frx":3488
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "F15_Remain.frx":349C
      TabIndex        =   8
      Top             =   4560
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51314689
      CurrentDate     =   39631
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F15_Remain.frx":3E6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F15_Remain.frx":5801
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F15_Remain.frx":7193
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F15_Remain.frx":8B25
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F15_Remain.frx":A4B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F15_Remain.frx":B191
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "F15_Remain.frx":CB23
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   6765
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4233
            Text            =   "PT Arjuna Wahana Putra"
            TextSave        =   "PT Arjuna Wahana Putra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "11/1/2008"
            Key             =   "Tanggal Sekarang "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "2:40 PM"
            Key             =   "Jam Sekarang"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PILIH NOMOR TAGIHAN"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TANGGAL"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WAKTU"
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS"
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KETERANGAN"
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "F15_Remain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean

Sub isi2()
Data1.Refresh
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do Until Combo1 = !nomor
        .MoveNext
    Loop
    isi_data
End If
End With
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1(0).SetFocus
End If

End Sub

Private Sub Combo1_Click()
If tambah = False Then
    isi2
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DTPicker1.SetFocus
End If
End Sub



Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        cmd_simpan
        tambah = True
        buka
        kosong
        Call remain_auto
        Combo1 = urut_remain
        'isi_cmb1
        Combo1.Enabled = False
    Else
        simpan
    End If
Case 1
    If Command1(1).Caption = "EDIT" Then
        cmd_simpan
        tambah = False
        buka
        Combo1.Enabled = False
    Else
        tutup
        cmd_awal
        Data1.Refresh
        isi_cmb
        Combo1.Enabled = True
        tambah = False
    End If
Case 2
    MsgBox "Apakah anda yakin menghapus data", vbYesNo, "Hapus Data"
    If vbYes Then
        Data1.Recordset.Delete
        Data1.Refresh
    End If
End Select
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Text1.Enabled = False Then
    isi_data
End If
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
End If

End Sub

Private Sub Form_Activate()
isi_cmb
Data1.Refresh
Data2.Refresh
End Sub

Private Sub Form_Load()
Call db_remain
kosong
cmd_awal
tutup
End Sub

Sub tutup()
DTPicker1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Check1.Enabled = False
End Sub

Sub buka()
DTPicker1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Check1.Enabled = True
End Sub

Sub kosong()
Text1 = ""
Text2 = ""
Check1.Value = 0
End Sub

Sub isi_data()
With Data1.Recordset
If Not .BOF Then
    Combo1 = !nomor
    DTPicker1 = !tgl
    Text1 = !waktu
    Text2 = !keterangan
    If !Status = True Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End If
End With
End Sub

Sub isi_cmb1()
Combo1.Clear
Data2.RecordSource = "select* from pembayaran where sisa<>0"
Data2.Refresh
With Data2.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Combo1.AddItem !nomor
        .MoveNext
    Loop
    Combo1.ListIndex = 0
End If
End With
End Sub

Sub isi_cmb()
Combo1.Clear
Data1.Refresh
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Combo1.AddItem !nomor
        .MoveNext
    Loop
    Combo1.ListIndex = 0
End If
End With
Data1.Refresh
End Sub


Sub cmd_awal()
Command1(0).Picture = ImageList1.ListImages(1).Picture
Command1(1).Picture = ImageList1.ListImages(2).Picture
Command1(2).Picture = ImageList1.ListImages(4).Picture
Command1(0).Caption = "TAMBAH"
Command1(1).Caption = "EDIT"
Command1(2).Visible = True
End Sub

Sub cmd_simpan()
Command1(0).Picture = ImageList1.ListImages(5).Picture
Command1(1).Picture = ImageList1.ListImages(4).Picture
Command1(2).Picture = ImageList1.ListImages(4).Picture
Command1(0).Caption = "SIMPAN"
Command1(1).Caption = "BATAL"
Command1(2).Visible = False
End Sub

Sub simpan()
Dim cek As Boolean
If Text1 = "" Or Text2 = "" Then
    MsgBox "Data belum lengkap", vbCritical, "Validasi Input"
    If Text1 = "" Then
        Text1.SetFocus
    ElseIf Text2 = "" Then
        Text2.SetFocus
    End If
Else
    With Data1.Recordset
        If tambah = True Then
            cek = False
            If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                If Combo1 = !nomor Then
                    cek = True
                    .MoveLast
                End If
                .MoveNext
            Loop
            End If
            If cek = True Then
                MsgBox "Nomor sudah ada, silahkan pilih yang lain", vbCritical, "Validasi Data"
            Else
                .AddNew
                !nomor = Combo1
                !tgl = DTPicker1
                !waktu = Format(Text1, "hh.mm")
                !keterangan = Text2
                If Check1.Value = 1 Then
                    !Status = True
                Else
                    !Status = False
                End If
                .Update
                Data1.Refresh
                cmd_awal
                tutup
                isi_cmb
                tambah = False
            End If
        Else
            .Edit
            !nomor = Combo1
            !tgl = DTPicker1
            !waktu = Format(Text1, "hh.mm")
            !keterangan = Text2
            If Check1.Value = 1 Then
                !Status = True
            Else
                !Status = False
            End If
            .Update
            Data1.Refresh
            cmd_awal
            tutup
            Combo1.Enabled = True
            isi_cmb
        End If
    End With
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
F01_Main.Enabled = True
F01_Main.Data1.Refresh
F01_Main.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check1.SetFocus
End If

End Sub

