VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form F19_cariMember 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CARI MEMBER"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   ClipControls    =   0   'False
   Icon            =   "F19_cariMember.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   5
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AMBIL"
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   6600
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   6600
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6800
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
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ALAMAT"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   11
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TELEPON"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   10
      Top             =   480
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA MEMBER"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMOR MEMBER"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SILAHKAN MASUKAN KATA KUNCI :"
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
      TabIndex        =   7
      Top             =   120
      Width           =   9480
   End
End
Attribute VB_Name = "F19_cariMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Data1.RecordSource = "select * from member where no_member like '*" & Text1 & "*' and nama_member like '*" & Text2 & "*' and telp_member like '*" & Text3 & "*' and alamat_member like '*" & Text4 & "*'"
    Data1.Refresh
    isi_list
    kosong
Case 1
    If Not Data1.Recordset.BOF Then
        If Text1 <> "" And Text2 <> "" Then
            F09_Nota.Combo1 = Text1
            Unload Me
        Else
            MsgBox "Silahkan pilih data terlebih dahulu", vbInformation, "Pilih Data"
            ListView1.SetFocus
        End If
    Else
        MsgBox "Silahkan data masih kosong/tidak diketemukan", vbInformation, "Cari Data"
    End If
Case 2
    kosong
    Text1.SetFocus
End Select
End Sub

Private Sub Form_Activate()
'isi_list
End Sub

Private Sub Form_Load()
kosong
limiter
Call db_cariMember
End Sub

Private Sub Form_Unload(Cancel As Integer)
F09_Nota.Enabled = True
F09_Nota.Show
End Sub

Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Sub limiter()
Text1.MaxLength = 10
Text2.MaxLength = 50
Text3.MaxLength = 50
Text4.MaxLength = 100
End Sub

Sub isi_list()
Dim head As ColumnHeader
Dim dtl As ListItem
ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear
Set head = ListView1.ColumnHeaders.Add(, , "NO", ListView1.Width / 6)
Set head = ListView1.ColumnHeaders.Add(, , "Nama", ListView1.Width / (6 / 2))
Set head = ListView1.ColumnHeaders.Add(, , "Telp", ListView1.Width / 6)
Set head = ListView1.ColumnHeaders.Add(, , "Alamat", ListView1.Width / (6 / 2))
ListView1.View = lvwReport
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Set dtl = ListView1.ListItems.Add(, , !no_member)
        dtl.SubItems(1) = !nama_member
        dtl.SubItems(2) = !telp_member
        dtl.SubItems(3) = !alamat_member
        .MoveNext
    Loop
    ListView1.SetFocus
'Else
'    MsgBox "Data tidak diketemukan...", vbInformation, "Tidak Ketemu"
End If
End With
Data1.Refresh
End Sub

Private Sub ListView1_Click()
If Not ListView1.ListItems.Count = 0 Then
isi
End If
End Sub

Sub isi()
Data1.Refresh
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do Until !no_member = ListView1.SelectedItem.Text
        .MoveNext
    Loop
    Text1 = !no_member
    Text2 = !nama_member
    Text3 = !telp_member
    Text4 = !alamat_member
Else
    kosong
End If
End With
End Sub

Private Sub ListView1_GotFocus()
If Not ListView1.ListItems.Count = 0 Then
    isi
End If
End Sub
