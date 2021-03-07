VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form F02_MsDb 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Database"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   ClipControls    =   0   'False
   Icon            =   "F02_MsDb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "DELETE ALL"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DELETE"
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "F02_MsDb.frx":3482
      Height          =   5175
      Left            =   120
      OleObjectBlob   =   "F02_MsDb.frx":3496
      TabIndex        =   2
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "Data1"
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
      Top             =   600
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0080FFFF&
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA DATABASE"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2025
   End
End
Attribute VB_Name = "F02_MsDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Data1.Caption = Combo1
Data1.RecordSource = Combo1
Data1.Refresh
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Not Data1.Recordset.BOF Then
        x = MsgBox("Apakah anda yakin menghapus data...???", vbYesNo, "Hapus Record")
        If x = vbYes And Not Data1.Recordset.BOF Then
            Data1.Recordset.Delete
        End If
    Else
        MsgBox "Data Masih Kosong", vbInformation, "Blank Data"
    End If
Case 1
    If Not Data1.Recordset.BOF Then
        x = MsgBox("Apakah anda yakin menghapus SELURUH DATA " & Combo1 & "...???", vbYesNo, "Hapus Semua Data")
        If x = vbYes Then
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
                Data1.Recordset.Delete
                Data1.Recordset.MoveNext
            Loop
            Data1.Refresh
        End If
    Else
        MsgBox "Data Masih Kosong", vbInformation, "Blank Data"
    End If
End Select
End Sub

Private Sub Form_Load()
Call db_Master
isi_cmb
End Sub

Sub isi_cmb()
Combo1.Clear
Combo1.AddItem "ByrBeli"
Combo1.AddItem "ByrJual"
Combo1.AddItem "Harga"
Combo1.AddItem "Kas"
Combo1.AddItem "Lokasi"
Combo1.AddItem "Member"
Combo1.AddItem "Nota"
Combo1.AddItem "Pembelian"
Combo1.AddItem "Penjualan"
Combo1.AddItem "Petugas"
Combo1.AddItem "Produk"
Combo1.AddItem "Stok"
Combo1.AddItem "Supplier"
Combo1.AddItem "Temp_Nota"
Combo1.AddItem "Remainder"
Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
F01_Main.Enabled = True
F01_Main.Show
End Sub
