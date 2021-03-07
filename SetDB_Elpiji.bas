Attribute VB_Name = "SetDB_Elpiji"
Option Explicit
Public db As String
Public ctl1 As Boolean
Public ctl2 As Boolean
Public ctl3 As Boolean
Public ctl4 As Boolean
Public ctl5 As Boolean

Private Sub buka_db()
    db = App.Path & "\dbelpiji.mdb"
End Sub

Public Sub db_main()
buka_db
F01_Main.Data1.DatabaseName = db
End Sub

Public Sub db_produk()
buka_db
F03_Produk.Data1.DatabaseName = db
F03_Produk.Data1.RecordSource = "produk"
End Sub

Public Sub db_supplier()
buka_db
F04_Supplier.Data1.DatabaseName = db
F04_Supplier.Data1.RecordSource = "supplier"
End Sub

Public Sub db_petugas()
buka_db
F05_Petugas.Data1.DatabaseName = db
F05_Petugas.Data1.RecordSource = "petugas"
End Sub

Public Sub db_member()
buka_db
F06_Member.Data1.DatabaseName = db
F06_Member.Data1.RecordSource = "member"
F06_Member.Data2.DatabaseName = db
F06_Member.Data2.RecordSource = "lokasi"
End Sub

Public Sub db_harga()
buka_db
F07_Harga.Data1.DatabaseName = db
F07_Harga.Data1.RecordSource = "HARGA"
F07_Harga.Data2.DatabaseName = db
F07_Harga.Data2.RecordSource = "produk"
End Sub

Public Sub db_stok()
buka_db
F08_TStok.Data1.DatabaseName = db
F08_TStok.Data1.RecordSource = "HARGA"
F08_TStok.Data2.DatabaseName = db
F08_TStok.Data2.RecordSource = "produk"
F08_TStok.Data3.DatabaseName = db
F08_TStok.Data3.RecordSource = "select * from stok order by kode_produk asc"
F08_TStok.Data4.DatabaseName = db
F08_TStok.Data5.DatabaseName = db
F08_TStok.Data5.RecordSource = "supplier"
End Sub

Public Sub db_IsiStok()
buka_db
F16_isiStok.Data1.DatabaseName = db
'F16_isiStok.Data1.RecordSource = "stok"
F16_isiStok.Data2.DatabaseName = db
F16_isiStok.Data2.RecordSource = "produk"
F16_isiStok.Data3.DatabaseName = db
F16_isiStok.Data3.RecordSource = "supplier"
F16_isiStok.Data4.DatabaseName = db
F16_isiStok.Data4.RecordSource = "HARGA"
F16_isiStok.Data5.DatabaseName = db
F16_isiStok.Data5.RecordSource = "pembelian"
F16_isiStok.Data6.DatabaseName = db
F16_isiStok.Data6.RecordSource = "byrbeli"
F16_isiStok.Data7.DatabaseName = db
F16_isiStok.Data7.RecordSource = "kas"
F16_isiStok.Data8.DatabaseName = db
F16_isiStok.Data8.RecordSource = "remainder"
End Sub

Public Sub db_lokasi()
buka_db
F17_Lokasi.Data1.DatabaseName = db
F17_Lokasi.Data1.RecordSource = "lokasi"
End Sub

Public Sub db_isiUlang()
buka_db
F18_isiUlang.Data1.DatabaseName = db
F18_isiUlang.Data1.RecordSource = "stok"
F18_isiUlang.Data2.DatabaseName = db
F18_isiUlang.Data2.RecordSource = "produk"
F18_isiUlang.Data3.DatabaseName = db
F18_isiUlang.Data3.RecordSource = "supplier"
F18_isiUlang.Data4.DatabaseName = db
F18_isiUlang.Data5.DatabaseName = db
F18_isiUlang.Data5.RecordSource = "pembelian"
F18_isiUlang.Data6.DatabaseName = db
F18_isiUlang.Data6.RecordSource = "byrbeli"
F18_isiUlang.Data7.DatabaseName = db
F18_isiUlang.Data7.RecordSource = "kas"
F18_isiUlang.Data8.DatabaseName = db
F18_isiUlang.Data8.RecordSource = "remainder"
End Sub

Public Sub db_Nota()
buka_db
With F09_Nota
.dt_harga.DatabaseName = db
.dt_harga.RecordSource = "harga"
.dt_lokasi.DatabaseName = db
.dt_lokasi.RecordSource = "lokasi"
.dt_member.DatabaseName = db
.dt_member.RecordSource = "member"
.dt_petugas.DatabaseName = db
.dt_petugas.RecordSource = "petugas"
.dt_produk.DatabaseName = db
.dt_produk.RecordSource = "Produk"
.dt_stok.DatabaseName = db
.dt_stok.RecordSource = "stok"
.dt_list.DatabaseName = db
.dt_list.RecordSource = "temp_nota"
.dt_nota.DatabaseName = db
.dt_nota.RecordSource = "Nota"
End With
End Sub

Public Sub db_cariMember()
buka_db
F19_cariMember.Data1.DatabaseName = db
F19_cariMember.Data1.RecordSource = "member"
End Sub

Public Sub db_ctknota()
buka_db
F20_ctkNota.Data1.DatabaseName = db
F20_ctkNota.Data1.RecordSource = "produk"
F20_ctkNota.Data2.DatabaseName = db
F20_ctkNota.Data3.DatabaseName = db
F20_ctkNota.Data3.RecordSource = "penjualan"
F20_ctkNota.Data4.DatabaseName = db
F20_ctkNota.Data4.RecordSource = "byrjual"
F20_ctkNota.Data5.DatabaseName = db
F20_ctkNota.Data5.RecordSource = "kas"
F20_ctkNota.Data6.DatabaseName = db
F20_ctkNota.Data6.RecordSource = "remainder"
End Sub

Public Sub db_UlangNota()
buka_db
F21_UlangNota.Data1.DatabaseName = db
F21_UlangNota.Data1.RecordSource = "nota"
End Sub

Public Sub db_TKas()
buka_db
F10_TKas.Data1.DatabaseName = db
F10_TKas.Data1.RecordSource = "kas"
End Sub

Public Sub db_TBayar()
buka_db
F11_TBayar.Data1.DatabaseName = db
F11_TBayar.Data1.RecordSource = "pembelian"
F11_TBayar.Data2.DatabaseName = db
F11_TBayar.Data2.RecordSource = "supplier"
F11_TBayar.Data3.DatabaseName = db
F11_TBayar.Data3.RecordSource = "produk"
F11_TBayar.Data4.DatabaseName = db
F11_TBayar.Data4.RecordSource = "penjualan"
F11_TBayar.Data5.DatabaseName = db
F11_TBayar.Data5.RecordSource = "member"
End Sub

Public Sub db_byrbeli()
buka_db
F22_ByrBeli.Data1.DatabaseName = db
F22_ByrBeli.Data1.RecordSource = "byrbeli"
End Sub

Public Sub db_EDITbyrbeli()
buka_db
F23_EditByrBeli.Data1.DatabaseName = db
End Sub

Public Sub db_byrJUAL()
buka_db
F24_ByrJual.Data1.DatabaseName = db
F24_ByrJual.Data1.RecordSource = "byrjual"
End Sub

Public Sub db_EDITbyrJUAL()
buka_db
F25_EditByrJual.Data1.DatabaseName = db
End Sub

Public Sub db_Master()
buka_db
F02_MsDb.Data1.DatabaseName = db
End Sub

Public Sub db_remain()
buka_db
With F15_Remain
    .Data1.DatabaseName = db
    .Data2.DatabaseName = db
'    .Data2.RecordSource = "select * from pembayaran where sisa<>0"
    .Data1.RecordSource = "Remainder"
End With
End Sub

Public Sub db_remain2()
buka_db
REmain_frm.Data1.DatabaseName = db
End Sub

