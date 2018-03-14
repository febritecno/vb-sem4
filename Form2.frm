VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   Caption         =   "Form2"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14955
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   14955
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   5640
      List            =   "Form2.frx":0002
      TabIndex        =   14
      Top             =   3360
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2640
      Top             =   4800
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DATA\Kuliah\tugasvb\dbakademik.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DATA\Kuliah\tugasvb\dbakademik.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "matakuliah"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   5640
      TabIndex        =   0
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   5640
      TabIndex        =   1
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   5640
      TabIndex        =   2
      Top             =   2760
      Width           =   4335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   8
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   7
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Picture         =   "Form2.frx":0004
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":040B
      Height          =   2895
      Left            =   2640
      TabIndex        =   3
      Top             =   5280
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "kode_mk"
         Caption         =   "kode_mk"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nama_mk"
         Caption         =   "nama_mk"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "sks"
         Caption         =   "sks"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "kode_dosen"
         Caption         =   "kode_dosen"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   120
      Top             =   7080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1720
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DATA\Kuliah\tugasvb\dbakademik.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DATA\Kuliah\tugasvb\dbakademik.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "dosen"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Kode Dosen"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "ENTRY DATA MATAKULIAH"
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Kode"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Nama MataKuliah"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "SKS"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   2760
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
mati (True)
Call control(True, True, True, False, False, True, True, True, True)
Text1.SetFocus
kosong
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Edit" Then
Call control(False, True, True, False, True, False, True, True, True)
Command2.Caption = "Save Edit"
Text2.SetFocus
Else
If Text1 = "" And Text2 = "" And Text3 = "" And Combo1 = "" Then
    MsgBox "Masih ada data yang kosong..!!!", vbCritical, "Error!"
        Else
        
db
With Adodc2.Recordset
    !kode_mk = Text1
    !nama_mk = Text2
    !sks = Text3
    !kode_dosen = Combo1
    .Update
End With
Call control(False, False, False, True, True, False, True, True, False)
Command2.Caption = "Edit"
    End If
End If

End Sub

Private Sub Command3_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
MsgBox "error"
Else
db
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields("kode_mk") = Text1
Adodc2.Recordset.Fields("nama_mk") = Text2
Adodc2.Recordset.Fields("sks") = Text3
Adodc2.Recordset.Fields("kode_dosen") = Combo1
Adodc2.Recordset.Update
MsgBox "Disimpan!", vbOKOnly, "Berhasil!"
        kosong
        Call control(False, False, False, True, True, False, True, True, False)
        Call Form_Load
End If
End Sub

Private Sub Command4_Click()
Dim hapus As String
db
    If Adodc2.Recordset.RecordCount <> 0 Then
        hapus = MsgBox("Yakin akan dihapus?", vbYesNo, "Peringatan...!")
        If hapus = vbYes Then
            If Adodc2.Recordset.EOF Then
                MsgBox "kosong"
            Else
                Adodc2.Recordset.Delete
                Adodc2.Recordset.MoveNext
                Call Form_Load
                End If
        End If
    Else
        MsgBox "Data kosong...", vbInformation, "Informasi!"
End If

End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Sub kosong()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub DataGrid1_Click()
isi
mati (False)
Call Form_Load
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
isi
mati (False)
Call Form_Load
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
isi
mati (False)
Call Form_Load
End Sub
Private Sub Form_Load()
Adodc1.Visible = False
l
isi
mati (False)
Call control(False, False, False, True, True, False, True, True, False)
End Sub
Sub isi()
If Adodc2.Recordset.EOF Or Adodc2.Recordset.BOF Then
Me.Show
Else

Text1.Text = Adodc2.Recordset.Fields("kode_mk")
Text2.Text = Adodc2.Recordset.Fields("nama_mk")
Text3.Text = Adodc2.Recordset.Fields("sks")
Command2.Caption = "Edit"
End If
End Sub

Sub mati(x)
Text1.Enabled = x
Text2.Enabled = x
Text3.Enabled = x
End Sub

Function control(t1, t2, t3, a1, a2, a3, a4, a5, d3)
Text1.Enabled = t1
Text2.Enabled = t2
Text3.Enabled = t3
Combo1.Enabled = d3
Command1.Enabled = a1
Command2.Enabled = a2
Command3.Enabled = a3
Command4.Enabled = a4
Command5.Enabled = a5
End Function



Sub l()
    'INI LOAD DARI TABLE DOSEN DI KIRIM KE COMBO1 PADA INPUTAN TABLE MATAKULIAH
    Adodc1.RecordSource = "select * from dosen"
    Do
        Combo1.AddItem Adodc1.Recordset.Fields("kode_dosen")
        Adodc1.Recordset.MoveNext
    Loop Until Adodc1.Recordset.EOF = True
End Sub
