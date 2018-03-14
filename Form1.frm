VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000011&
   Caption         =   "ENTRY DATA DOSEN"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2895
      Left            =   2760
      TabIndex        =   12
      Top             =   4920
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
      ColumnCount     =   3
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "nama_dosen"
         Caption         =   "nama_dosen"
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
         DataField       =   "no_hp"
         Caption         =   "no_hp"
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
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2760
      Top             =   4440
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Left            =   5040
      Picture         =   "Form1.frx":0015
      TabIndex        =   11
      Top             =   3840
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
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
      Left            =   7080
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
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
      Left            =   8160
      TabIndex        =   8
      Top             =   3840
      Width           =   975
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
      Left            =   9120
      TabIndex        =   7
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   5760
      TabIndex        =   3
      Top             =   2880
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   5760
      TabIndex        =   2
      Top             =   2280
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   5760
      TabIndex        =   1
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "No HP"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Dosen"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Kode"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "ENTRY DATA DOSEN"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
mati (True)
Call control(True, True, True, False, False, True, True, True)
Text1.SetFocus
kosong
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Edit" Then
Call control(False, True, True, False, True, False, True, True)
Command2.Caption = "Save Edit"
Text2.SetFocus
Else
If Text1 = "" And Text2 = "" And Text3 = "" Then
    MsgBox "Masih ada data yang kosong..!!!", vbCritical, "Error!"
        Else
           If Adodc2.Recordset.BOF Or Adodc1.Recordset.EOF Then
        MsgBox "error"
        Else
db
With Adodc1.Recordset
    !kode_dosen = Text1
    !nama_dosen = Text2
    !no_hp = Text3
    .Update
End With
Call control(False, False, False, True, True, False, True, True)
Command2.Caption = "Edit"
    End If
    End If
End If

End Sub

Private Sub Command3_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
MsgBox "error"
Else
db
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("kode_dosen") = Text1
Adodc1.Recordset.Fields("nama_dosen") = Text2
Adodc1.Recordset.Fields("no_hp") = Text3
Adodc1.Recordset.Update
MsgBox "Disimpan!", vbOKOnly, "Berhasil!"
        kosong
        Call control(False, False, False, True, True, False, True, True)
        Call Form_Load
End If
End Sub

Private Sub Command4_Click()
Dim hapus As String
db
    If Adodc1.Recordset.RecordCount <> 0 Then
        hapus = MsgBox("Yakin akan dihapus?", vbYesNo, "Peringatan...!")
        If hapus = vbYes Then
            If Adodc1.Recordset.EOF Then
                MsgBox "kosong"
            Else
                Adodc1.Recordset.Delete
                Adodc1.Recordset.MoveNext
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
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
isi
mati (False)
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
isi
mati (False)
End Sub
Private Sub Form_Load()
isi
mati (False)
Call control(False, False, False, True, True, False, True, True)
End Sub
Sub isi()
If Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF Then
Me.Show
Else
Text1.Text = Adodc1.Recordset.Fields("kode_dosen")
Text2.Text = Adodc1.Recordset.Fields("nama_dosen")
Text3.Text = Adodc1.Recordset.Fields("no_hp")
End If
End Sub

Sub mati(x)
Text1.Enabled = x
Text2.Enabled = x
Text3.Enabled = x
End Sub

Function control(t1, t2, t3, a1, a2, a3, a4, a5)
Text1.Enabled = t1
Text2.Enabled = t2
Text3.Enabled = t3
Command1.Enabled = a1
Command2.Enabled = a2
Command3.Enabled = a3
Command4.Enabled = a4
Command5.Enabled = a5
End Function
