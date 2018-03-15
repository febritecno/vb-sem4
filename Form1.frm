VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{7FEC7313-D161-427C-A141-48E17931414B}#1.0#0"; "truedc8.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "ENTRY DATA DOSEN"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin TDBNumber6Ctl.TDBNumber Text3 
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   2880
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   661
      Calculator      =   "Form1.frx":0000
      Caption         =   "Form1.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form1.frx":008C
      Keys            =   "Form1.frx":00AA
      Spin            =   "Form1.frx":00F4
      AlignHorizontal =   2
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   1
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "+62 ########;(########);+62 "
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "+62 ###,###,######"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999990
      MinValue        =   -999999999999990
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBText6Ctl.TDBText Text2 
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   2280
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   661
      Caption         =   "Form1.frx":011C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form1.frx":0188
      Key             =   "Form1.frx":01A6
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBText6Ctl.TDBText Text1 
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   1680
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   661
      Caption         =   "Form1.frx":01EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form1.frx":0256
      Key             =   "Form1.frx":0274
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   0
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TrueData80Ctl.TData TData1 
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   4560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "TData1"
      BackColor       =   8421504
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ErrorMsgCaption =   ""
      Filtered        =   0   'False
      DataMode        =   0
      DataMember      =   ""
      NameSubstitute  =   ""
      ConnectionString=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DATA\Kuliah\tugasvb\dbakademik.mdb;Persist Security Info=False"
      ConnectStringType=   1
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\DATA\Kuliah\tugasvb\dbakademik.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      CursorLocation  =   3
      ConnectionTimeout=   15
      CommandTimeout  =   30
      RecordSource    =   "dosen"
      CursorType      =   3
      CommandType     =   2
      MaxRecords      =   0
      LinkType        =   0
      Master          =   ""
      CallDataRead    =   0   'False
      ConvertNullToEmpty=   -1  'True
      DesignConnection=   -1  'True
      DesignTimeout   =   5
      MousePointer    =   0
      Enabled         =   -1  'True
      BOFAction       =   0
      EOFAction       =   0
      QueryMode       =   0   'False
      Orientation     =   0
      ButtonFirst     =   -1  'True
      ButtonNext      =   -1  'True
      ButtonPage      =   0   'False
      ButtonAdd       =   0   'False
      ButtonDelete    =   0   'False
      ButtonUpdate    =   0   'False
      ButtonCancel    =   0   'False
      ButtonBookmark  =   0   'False
      ButtonFind      =   0   'False
      ButtonQuery     =   0   'False
      Tooltips        =   0   'False
      PageSize        =   10
      ConfirmDelete   =   -1  'True
      ConfirmCancel   =   0   'False
      LockType        =   3
      CallDataWrite   =   0   'False
      ConvertEmptyToNull=   -1  'True
      ResyncAfterUpdate=   0   'False
      ManualUpdate    =   0   'False
      RefreshOnSrcChange=   -1  'True
      CacheSize       =   50
      Mode            =   0
      ErrorMsgRestore =   -1  'True
      AutoRefresh     =   2
      AllowEarlyOpen  =   0   'False
      SafeMode        =   0   'False
      Virgin          =   -1  'True
      Fields.Count    =   3
      Fields(0).Name  =   "kode_dosen"
      Fields(0).DisplayName=   "kode_dosen"
      Fields(0).FieldKind=   0
      Fields(0).DataSourceField=   "kode_dosen"
      Fields(0).MaxLength=   10
      Fields(1).Name  =   "nama_dosen"
      Fields(1).DisplayName=   "nama_dosen"
      Fields(1).FieldKind=   0
      Fields(1).DataSourceField=   "nama_dosen"
      Fields(1).MaxLength=   100
      Fields(2).Name  =   "no_hp"
      Fields(2).DisplayName=   "no_hp"
      Fields(2).FieldKind=   0
      Fields(2).DataSourceField=   "no_hp"
      Fields(2).MaxLength=   20
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":02B8
      Height          =   2895
      Left            =   3120
      TabIndex        =   9
      Top             =   5040
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
      Top             =   19440
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
      Caption         =   ""
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
      Picture         =   "Form1.frx":02CD
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NO HP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA DOSEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "KODE DOSEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENTRY DATA DOSEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   720
      Width           =   6735
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
db
a
With TData1.Recordset
    !kode_dosen = Text1
    !nama_dosen = Text2
    !no_hp = Text3
    .Update
End With
Call control(False, False, False, True, True, False, True, True)
Command2.Caption = "Edit"
    End If
End If

End Sub

Private Sub Command3_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Value = "" Then
MsgBox "error"
Else
db
a
TData1.Recordset.AddNew
TData1.Recordset.Fields("kode_dosen") = Text1
TData1.Recordset.Fields("nama_dosen") = Text2
TData1.Recordset.Fields("no_hp") = Text3
TData1.Recordset.Update
MsgBox "Disimpan!", vbOKOnly, "Berhasil!"
        kosong
        Call control(False, False, False, True, True, False, True, True)
        Call Form_Load
End If
End Sub

Private Sub Command4_Click()
Dim hapus As String
db
 a
    If TData1.Recordset.RecordCount <> 0 Then
        hapus = MsgBox("Yakin akan dihapus?", vbYesNo, "Peringatan...!")
        If hapus = vbYes Then
            If TData1.Recordset.EOF Then
                MsgBox "kosong"
            Else
                TData1.Recordset.Delete
                TData1.Recordset.MoveNext
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
Text3.Value = ""
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
isi
Command2.Caption = "Edit"
mati (False)
Call control(False, False, False, True, True, False, True, True)
End Sub
Sub isi()
On Error Resume Next
Text1.Text = TData1.Recordset.Fields("kode_dosen")
Text2.Text = TData1.Recordset.Fields("nama_dosen")
Text3.Value = TData1.Recordset.Fields("no_hp")

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
