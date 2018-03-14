VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MENU"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   15615
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000008&
      Height          =   8805
      Left            =   0
      ScaleHeight     =   585
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   0
      Top             =   0
      Width           =   1185
      Begin VB.CommandButton Command4 
         Height          =   4095
         Left            =   0
         TabIndex        =   4
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "INPUT NILAI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "MATA KULIAH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "DOSEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   1575
         Left            =   -240
         Picture         =   "MDIForm1.frx":8E180
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   15615
      TabIndex        =   3
      Top             =   0
      Width           =   15615
   End
   Begin VB.Menu menu 
      Caption         =   "File"
      Begin VB.Menu abt 
         Caption         =   "About"
      End
      Begin VB.Menu ext 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub abt_Click()
frmAbout.Show
End Sub
Private Sub Command1_Click()
Form1.Show
Unload Form2
Unload Form3
End Sub

Private Sub Command2_Click()
Form2.Show
Unload Form1
Unload Form3
End Sub

Private Sub Command3_Click()
Form3.Show
Unload Form1
Unload Form2
End Sub

Private Sub ext_Click()
End
End Sub

