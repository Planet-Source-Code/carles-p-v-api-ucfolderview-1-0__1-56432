VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ucFolderView 1.0.1 - Test"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.ucFolderView ucFolderView 
      Height          =   6045
      Left            =   150
      TabIndex        =   9
      Top             =   180
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   10663
   End
   Begin VB.CheckBox chkTrackSelect 
      Caption         =   "TrackSelect"
      Height          =   285
      Left            =   4680
      TabIndex        =   8
      Top             =   3765
      Width           =   2640
   End
   Begin VB.CheckBox chkSingleExpand 
      Caption         =   "SingleExpand (+[Ctl]: no collapse)"
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   3360
      Width           =   3285
   End
   Begin VB.CheckBox chkHideSelection 
      Caption         =   "HideSelection"
      Height          =   285
      Left            =   4680
      TabIndex        =   6
      Top             =   2955
      Width           =   2640
   End
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   4455
      TabIndex        =   1
      Top             =   435
      Width           =   3885
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   435
      Left            =   7245
      TabIndex        =   2
      Top             =   825
      Width           =   1095
   End
   Begin VB.CheckBox chkHasButtons 
      Caption         =   "HasButtons"
      Height          =   285
      Left            =   4680
      TabIndex        =   4
      Top             =   2145
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.CheckBox chkHasLines 
      Caption         =   "HasLines"
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Top             =   2550
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.Label lblProperties 
      Caption         =   "Properties (default)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4455
      TabIndex        =   3
      Top             =   1740
      Width           =   1905
   End
   Begin VB.Label lblPath 
      Caption         =   "Change path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4455
      TabIndex        =   0
      Top             =   165
      Width           =   1815
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    
    Call ucFolderView.Initialize
End Sub


Private Sub ucFolderView_ChangeBefore(ByVal NewPath As String, Cancel As Boolean)
    
    '-- Check paths here...
End Sub

Private Sub ucFolderView_ChangeAfter(ByVal OldPath As String)

    txtPath.Text = ucFolderView.Path
    txtPath.SelStart = Len(txtPath.Text)
End Sub

Private Sub cmdApply_Click()
    
    ucFolderView.Path = txtPath.Text
End Sub

'//

Private Sub chkHasButtons_Click()
    ucFolderView.HasButtons = CBool(chkHasButtons)
End Sub

Private Sub chkHasLines_Click()
    ucFolderView.HasLines = CBool(chkHasLines)
End Sub

Private Sub chkHideSelection_Click()
    ucFolderView.HideSelection = CBool(chkHideSelection)
End Sub

Private Sub chkSingleExpand_Click()
    ucFolderView.SingleExpand = CBool(chkSingleExpand)
End Sub

Private Sub chkTrackSelect_Click()
    ucFolderView.TrackSelect = CBool(chkTrackSelect)
End Sub


