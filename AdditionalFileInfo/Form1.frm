VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Additional File Information by Vanja Fuckar"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   3720
      Width           =   5775
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   240
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get From File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   240
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get From Current Module"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "All Available Information in String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3480
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Extracted Information from String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Inicializiraj!
Text1 = InitFileInfo

'Imena u zagradi moraju biti baš takva kako sam napisao!
'Naravno u EXEu æe se pojaviti te informacije,a ne u IN-design modu!
Text2 = ""
Text2 = "Comments:" & FindFromStringInfo("Comments") & vbCrLf
Text2 = Text2 & "Product Name:" & FindFromStringInfo("ProductName") & vbCrLf
Text2 = Text2 & "File Version:" & FindFromStringInfo("FileVersion") & vbCrLf
Text2 = Text2 & "Legal Copyright:" & FindFromStringInfo("LegalCopyright") & vbCrLf
Text2 = Text2 & "File Description:" & FindFromStringInfo("FileDescription") & vbCrLf
Text2 = Text2 & "Legal Trademarks:" & FindFromStringInfo("LegalTrademarks") & vbCrLf
Text2 = Text2 & "Company Name:" & FindFromStringInfo("CompanyName")

End Sub

Private Sub Command2_Click()
cd1.ShowOpen
If Len(cd1.Filename) = 0 Then Exit Sub

Dim IsAvailable As Long

Text1 = InitFileInfo(cd1.Filename, IsAvailable)

If IsAvailable = 0 Then MsgBox "File information doesn't exist!", vbExclamation, "Info!": Exit Sub



Text2 = ""
Text2 = "Comments:" & FindFromStringInfo("Comments") & vbCrLf
Text2 = Text2 & "Product Name:" & FindFromStringInfo("ProductName") & vbCrLf
Text2 = Text2 & "File Version:" & FindFromStringInfo("FileVersion") & vbCrLf
Text2 = Text2 & "Legal Copyright:" & FindFromStringInfo("LegalCopyright") & vbCrLf
Text2 = Text2 & "File Description:" & FindFromStringInfo("FileDescription") & vbCrLf
Text2 = Text2 & "Legal Trademarks:" & FindFromStringInfo("LegalTrademarks") & vbCrLf
Text2 = Text2 & "Company Name:" & FindFromStringInfo("CompanyName")
End Sub

