VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CAFE SOFTWARE"
   ClientHeight    =   2880
   ClientLeft      =   4035
   ClientTop       =   3075
   ClientWidth     =   4245
   Icon            =   "firstform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4245
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3360
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   630
      TabIndex        =   2
      Top             =   1575
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Payment"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   630
      TabIndex        =   1
      Top             =   975
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   630
      TabIndex        =   0
      Top             =   375
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Version is Under Construction ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
addcusto.Show vbModal
End Sub

Private Sub Command2_Click()
Form2.Show vbModal
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
frmAbout.Show vbModal
End Sub
