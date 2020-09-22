VERSION 5.00
Begin VB.Form addcusto 
   BackColor       =   &H8000000A&
   Caption         =   "ADD CUSTOMER INFORMATION"
   ClientHeight    =   5490
   ClientLeft      =   2400
   ClientTop       =   1920
   ClientWidth     =   7350
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7350
   Begin VB.CommandButton Command2 
      Caption         =   "Save Customer Information"
      Height          =   495
      Left            =   1936
      TabIndex        =   21
      Top             =   4598
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   2588
      TabIndex        =   20
      Top             =   2805
      Width           =   2415
   End
   Begin VB.CommandButton systemno 
      BackColor       =   &H00C0E0FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   5468
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2198
      Width           =   1215
   End
   Begin VB.CommandButton systemno 
      BackColor       =   &H00C0E0FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   5468
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1718
      Width           =   1215
   End
   Begin VB.CommandButton systemno 
      BackColor       =   &H00C0E0FF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   4268
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2198
      Width           =   1215
   End
   Begin VB.CommandButton systemno 
      BackColor       =   &H00C0E0FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   4268
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1718
      Width           =   1215
   End
   Begin VB.CommandButton systemno 
      BackColor       =   &H00C0E0FF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3068
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2198
      Width           =   1215
   End
   Begin VB.CommandButton systemno 
      BackColor       =   &H00C0E0FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   3068
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1718
      Width           =   1215
   End
   Begin VB.CommandButton systemno 
      BackColor       =   &H00C0E0FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1868
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2198
      Width           =   1215
   End
   Begin VB.CommandButton systemno 
      BackColor       =   &H00C0E0FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1868
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1718
      Width           =   1215
   End
   Begin VB.CommandButton systemno 
      BackColor       =   &H00C0E0FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   668
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2198
      Width           =   1215
   End
   Begin VB.CommandButton systemno 
      BackColor       =   &H00C0E0FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   668
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1718
      Width           =   1215
   End
   Begin VB.CommandButton cmd_time 
      Caption         =   "Current Time"
      Height          =   285
      Left            =   4995
      TabIndex        =   4
      Top             =   4118
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2595
      TabIndex        =   2
      Top             =   4125
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2595
      TabIndex        =   0
      Top             =   405
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "System Number"
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
      Left            =   668
      TabIndex        =   22
      Top             =   2798
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000A&
      Caption         =   "Allocated System"
      Height          =   255
      Left            =   4148
      TabIndex        =   19
      Top             =   3398
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      Caption         =   "Free System"
      Height          =   255
      Left            =   2108
      TabIndex        =   18
      Top             =   3398
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3788
      TabIndex        =   17
      Top             =   3398
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1748
      TabIndex        =   16
      Top             =   3398
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click on the System number to allocate it"
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
      Left            =   428
      TabIndex        =   15
      Top             =   1118
      Width           =   6495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   3015
      Left            =   428
      Shape           =   4  'Rounded Rectangle
      Top             =   878
      Width           =   6495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "In Time"
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
      Left            =   668
      TabIndex        =   3
      Top             =   4118
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Customer Name"
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
      Left            =   720
      TabIndex        =   1
      Top             =   405
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   5175
      Left            =   188
      Shape           =   4  'Rounded Rectangle
      Top             =   158
      Width           =   6975
   End
End
Attribute VB_Name = "addcusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub cmd_time_Click()
    Text3.Text = Now
End Sub



Private Sub Command2_Click()
If Len(Text1.Text) = 0 Then
    MsgBox "Please enter the customer name", vbInformation, "Customer name !!!"
ElseIf Len(Text4.Text) = 0 Then
    MsgBox "Please Assign the system number", vbInformation, "System Number !!!"
ElseIf Len(Text3.Text) = 0 Then
    MsgBox "Please assign the Incoming time", vbInformation, "In_Time !!!"
Else
    rs.AddNew
    rs.Fields(0).Value = Text1.Text
    rs.Fields(1).Value = Text4.Text
    rs.Fields(2).Value = Text3.Text
    rs.Update
    Unload addcusto
    
    
End If

End Sub

Private Sub Form_Load()
    
    dn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CAFE_DATABASE.mdb;Persist Security Info=False"
    rs.Open "select * from CURRENT_CUSTOMER", dn, adOpenDynamic, adLockOptimistic
    
    While rs.EOF <> True
        systemno(rs.Fields(1).Value - 1).BackColor = &HFF8080
        systemno(rs.Fields(1).Value - 1).Enabled = False
        rs.MoveNext
    Wend
    
    
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dn.Close
End Sub

Private Sub systemno_Click(Index As Integer)
Text4.Text = Index + 1
End Sub
