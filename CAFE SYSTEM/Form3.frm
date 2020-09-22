VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Calculate Customer Payment"
   ClientHeight    =   4125
   ClientLeft      =   3000
   ClientTop       =   2520
   ClientWidth     =   5730
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5730
   Begin VB.ComboBox text8 
      Height          =   315
      ItemData        =   "Form3.frx":0442
      Left            =   2280
      List            =   "Form3.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   15
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calcel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AMOUNT PAID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   13
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2235
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2235
      TabIndex        =   8
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2235
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2235
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2235
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2235
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   3855
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Received By"
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
      Left            =   360
      TabIndex        =   12
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Amt"
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
      Left            =   360
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Hrs."
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
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Out_Time"
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
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "In_Time"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "System_Number"
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
      Left            =   360
      TabIndex        =   3
      Top             =   840
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
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Public SYS_NO As String

Private Sub Command1_Click()

If Len(text8.Text) = 0 Then
    MsgBox "Please Enter Receiver name", vbInformation, "Enter All the Details..."
Else
    rs.AddNew
    rs.Fields(0).Value = Text1.Text
    rs.Fields(1).Value = Text2.Text
    rs.Fields(2).Value = Text3.Text
    rs.Fields(3).Value = Text4.Text
    rs.Fields(4).Value = Text5.Text
    rs.Fields(5).Value = Text7.Text
    rs.Fields(6).Value = text8.Text
    rs.Update
    RS1.Delete
    Unload Me

End If


End Sub

Private Sub Command2_Click()
    Unload Me
    Form2.Show vbModal
End Sub

Private Sub Form_Load()
dn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CAFE_DATABASE.mdb;Persist Security Info=False"
rs.Open "select * from MASTER_TABLE", dn, adOpenDynamic, adLockOptimistic
RS1.Open "select * from CURRENT_CUSTOMER WHERE SYSTEM_NO =" & SYS_NO, dn, adOpenDynamic, adLockOptimistic

End Sub

Private Sub Form_Unload(Cancel As Integer)
dn.Close
End Sub

Private Sub Text6_LostFocus()
    If Val(Text9) <= 15 Then
        Text7 = 5
    ElseIf Val(Text9) <= 30 Then
        Text7 = 10
    ElseIf Val(Text9) <= 45 Then
        Text7 = 15
    ElseIf Val(Text9) <= 60 Then
        Text7 = 20
    Else
        Text7 = (Val(Text9) * 20) / 60
    End If
End Sub

