VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Customer and System Details ..."
   ClientHeight    =   6165
   ClientLeft      =   2910
   ClientTop       =   1530
   ClientWidth     =   6030
   FillStyle       =   0  'Solid
   ForeColor       =   &H00404040&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6030
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "10"
      Enabled         =   0   'False
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   29
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "9"
      Enabled         =   0   'False
      Height          =   495
      Index           =   8
      Left            =   120
      TabIndex        =   28
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "8"
      Enabled         =   0   'False
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   27
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "7"
      Enabled         =   0   'False
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "6"
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "5"
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "4"
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "3"
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "2"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "1"
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Calculate"
      Height          =   495
      Index           =   9
      Left            =   4680
      TabIndex        =   19
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Calculate"
      Height          =   495
      Index           =   8
      Left            =   4680
      TabIndex        =   18
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Calculate"
      Height          =   495
      Index           =   7
      Left            =   4680
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Calculate"
      Height          =   495
      Index           =   6
      Left            =   4680
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Calculate"
      Height          =   495
      Index           =   5
      Left            =   4680
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Calculate"
      Height          =   495
      Index           =   4
      Left            =   4680
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Calculate"
      Height          =   495
      Index           =   3
      Left            =   4680
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Calculate"
      Height          =   495
      Index           =   2
      Left            =   4680
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Calculate"
      Height          =   495
      Index           =   1
      Left            =   4680
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Calculate"
      Height          =   495
      Index           =   0
      Left            =   4680
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   9
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   8
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   7
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   5
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H8000000F&
      Height          =   495
      Index           =   9
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   8
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   7
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   6
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   5
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   4
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   3
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   2
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS1_COUNT As New ADODB.Recordset
Dim in_time As Date
Dim out_time As Date
Dim h As Double
Dim m As Double

Private Sub Command2_Click(Index As Integer)
RS1.Close
RS1_COUNT.Close
RS1.Open "select * from CURRENT_CUSTOMER WHERE SYSTEM_NO=" & Index + 1
RS1_COUNT.Open "select count(*) from CURRENT_CUSTOMER WHERE SYSTEM_NO=" & Index + 1


If RS1_COUNT.Fields(0).Value = 1 Then

Form3.SYS_NO = Index + 1



in_time = RS1.Fields(2).Value




out_time = Now

Form3.Text1 = RS1.Fields(0).Value
Form3.Text2 = RS1.Fields(1).Value
Form3.Text3 = RS1.Fields(2).Value
Form3.Text4 = out_time

Form3.Text9 = Clear

Form3.Text9 = DateDiff("n", in_time, out_time)

    m = Val(Form3.Text9) Mod 60
    h = (Form3.Text9) / 60
    h = Int(h)
    
    Form3.Text5.Text = Val(h)
    Form3.Text9.Text = Val(m)

    

''''''



'If DatePart("h", in_time) >= 22 Then
    'MsgBox DatePart("h", in_time)
    'If DatePart("h", in_time) = 24 Then
     '       If Val(Form3.Text9) <= 15 Then
      '          Form3.Text7 = 3.75
       '     ElseIf Val(Form3.Text9) <= 30 Then
       '         Form3.Text7 = 7.5
       '     ElseIf Val(Form3.Text9) <= 45 Then
       '         Form3.Text7 = 11.25
       '     ElseIf Val(Form3.Text9) <= 60 Then
       '         Form3.Text7 = 15
       '     Else
       '         Form3.Text7 = (Val(Form3.Text9) * 15) / 60
       '     End If
      'Form3.Text7 = Val(Form3.Text7) + (Val(Form3.Text5) * 15)
    'End If
    
If DatePart("h", in_time) >= 0 Then
    
    If DatePart("h", in_time) <= 6 Then
            If Val(Form3.Text9) <= 15 Then
                Form3.Text7 = 3.75
            ElseIf Val(Form3.Text9) <= 30 Then
                Form3.Text7 = 7.5
            ElseIf Val(Form3.Text9) <= 45 Then
                Form3.Text7 = 11.25
            ElseIf Val(Form3.Text9) <= 60 Then
                Form3.Text7 = 15
            Else
                Form3.Text7 = (Val(Form3.Text9) * 15) / 60
            End If
            Form3.Text7 = Val(Form3.Text7) + (Val(Form3.Text5) * 15)
            
    ElseIf DatePart("h", in_time) <= 23 Then
            If Val(Form3.Text9) <= 15 Then
                Form3.Text7 = 5
            ElseIf Val(Form3.Text9) <= 30 Then
                Form3.Text7 = 10
            ElseIf Val(Form3.Text9) <= 45 Then
                Form3.Text7 = 15
            ElseIf Val(Form3.Text9) <= 60 Then
                Form3.Text7 = 20
            Else
                Form3.Text7 = (Val(Form3.Text9) * 20) / 60
            End If
            Form3.Text7 = Val(Form3.Text7) + (Val(Form3.Text5) * 20)
    End If

End If



 '   If Val(Form3.Text9) <= 15 Then
 '       Form3.Text7 = 5
 '   ElseIf Val(Form3.Text9) <= 30 Then
 '       Form3.Text7 = 10
 '   ElseIf Val(Form3.Text9) <= 45 Then
 '       Form3.Text7 = 15
 '   ElseIf Val(Form3.Text9) <= 60 Then
 '       Form3.Text7 = 20
 '   Else
 '       Form3.Text7 = (Val(Form3.Text9) * 20) / 60
 '   End If

 '    Form3.Text7 = Val(Form3.Text7) + (Val(Form3.Text5) * 20)
    

Unload Me
Form3.Show vbModal
Else
    MsgBox "There is no Customer on that System...", vbInformation, "No Customer Found On that System"
    
End If

End Sub

Private Sub Form_Load()
   dn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CAFE_DATABASE.mdb;Persist Security Info=False"
   rs.Open "select * from CURRENT_CUSTOMER", dn, adOpenDynamic, adLockOptimistic
   RS1.Open "select * from CURRENT_CUSTOMER", dn, adOpenDynamic, adLockOptimistic
   RS1_COUNT.Open "select count(*) from current_customer", dn, adOpenDynamic, adLockOptimistic
   While rs.EOF <> True
         Label1(rs.Fields(1).Value - 1).Caption = rs.Fields(0).Value
         rs.MoveNext
    Wend
   
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dn.Close
End Sub
