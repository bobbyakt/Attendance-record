VERSION 5.00
Begin VB.Form regfrm 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Enter your password  :"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Enter your unique ID :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Registration"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "regfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection


Private Sub Command1_Click()
Set rL = New ADODB.Recordset
rL.Open "select * from loginRecord", conn, adOpenDynamic, adLockOptimistic
rL.AddNew
rL.Fields("ID") = Text1.Text
rL.Fields("pass") = Text2.Text

rL.Update
End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database2.mdb;Persist Security Info=False"
conn.Open
End Sub
