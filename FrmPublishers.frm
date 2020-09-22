VERSION 5.00
Begin VB.Form FrmPublishers 
   Caption         =   "FrmPublishers"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Publishers Details           "
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton CmdDel 
         Caption         =   "Delete Record"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2280
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1920
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   2280
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1560
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   2280
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1200
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   840
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   480
         Width           =   5535
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4800
         TabIndex        =   9
         Top             =   5280
         Width           =   1455
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         Top             =   5280
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   2535
         Index           =   0
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "FrmPublishers.frx":0000
         Top             =   2640
         Width           =   5535
      End
      Begin VB.Label Label6 
         Caption         =   "Comments"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "City"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Company Name"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Publisher Name"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Publisher#"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmPublishers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs As ADODB.Recordset

Private Sub CmdClose_Click()
On Error GoTo quit

rs.Close
Set rs = Nothing



quit:
Unload Me
End Sub

Private Sub CmdDel_Click()
On Error GoTo Error1

response = MsgBox("Are you sure?", vbCritical + vbYesNo, "Delete Record")


If response = vbNo Then Exit Sub

rs.Delete
rs.Close
Set rs = Nothing
Unload Me
FrmMain.ListColumns
FrmMain.LoadDetails

Exit Sub

Error1:
MsgBox Err.Description, vbCritical

End Sub

Private Sub CmdSave_Click()

On Error Resume Next

 If Len(Text1(1).Text) Then rs!Comments = "" & Text1(0).Text
 If Len(Text1(2).Text) Then rs!Name = "" & Text1(2).Text
 If Len(Text1(3).Text) Then rs![Company Name] = "" & Text1(3).Text
 If Len(Text1(4).Text) Then rs!address = "" & Text1(4).Text
 If Len(Text1(5).Text) Then rs!City = "" & Text1(5).Text


rs.Update
rs.Close
Set rs = Nothing
Unload Me
FrmMain.ListColumns
FrmMain.LoadDetails
End Sub



Private Sub Form_Load()

For i = 0 To 5
Text1(i).Text = ""          'Clear text Boxes
Next i


Call Get_Publishers         'Load Profile


End Sub

 Private Sub Get_Publishers()
 
Dim StrQry As String
Dim i As Integer
StrQry = "Select pubid, Name, [Company Name], " & _
"Address, City, Comments From Publishers where pubid = " & FrmMain.PubID

Set rs = New ADODB.Recordset

    rs.Open StrQry, FrmMain.DBConn, adOpenKeyset, adLockOptimistic

Text1(0).Text = "" & rs!Comments
Text1(1).Text = "" & rs!PubID
Text1(2).Text = "" & rs!Name
Text1(3).Text = "" & rs![Company Name]
Text1(4).Text = "" & rs!address
Text1(5).Text = "" & rs!City


End Sub

