VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "Refresh List"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin MSComctlLib.ListView lvDetails 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by Chris Hatton in VB 6.0.
'This is a simple example of using VB to View a Microsoft Access
'database on a local machine.
'Be sure to include in your project references the Microsoft
'ActiveX Data Objects 2.5 Library.
'be sure to have the database in C:\Program Files\Microsoft Visual Studio\VB98\BIBLIO.MDB
'Feel free to email me with your comments or suggestions. chris@hatton.com

Public DBConn As ADODB.Connection
Public PubID As Long

Sub ListColumns()

lvDetails.FullRowSelect = True
lvDetails.View = lvwReport
lvDetails.GridLines = True
lvDetails.ColumnHeaders.Clear
lvDetails.ListItems.Clear

With lvDetails.ColumnHeaders                    ' Create column headers.
    .Add , , "Publisher#", 1000
    .Add , , "Publisher Name", 4400
    .Add , , "Company Name", 3500
    .Add , , "Address", 2000
    .Add , , "City", 1500
    .Add , , "Comments", 4500

End With

End Sub

Sub LoadDetails()
On Error GoTo ErrorOpen

Dim rs As ADODB.Recordset
Dim i As Integer

Set DBConn = New ADODB.Connection                                               '
                                                                                '   setting up connection.
DBConn.Provider = "Microsoft.JET.OLEDB.4.0"                                     '
                                                                                '
DBConn.Open "C:\Program Files\Microsoft Visual Studio\VB98\BIBLIO.MDB"

Set rs = New ADODB.Recordset

rs.Open "Select pubid, Name, [Company Name], " & _
"Address, City, Comments From Publishers", DBConn, adOpenKeyset, adLockReadOnly ' Select fields

Do

For i = 1 To rs.RecordCount

With lvDetails

           .ListItems.Add , , rs!PubID      'Create the first item in the column

    
    If Not rs!Name = "" Then .ListItems(i).ListSubItems.Add , , rs!Name                 'list subitems in the next column
    If Not rs![Company Name] = "" Then .ListItems(i).ListSubItems.Add , , rs![Company Name]
    If Not rs!address = "" Then .ListItems(i).ListSubItems.Add , , rs!address
    If Not rs!City = "" Then .ListItems(i).ListSubItems.Add , , rs!City
    If Not rs!Comments = "" Then .ListItems(i).ListSubItems.Add , , rs!Comments
    
    
End With

    rs.MoveNext
Next i

        
        
        Loop While rs.EOF = False                                   'if End Of Record is true then close the recordset.
        
        
        
rs.Close
Set rs = Nothing
Exit Sub

ErrorOpen:


MsgBox Err.Description, vbCritical
End

End Sub


Private Sub CmdRefresh_Click()
ListColumns
LoadDetails
End Sub


Private Sub Form_Load()
ListColumns                 'Load Columheaders
LoadDetails                 'Fill form
End Sub

Private Sub Form_Resize()
lvDetails.Width = FrmMain.Width - 150
lvDetails.Height = FrmMain.Height - 1200
End Sub

Private Sub Form_Unload(Cancel As Integer)
DBConn.Close
Set DBConn = Nothing
End Sub

Private Sub lvDetails_DblClick()

Dim FrmEdit As New frmpublishers
PubID = lvDetails.SelectedItem.Text

FrmEdit.Show



End Sub
