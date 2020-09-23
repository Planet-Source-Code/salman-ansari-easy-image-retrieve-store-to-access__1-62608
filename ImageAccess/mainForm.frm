VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form mainForm 
   Caption         =   "Images Store/Retrieve From Access"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   330
      Left            =   1935
      TabIndex        =   7
      Top             =   2610
      Width           =   960
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1125
      TabIndex        =   6
      Top             =   2610
      Width           =   690
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   330
      Left            =   810
      TabIndex        =   3
      Top             =   2160
      Width           =   600
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   330
      Left            =   2295
      TabIndex        =   2
      Top             =   2160
      Width           =   600
   End
   Begin MSComDlg.CommonDialog browseDialog 
      Left            =   2385
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   330
      Left            =   3240
      TabIndex        =   1
      Top             =   585
      Width           =   1050
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   330
      Left            =   3240
      TabIndex        =   0
      Top             =   135
      Width           =   1050
   End
   Begin VB.Label Label3 
      Caption         =   "Open"
      Height          =   240
      Left            =   135
      TabIndex        =   8
      Top             =   2610
      Width           =   600
   End
   Begin VB.Image Picture1 
      Height          =   1995
      Left            =   135
      Top             =   90
      Width           =   2985
   End
   Begin VB.Label Label2 
      Caption         =   "Browse"
      Height          =   240
      Left            =   135
      TabIndex        =   5
      Top             =   2205
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "ID : "
      Height          =   195
      Left            =   810
      TabIndex        =   4
      Top             =   2655
      Width           =   285
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'       created by : salman ansari
'       email : salmanansari99@hotmail.com
'
'       Description : an easy to use program for storing,
'       retrieving images to/from access database. the images
'       are stored within database. they are retrieved in written to
'       file when retrieved to be used later.



Option Explicit

Public connString As String
Public connMain As New ADODB.Connection
Public rs As ADODB.Recordset
Public htmTitle As String
Dim pBag As PropertyBag
Dim pByteA() As Byte

Private Sub cmdBrowse_Click()
    Dim fileName As String
    
    On Error Resume Next
    browseDialog.ShowOpen
    
    fileName = browseDialog.fileName
    Picture1.Picture = LoadPicture(fileName)

End Sub

Private Sub cmdLoad_Click()
     On Error Resume Next
     pByteA = rs.Fields("emppic").Value
     Set pBag = New PropertyBag
     pBag.Contents = pByteA
     Set Picture1 = pBag.ReadProperty("MyPicture")
     txtID.Text = rs("empid")

     
     SavePicture Picture1, "C:\temp.bmp"
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rs.MoveNext
    cmdLoad_Click
    
End Sub

Private Sub cmdPrev_Click()
    On Error Resume Next
    rs.MovePrevious
    cmdLoad_Click
End Sub

Private Sub cmdSave_Click()
    
    Dim strQuery
               
    'Create propertybag
    Set pBag = New PropertyBag
    
    'Write object
    pBag.WriteProperty "MyPicture", Picture1.Picture
    
    'Fill array with binary data of pic
    pByteA = pBag.Contents
    
    'Write data to database
    'strQuery = "INSERT INTO Employees (emppic) VALUES (" & pByteA & ")"
    'connMain.Execute strQuery
    
    rs.AddNew
    rs.Fields("emppic").Value = pByteA
    rs.Update

End Sub


'Get record from db from primary key, display picture and
'save picture to file
Private Sub loadRecord(strkey As String)
    Dim strQuery As String
    Dim objRs As Recordset
    
    Set objRs = New ADODB.Recordset
    
    strQuery = "Select * from employees where empid = " & strkey
    objRs.Open strQuery, connMain
    
    On Error Resume Next
    pByteA = objRs.Fields("emppic").Value
    Set pBag = New PropertyBag
    pBag.Contents = pByteA
    Set Picture1 = pBag.ReadProperty("MyPicture")
    SavePicture Picture1, "C:\temp.bmp"
    
    objRs.Close
    
End Sub

Private Sub Command1_Click()
    loadRecord txtID.Text
End Sub

Private Sub Form_Load()

    connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\demoDB.mdb;Persist Security Info=False;"
    connMain.Open connString
    
    Set rs = New ADODB.Recordset
    rs.Open "Employees", connMain, adOpenKeyset, adLockPessimistic, adCmdTable
    
    cmdLoad_Click
End Sub

