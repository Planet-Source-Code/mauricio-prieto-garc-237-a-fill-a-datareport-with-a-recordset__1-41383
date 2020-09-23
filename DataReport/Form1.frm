VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fill a DataReport wiht a Recordset"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      Caption         =   "Order by Name"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Order by Number"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Disconnect"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Report"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cn As ADODB.Connection
Dim OrderKey As String
' in order to use this type of variable you must match
' the "Microsof tActivex Data Objects 2.x (i.e.,"2.0","2.6","2.7",..)"
' or others more recent placed in: Project -> References...

Private Sub Command1_Click()

On Error GoTo ErrHandler

    Set Cn = New ADODB.Connection
    Cn.ConnectionString = "DBQ=" & App.Path & "\databasename.mdb;DefaultDir=" & App.Path & ";Driver={Microsoft Access Driver (*.mdb)};DriverId=25;FIL=MS Access;FILEDSN=;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
    Cn.Open
    Label1.Caption = "Data Base State: Connected !"
    Command1.Enabled = False
    Command2.Enabled = True
    Command2.SetFocus
    Command3.Enabled = True
    Exit Sub
    
ErrHandler:
    MsgBox "The application could not connect the data base file!!!" + Chr(13) & Chr(10) + Err.Description, vbInformation, "Sorry!!!"
    Err.Clear
    End
    
End Sub

Private Sub Command2_Click()

Dim RsRC As New ADODB.Recordset

    If Option1 = True Then
        OrderKey = "ClientPrimaryKey"
    Else
        OrderKey = "ClientName"
    End If
    Set RsRC = Cn.Execute("Select ClientPrimaryKey,ClientName,ClientRepr,ClientAddress,ClientPhone,ClientCelPhone,ClientFedID,ClientDate from Client order by(" & OrderKey & ")")
    With DataReport1
        .DataMember = vbNullString
        Set .DataSource = RsRC
        .Caption = "Hello !!!" ' you can write here the DataReport window's name
        With .Sections("Sección1").Controls
            .Item("tClientPK").DataField = RsRC.Fields(0).Name
            .Item("tClientN").DataField = RsRC.Fields(1).Name
            .Item("tClientR").DataField = RsRC.Fields(2).Name
            .Item("tClientA").DataField = RsRC.Fields(3).Name
            .Item("tClientP").DataField = RsRC.Fields(4).Name
            .Item("tClientCP").DataField = RsRC.Fields(5).Name
            .Item("tClientFID").DataField = RsRC.Fields(6).Name
            .Item("tClientD").DataField = RsRC.Fields(7).Name
        End With
        .Sections("Sección4").Controls.Item("tDate").Caption = Format(Now, "dddd, dd-mmmm-yyyy")
        DataReport1.WindowState = 2
        .Show
    End With
End Sub

Private Sub Command3_Click()
    If Cn.State = adStateOpen Then
        Cn.Close
        Command1.Enabled = True
        Command1.SetFocus
        Command2.Enabled = False
        Command3.Enabled = False
        Label1.Caption = "Data Base State: Disconnected !"
    End If
End Sub

Private Sub Command4_Click()
    MsgBox "Thank's, Have a nice day !!!", vbInformation, "Mauricio say..."
    End
End Sub

Private Sub Form_Load()
    Command1.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = False
    Label1.Caption = "Data Base State: Disconnected !"
End Sub

