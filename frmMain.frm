VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BioDB [A collection of biographies]"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Height          =   375
      Left            =   9840
      Picture         =   "frmMain.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "About"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdCopy 
      DownPicture     =   "frmMain.frx":0544
      Height          =   375
      Left            =   9840
      Picture         =   "frmMain.frx":0986
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Copy To Clipboard"
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.ListBox lstAuthors 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   6300
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Click for biography"
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtBio 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "frmMain.frx":0EB8
      ToolTipText     =   "Author Biography"
      Top             =   2520
      Width           =   5055
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   11033
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Classical Collection"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "BioCount: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Died:"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Born:"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblDied 
      BackStyle       =   0  'Transparent
      Caption         =   "Death"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   1800
      Width           =   5775
   End
   Begin VB.Label lblBorn 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of birth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   735
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0E0FF&
      BorderWidth     =   15
      Height          =   4575
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   220
      Left            =   4200
      Top             =   1560
      Width           =   5775
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFC0C0&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   4200
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Classical Collection"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  BioDB Biography Database
'  Just a database for your literature class
'=========================================================================================
'  Created By: Behrooz Sangani <bs20014@yahoo.com>
'  Published Date: 31/05/2002
'  WebSite: http://www.geocities.com/bs20014
'           Visit my web site for more literature...
'  Legal Copyright: Behrooz Sangani © 31/05/2002
'  Copyright DOES NOT include Database
'=========================================================================================

Dim db As Database
Dim rs As Recordset
Dim strSQL As String
'=========================================================================================
Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub 'cmdAbout_Click()
'=========================================================================================
Private Sub cmdCopy_Click()
    Dim txt As String
    txt = lblName.Caption & vbCrLf & vbCrLf & "Born: " & _
        lblBorn.Caption & vbCrLf & "Died: " & lblDied.Caption & _
        vbCrLf & vbCrLf & txtBio
    Clipboard.SetText txt
    lstAuthors.SetFocus
End Sub 'cmdCopy_Click()
'=========================================================================================
Private Sub Form_Load()
    Me.Show
    DoEvents
    LoadAuthors
End Sub 'Form_Load()
'=========================================================================================
Sub LoadAuthors()
    PB1.Visible = True
    Set db = OpenDatabase(AppPath & "authors.mdb")
    strSQL = "select * from bio"
    Set rs = db.OpenRecordset(strSQL)
    rs.MoveLast
    rs.MoveFirst
    PB1.Max = rs.RecordCount
    lstAuthors.Clear
    For i = 1 To rs.RecordCount
        lstAuthors.AddItem rs.Fields("Name")
        rs.MoveNext
        PB1 = PB1 + 1
    Next i
    rs.MoveFirst
    PB1 = 0
    PB1.Visible = False
    lblCount.Caption = lblCount.Caption & rs.RecordCount
    lstAuthors.Selected(0) = True
    LoadAuthor lstAuthors.Text
End Sub 'LoadAuthors()
'=========================================================================================
Sub LoadAuthor(AuthorName As String)
    strSQL = "select * from bio where " & "Name" & " like '" & AuthorName & "'"
    Set rs = db.OpenRecordset(strSQL)
    lblName.Caption = rs.Fields("Name")
    lblBorn.Caption = rs.Fields("Born")
    lblDied.Caption = rs.Fields("Died")
    txtBio.Text = rs.Fields("Biography")
End Sub 'LoadAuthor(AuthorName As String)
'=========================================================================================
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = &HFF&
End Sub 'Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'=========================================================================================
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set rs = Nothing
    Set db = Nothing
End Sub 'Form_Unload(Cancel As Integer)
'=========================================================================================
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = &H8080FF
End Sub 'Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'=========================================================================================
Private Sub lstAuthors_Click()
    On Error Resume Next
    LoadAuthor lstAuthors.Text
End Sub 'lstAuthors_Click()
'=========================================================================================
Function AppPath() As String
    Dim sAns As String
    sAns = App.Path
    If Right(App.Path, 1) <> "\" Then sAns = sAns & "\"
    AppPath = sAns
End Function 'AppPath() As String
'=========================================================================================
Private Sub lstAuthors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = &HFF&
End Sub 'lstAuthors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'=========================================================================================
