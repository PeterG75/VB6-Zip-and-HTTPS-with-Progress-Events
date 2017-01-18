VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Abort"
      Height          =   315
      Left            =   9480
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   1560
      Width           =   10815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Step 3: Re-zip"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Step 2: Unzip"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Step 1: Download a .zip for Testing"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim WithEvents myHttp As ChilkatHttp
Attribute myHttp.VB_VarHelpID = -1
Dim WithEvents myZip As ChilkatZip
Attribute myZip.VB_VarHelpID = -1


' Download a .zip
Private Sub Command1_Click()
    Dim success As Long
    
    Dim fac As New CkFileAccess
    success = fac.DirEnsureExists("c:/aaworkarea")
        
    Dim url As String
    url = "http://chilkatdownload.com/9.5.0.65/chilkatax-9.5.0-win32.zip"
    
    Text1.Text = ""
    ProgressBar1.value = 0
    
    myHttp.HeartbeatMs = 50
    success = myHttp.Download(url, "c:/aaworkarea/ChilkatActiveX.zip")
    If (success <> 1) Then
        Text1.Text = Text1.Text & vbCrLf & myHttp.LastErrorText
    Else
        MsgBox "Success."
    End If
    
    
End Sub

' Unzip what was downloaded in Step 1.
Private Sub Command2_Click()

    myZip.HeartbeatMs = 50
    
    Text1.Text = ""
    ProgressBar1.value = 0
    
    Dim success As Long
    
    Dim fac As New CkFileAccess
    success = fac.DirEnsureExists("c:/aaworkarea/temp")
    
    success = myZip.OpenZip("c:/aaworkarea/ChilkatActiveX.zip")
    If (success <> 1) Then
        Text1.Text = myZip.LastErrorText
        Exit Sub
    End If
    
    Dim count As Long
    count = myZip.Unzip("c:/aaworkarea/temp")
    If (count < 0) Then
        Text1.Text = myZip.LastErrorText
        Exit Sub
    End If
    
    myZip.CloseZip
    
End Sub

' Re-zip what was unzipped in Step 2.
Private Sub Command3_Click()

    myZip.HeartbeatMs = 50
    
    Text1.Text = ""
    ProgressBar1.value = 0
    
    Dim success As Long
    success = myZip.NewZip("c:/aaworkarea/MyNewZip.zip")
    
    recurse = 1
    success = myZip.AppendFiles("c:/aaworkarea/temp", recurse)
    If (success <> 1) Then
        Text1.Text = myZip.LastErrorText
        Exit Sub
    End If
    
    success = myZip.WriteZipAndClose()
    If (success <> 1) Then
        Text1.Text = myZip.LastErrorText
        Exit Sub
    End If
    
End Sub

Private Sub Form_Load()
    Set myHttp = New ChilkatHttp
    success = myHttp.UnlockComponent("Anything for 30-day trial")
    Set myZip = New ChilkatZip
    success = myZip.UnlockComponent("Anything for 30-day trial")
    
End Sub

Private Sub myHttp_AbortCheck(abort As Long)
    DoEvents
End Sub

Private Sub myHttp_PercentDone(ByVal percent As Long, abort As Long)
    ProgressBar1.value = percent
End Sub

Private Sub myHttp_ProgressInfo(ByVal name As String, ByVal value As String)
  Text1.Text = Text1.Text & vbCrLf & name & ": " & value
End Sub

Private Sub myZip_AbortCheck(abort As Long)
    DoEvents
End Sub


Private Sub myZip_PercentDone(ByVal percent As Long, abort As Long)
    ProgressBar1.value = percent
End Sub

Private Sub myZip_ProgressInfo(ByVal name As String, ByVal value As String)
  Text1.Text = Text1.Text & vbCrLf & name & ": " & value
End Sub
