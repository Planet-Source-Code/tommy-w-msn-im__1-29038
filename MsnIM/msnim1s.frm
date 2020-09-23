VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Mini IM"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   3870
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   0
      Width           =   4815
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         DragIcon        =   "msnim1s.frx":0000
         DragMode        =   1  'Automatic
         Height          =   300
         Left            =   4365
         Picture         =   "msnim1s.frx":0CCA
         ScaleHeight     =   300
         ScaleWidth      =   240
         TabIndex        =   9
         Top             =   90
         Width           =   240
      End
      Begin VB.ComboBox cmbUsers 
         Height          =   315
         Left            =   255
         TabIndex        =   7
         Top             =   75
         Width           =   2505
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set"
         Height          =   315
         Left            =   3675
         TabIndex        =   6
         Top             =   90
         Width           =   660
      End
      Begin VB.CommandButton Command1 
         Caption         =   "del"
         Height          =   315
         Left            =   2835
         TabIndex        =   5
         Top             =   90
         Width           =   345
      End
      Begin VB.CommandButton Command2 
         Caption         =   "add"
         Height          =   315
         Left            =   3210
         TabIndex        =   4
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox txtLog 
      Height          =   2265
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "msnim1s.frx":1054
      Top             =   570
      Width           =   4515
   End
   Begin VB.TextBox txtSend 
      Height          =   705
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2895
      Width           =   4530
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      Caption         =   "Created by Thomasunt Products (homeworkkid@msn.com)"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3615
      Width           =   4785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   This code simulates an array of IM Windows
'in MSN Messenger. Whoever's email you have set
'will be the active receiver, sender. In other
'words, if you had blah@hotmail.com set, all
'all messages sent will go to that address and
'all messages received will be filtered to only
'ones from that address. You may use this code
'to gain knowledge of the Messenger 1.0 Type
'Library or any other knowledge acquirable.

'Feel free to use this code for whatever you wish
'   Thomasunt Products:
'       Because free is the way it should be

Public WithEvents mm As MsgrObject 'Main MSN control
Attribute mm.VB_VarHelpID = -1
Dim bstrMsgHeader As String 'Message Header
Public User1 As IMsgrUser 'The User Obj
Public User1S As IMsgrIMSession 'The User Session Obj
Public Users As IMsgrUsers 'The User Array (Contact List)
Public Service As IMsgrService 'The MSN Service
Dim ctrldown As Boolean 'If ctrl is down in txtsend

Private Sub cmbUsers_Click()
Set User1 = mm.CreateUser(cmbUsers.Text, mm.Services.PrimaryService) 'Set User Obj
txtLog = txtLog & vbCrLf & "Switched to talk to " & User1.EmailAddress & ":" 'Notify change in txtlog
End Sub

Private Sub cmbUsers_DragDrop(Source As Control, X As Single, Y As Single)
If Source = Picture1 Then MsgBox "This combo box sets the e-mail object once a new e-mail is clicked or when you click set."
End Sub

Private Sub cmdSet_Click()
On Error Resume Next
Dim i As Integer
Dim add As Boolean
Set User1 = mm.CreateUser(cmbUsers.Text, mm.Services.PrimaryService)
For i = 0 To cmbUsers.ListCount - 1 'This section checks whether the e-mail is already in the list or not
If cmbUsers.List(i) = cmbUsers.Text Then add = True
Next i
If add = False Then cmbUsers.AddItem cmbUsers.Text
txtLog = txtLog & vbCrLf & "Switched to talk to " & User1.EmailAddress & ":"
End Sub

Private Sub cmdSet_DragDrop(Source As Control, X As Single, Y As Single)
If Source = Picture1 Then MsgBox "This is the Set button, it sets the object for the e-mail in the combo box."
End Sub

Private Sub Command1_Click()
On Error Resume Next
cmbUsers.RemoveItem cmbUsers.ListIndex 'destroys selected item
End Sub

Private Sub Command1_DragDrop(Source As Control, X As Single, Y As Single)
If Source = Picture1 Then MsgBox "This button deletes the currently selected item of the combo box."
End Sub

Private Sub Command2_Click()
On Error Resume Next
cmbUsers.AddItem cmbUsers.Text 'adds current e-mail (done when set is clicked)
End Sub

Private Sub Command2_DragDrop(Source As Control, X As Single, Y As Single)
If Source = Picture1 Then MsgBox "This button adds the current e-mail to the list (also done when object is set)"
End Sub

Private Sub Form_Load()
On Error Resume Next
frmSplash.Show 'Show splash screen
Set mm = New MsgrObject
''bstrMsgHeader = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=MS%20Sans%20Serif; EF=; CO=0; CS=0; PF=0"
''User1.SendText bstrMsgHeader, "Hi! Test..", MMSGTYPE_ALL_RESULTS
'User1S.SendText bstrMsgHeader, "If you get this message, say Hi", MMSGTYPE_ALL_RESULTS
Dim i As Integer
Set Users = mm.List(MLIST_CONTACT)
frmSplash.pb1.Max = Users.Count - 1
For i = 0 To Users.Count - 1
frmSplash.pb1.Value = i 'Increase progress bar
'scans contact list
If Users.Item(i).State <> MSTATE_INVISIBLE Then
If Users.Item(i).State <> MSTATE_OFFLINE Then
If Users.Item(i).State <> MSTATE_UNKNOWN Then
'if online, add to combo box
cmbUsers.AddItem Users.Item(i).EmailAddress
End If
End If
End If
Next i
Unload frmSplash
cmbUsers.ListIndex = 0
End Sub

Public Function SendTheText()
On Error Resume Next
Set User1S = mm.CreateIMSession(User1)
User1S.SendText bstrMsgHeader, txtSend, MMSGTYPE_ERRORS_ONLY
txtLog = txtLog & vbCrLf & mm.LocalFriendlyName & ": " & txtSend
lstmsg = txtSend
txtSend = ""
End Function

Private Sub Form_Resize()
'Form Resize Settings
lblAuthor.Top = Me.Height - (lblAuthor.Height * 2.5) - 300
lblAuthor.Width = Me.Width
txtSend.Top = Me.Height - (txtSend.Height * 2) - 400
txtLog.Height = Me.Height - 2400
txtLog.Width = Me.Width - 200
txtSend.Width = Me.Width - 200
If Me.Height < 4365 Then Me.Height = 4365
If Me.Width < 4800 Then Me.Width = 4800
End Sub

Private Sub mm_OnSendResult(ByVal hr As Long, ByVal lCookie As Long)
On Error Resume Next
txtLog = txtLog & vbCrLf & "Error sending last message." 'on send error
End Sub

Private Sub mm_OnTextReceived(ByVal pIMSession As Messenger.IMsgrIMSession, ByVal pSourceUser As Messenger.IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean)
On Error Resume Next
'If the email is that of the user you are talking to, add their text to the log
If pSourceUser.EmailAddress = User1.EmailAddress And bstrMsgText <> vbCrLf Then txtLog = txtLog & vbCrLf & pSourceUser.FriendlyName & ": " & bstrMsgText
End Sub

Private Sub txtLog_Change()
On Error Resume Next
txtLog.SelStart = Len(txtLog) 'Keeps txtlog caught up
End Sub

Private Sub txtLog_DragDrop(Source As Control, X As Single, Y As Single)
If Source = Picture1 Then MsgBox "This is the text log box. The log for the chat is here. Note you can't copy text without right-clicking text and clicking copy (not ctrl & c)"
End Sub

Private Sub txtSend_Change()
On Error Resume Next
If txtSend = vbCrLf Then txtSend = "" 'fixes problem of new line after send
End Sub

Private Sub txtSend_DragDrop(Source As Control, X As Single, Y As Single)
If Source = Picture1 Then MsgBox "This is the text send box. Text in this box is text that will be sent."
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 17 Then ctrldown = True 'if ctrl is down
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 And ctrldown = False Then SendTheText 'if ctrl isn't down when enter is pressed, send the text
End Sub

Private Sub txtSend_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 17 Then ctrldown = False 'if ctrl is up
End Sub
