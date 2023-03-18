VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormSendMail 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Remember"
      Height          =   375
      Left            =   9840
      TabIndex        =   26
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Remember"
      Height          =   375
      Left            =   9840
      TabIndex        =   25
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Remember"
      Height          =   375
      Left            =   9840
      TabIndex        =   24
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Remember"
      Height          =   375
      Left            =   9840
      TabIndex        =   23
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remember"
      Height          =   375
      Left            =   9840
      TabIndex        =   22
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remember details"
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   240
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show/hide password "
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1200
      Top             =   4320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select attachment"
      Height          =   495
      Left            =   5400
      TabIndex        =   17
      Top             =   4680
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFileAttachment 
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   4080
      Width           =   4335
   End
   Begin VB.TextBox txtFrom 
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   1440
      Width           =   4335
   End
   Begin VB.TextBox txtSubject 
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox txtRecipient 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   240
      Width           =   4335
   End
   Begin VB.TextBox txtSmtpServer 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox TextBody 
      Height          =   1815
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmSendMail.frx":0000
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEND MAIL"
      Height          =   495
      Left            =   7680
      TabIndex        =   0
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "SMTP SERVER DETAILS :"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "Processing, please wait ..."
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
      Left            =   360
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label8 
      Caption         =   "Attachment"
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Message body"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "From"
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Subject"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Recipient"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "SMTP Server"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "FormSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msgA As Object 'Define the CDO
Dim sFilePath As String


Private Sub DrawRadialLines()

    Const ksngPI            As Single = 3.14159!
    Const ksngCircle        As Single = 2! * ksngPI


    'original params
    'Const ksngInnerRadius   As Single = 130!
    'Const ksngOuterRadius   As Single = 260!


    Const ksngInnerRadius   As Single = 260!
    Const ksngOuterRadius   As Single = 520!
    
    'positioning params
    Const ksngCenterX       As Single = 1800!
    Const ksngCenterY       As Single = 4200!

    

    Const klSegmentCount    As Long = 12

    Const klLineWidth       As Long = 3

    Static s_lActiveSegment As Integer              ' The "selected" segment.

    Dim lSegment            As Long
    Dim sngRadians          As Single
    Dim sngX1               As Single
    Dim sngY1               As Single
    Dim sngX2               As Single
    Dim sngY2               As Single
    Dim cLineColour         As OLE_COLOR

    Me.DrawWidth = klLineWidth

    ' Overdraw previous graphic.
    Me.Line (ksngCenterX - ksngOuterRadius - Screen.TwipsPerPixelX * 2, ksngCenterY - ksngOuterRadius - Screen.TwipsPerPixelY * 2)-(ksngCenterX + ksngOuterRadius + Screen.TwipsPerPixelX * 2, ksngCenterY + ksngOuterRadius + Screen.TwipsPerPixelY * 2), Me.BackColor, BF

    For lSegment = 0 To klSegmentCount - 1

        '
        ' Work out the coordinates for the line to be draw from the outside circle to the inside circle.
        '

        sngRadians = (ksngCircle * CSng(lSegment)) / klSegmentCount

        sngX1 = (ksngOuterRadius * Cos(sngRadians)) + ksngCenterX
        sngY1 = (ksngOuterRadius * Sin(sngRadians)) + ksngCenterY
        sngX2 = (ksngInnerRadius * Cos(sngRadians)) + ksngCenterX
        sngY2 = (ksngInnerRadius * Sin(sngRadians)) + ksngCenterY

        ' Work out how many segments away from the "current segment" we are.
        ' The current segment should be the darkest, and the further away from this segment we are, the lighter the colour should be.
        Select Case Abs(Abs(s_lActiveSegment - lSegment) - klSegmentCount \ 2)
        Case 0!
            'cLineColour = RGB(0, 0, 255)
            'original blue color
            
            cLineColour = RGB(0, 0, 0)
            'black
            
        Case 1!
            'cLineColour = RGB(63, 63, 255)
            cLineColour = &H808080
            
        Case 2!
            'cLineColour = RGB(117, 117, 255)
             cLineColour = &HC0C0C0
        Case Else
            cLineColour = RGB(181, 181, 255)
            'cLineColour = &H808080
            cLineColour = &HE0E0E0
        End Select

        Me.Line (sngX1, sngY1)-(sngX2, sngY2), cLineColour
        'Picture1.Line (sngX1, sngY1)-(sngX2, sngY2), cLineColour
        'Image1.Line (sngX1, sngY1)-(sngX2, sngY2), cLineColour

    Next lSegment

    ' Move the current segment on by one.
    s_lActiveSegment = (s_lActiveSegment + 1) Mod klSegmentCount

End Sub

    
    
'username'password'server'recipient'subject'from'message'file-attachment

Private Function MailSend(xUsername, xPassword, xServer, xMailTo, xSubject, xFrom, xMainText, xFilepath)

On Error GoTo errorsub

    Set msgA = CreateObject("CDO.Message") 'set the CDO to reffer as.
    
    
    msgA.From = xFrom
    msgA.To = xMailTo 'get targeted mail from command
    msgA.Subject = xSubject 'get subject from command
    msgA.HTMLBody = xMainText 'Main Text - You may use HTML tags here, for example <BR> to immitate "VBCRLF" (start new line) etc.
    ' HTMLBODY is a STRING, do not try to link a multilined textbox to it without using the ''replace'' function for 'VBCRLf' with '<BR>' (example later)
    
    If xFilepath <> "" Then
    
     If Dir(xFilepath) <> "" Then
      MsgBox "File exists! adding"
    
        
        If Trim$(xFilepath) <> vbNullString Then
                msgA.AddAttachment (xFilepath)
        End If
          
    End If
    
    End If
    
        
        
    'mail Username (from which mail will be sent)
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = xUsername
    'mail Password
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = xPassword
    
    'Mail Server address.
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = xServer
    
    'To set SMTP over the network = 2
    'To set Local SMTP = 1
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    
    'Type of Authenthication
    '0 - None
    '1 - Base 64 encoded (Normal)
    '2 - NTLM
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    
    'Port
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    
    'Send using SSL True\False
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    
    'Update values of the SMTP configuration
    msgA.Configuration.Fields.Update
    
    'Send it.
    msgA.Send
    
    
    
    
    ' On error ... (using on error resume next is over rated, realy.)
      
    
    
    'If Err.Number <> 0 Then MsgBox "Mail Error: " & Err.Description
 
errorsub:
If Err.Number <> 0 Then

MsgBox "Mail Error: " & Err.Description & vbCrLf & "Error number : " & Err.Number
MailSend = Err.Number


End If


End Function

Private Sub Check1_Click()

If Check1.Value = 1 Then
    txtPassword.PasswordChar = ""
Else
    txtPassword.PasswordChar = "*"
End If

End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
Label9.Visible = True


'sFilePath = "C:\my-source\sfilepath.txt"

Replace TextBody.Text, vbCrLf, "<br>"

'usage
'username'password'server'recipient'subject'from'message'file-attachment

If MailSend(txtUserName.Text, txtPassword.Text, txtSmtpServer.Text, txtRecipient.Text, txtSubject.Text, txtFrom.Text, TextBody.Text, txtFileAttachment.Text) = 0 Then
    
    MsgBox "Mail Sent!"

Else

    'MsgBox "Mail Send Error!"
End If

Me.Cls

Label9.Visible = False
Timer1.Enabled = False



End Sub

Private Sub Command2_Click()
CommonDialog1.Filter = "All files (*.*)|*.*"
CommonDialog1.DefaultExt = "*.*"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen

'The FileName property gives you the variable you need to use
txtFileAttachment.Text = CommonDialog1.FileName
End Sub

Private Sub Command3_Click()
Call SaveSetting("sendmail", "configs", "username", txtUserName.Text)
Call SaveSetting("sendmail", "configs", "password", txtPassword.Text)
Call SaveSetting("sendmail", "configs", "smtpserver", txtSmtpServer.Text)

MsgBox "details saved"



End Sub

Private Sub Command4_Click()
Call SaveSetting("sendmail", "configs", "recipient", txtRecipient.Text)
MsgBox "details saved"

End Sub

Private Sub Command5_Click()
Call SaveSetting("sendmail", "configs", "subject", txtSubject.Text)
MsgBox "details saved"

End Sub

Private Sub Command6_Click()
Call SaveSetting("sendmail", "configs", "from", txtFrom.Text)
MsgBox "details saved"

End Sub

Private Sub Command7_Click()
Call SaveSetting("sendmail", "configs", "message", TextBody.Text)
MsgBox "details saved"

End Sub

Private Sub Command8_Click()
Call SaveSetting("sendmail", "configs", "attachment", txtFileAttachment.Text)
MsgBox "details saved"

End Sub

Private Sub Form_Load()
txtUserName.Text = GetSetting("sendmail", "configs", "username")
txtPassword.Text = GetSetting("sendmail", "configs", "password")
txtSmtpServer.Text = GetSetting("sendmail", "configs", "smtpserver")
txtRecipient.Text = GetSetting("sendmail", "configs", "recipient")
txtSubject.Text = GetSetting("sendmail", "configs", "subject")
txtFrom.Text = GetSetting("sendmail", "configs", "from")
TextBody.Text = GetSetting("sendmail", "configs", "message")
txtFileAttachment.Text = GetSetting("sendmail", "configs", "attachment")


End Sub

Private Sub Timer1_Timer()
DrawRadialLines

End Sub

