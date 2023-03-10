VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormSendMail 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Select attachment"
      Height          =   495
      Left            =   5400
      TabIndex        =   17
      Top             =   4680
      Width           =   2175
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
      Width           =   4935
   End
   Begin VB.TextBox txtFrom 
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   1440
      Width           =   4935
   End
   Begin VB.TextBox txtSubject 
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   840
      Width           =   4935
   End
   Begin VB.TextBox txtRecipient 
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   240
      Width           =   4935
   End
   Begin VB.TextBox txtSmtpServer 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox TextBody 
      Height          =   1815
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmSendMail.frx":0000
      Top             =   2040
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEND MAIL"
      Height          =   495
      Left            =   7800
      TabIndex        =   0
      Top             =   4680
      Width           =   2535
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
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "From"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Subject"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Recipient"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "SMTP Server"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
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
    
    
'username'password'server'recipient'subject'from'message'file-attachment

Private Function MailSend(xUsername, xPassword, xServer, xMailTo, xSubject, xFrom, xMainText, xFilepath)
    Set msgA = CreateObject("CDO.Message") 'set the CDO to reffer as.
    
    
    'msgA.sender = "tom@bretthq.com"
    msgA.From = xFrom
    msgA.To = xMailTo 'get targeted mail from command
    msgA.Subject = xSubject 'get subject from command
    msgA.HTMLBody = xMainText 'Main Text - You may use HTML tags here, for example <BR> to immitate "VBCRLF" (start new line) etc.
    ' HTMLBODY is a STRING, do not try to link a multilined textbox to it without using the ''replace'' function for 'VBCRLf' with '<BR>' (example later)
    
    'Notice, i simplified it, however, you may use more values depending on your needs, such as:
    '.Bcc = "mail@mail.com" ' - BCC..
    '.Cc = "mail@mail.com" ' - CC..
    '.From
    '.CreateMHTMLBody ("www.mywebsite.com/index.html) 'send an entire webpage from a site
    '.CreateMHTMLBody ("c:\program files\download.htm) 'Send an entire webpage from your PC
    '.AddAttachment ("c:\myfile.zip") 'Send a file from your pc (notice uploading may take a while depending on your connection)
        
     If Dir(xFilepath) <> "" Then
      MsgBox "File exists"
    
        
     If Trim$(xFilepath) <> vbNullString Then
            msgA.AddAttachment (xFilepath)
        End If
          
    End If
        
        
    'Gmail Username (from which mail will be sent)
    msgA.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = xUsername
    'Gmail Password
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
    MailSend = Err.Number
    If Err.Number <> 0 Then MsgBox "Mail Error: " & Err.Description
 
End Function

Private Sub Command1_Click()


'Usage: GmailSend ("USERNAME","PASSWORD","SendTo@mail.com","Subject","Text Body <br> New line Here"
'As i was saying, to multiline textbox won't work here. so you'll have to use the Replace function BEFORE sending the mail.


sFilePath = "C:\my-source\sfilepath.txt"

Replace TextBody.Text, vbCrLf, "<br>"

'username'password'server'recipient'subject'from'message'file-attachment

If MailSend(txtUserName.Text, txtPassword.Text, txtSmtpServer.Text, txtRecipient.Text, txtSubject.Text, txtFrom.Text, TextBody.Text, txtFileAttachment.Text) = 0 Then
    
    MsgBox "Mail Sent!"

Else

    MsgBox "Mail Send Error!"
End If


End Sub

Private Sub Command2_Click()
CommonDialog1.Filter = "All files (*.*)|*.*"
CommonDialog1.DefaultExt = "*.*"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen

'The FileName property gives you the variable you need to use
txtFileAttachment.Text = CommonDialog1.FileName
End Sub

