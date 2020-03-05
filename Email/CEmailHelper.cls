VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEmailHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Purpose: Generates CEmail objects using the current user's Outlook account
' The '@PredeclaredId means that there is always an instance of this class available.  This means that we can have a single outlook instance for all email objects
Option Compare Database
Option Explicit
'@Folder("Email.Wrapper")

'@PredeclaredId

''''''''''''''''''''
'MEMBERS'
''''''''''''''''''''
Private Const MAIL_ITEM As Integer = 0
Private m_objoutlookapp As Object
Private m_objFso  As FileSystemObject


'''''''''''''''''''''''''''''
'CONSTRUCTOR'
'''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_objoutlookapp = CreateObject("Outlook.Application")        ' make the Outlook instance
    Set m_objFso = New FileSystemObject
End Sub


''''''''''''''''''''''''''
'DESTRUCTOR'
''''''''''''''''''''''''''
Private Sub Class_Terminate()
    On Error Resume Next
    m_objoutlookapp.Close        ' close the outlook instance
    
    On Error GoTo 0
    Set m_objoutlookapp = Nothing
    Set m_objFso = Nothing
End Sub


''''''''''''''''''''
'METHODS'
''''''''''''''''''''

Public Function CreateMailItem() As CEmail

    Dim objEmail  As CEmail: Set objEmail = New CEmail
    Dim objItem   As Object: Set objItem = m_objoutlookapp.CreateItem(MAIL_ITEM)        ' Create the mail item
    
    Set objEmail.MailItem = objItem        ' set the CEmail's MailItem
    Set createMailItem = objEmail        ' Return the CEmail
End Function

Public Sub SendEmail(ByVal emailType As IEmailType)
    Dim objEmail  As CEmail: Set objEmail = Factory.createEmail        ' Create an Email Object
    
    With objEmail        ' set the email's properties from the email type
        .addRecipient emailType.Recipients
        .cc = emailType.cc
        .Subject = emailType.Subject
        .HtmlBody = "<html> "
        .HtmlBody = .Body & HtmlHead
        .HtmlBody = .Body & "<body>"
        .HtmlBody = .Body & "<div class=""content"">" & emailType.Description & "</div>"
        .HtmlBody = .Body & "</body>"
        .HtmlBody = .Body & "</html>"
        
        AddAttachments objEmail, emailType.Attachments
        
        On Error GoTo EMAIL_SEND_ERR
        .send        ' send the email.  This can fail
    End With

EXIT_SUB:
    On Error GoTo 0
    Exit Sub
    
EMAIL_SEND_ERR:            ' Give the user a message telling an email couldn't be sent and to contact the admins
    MsgBox _
        "The database attempted to send an email with the subject:" & vbNewLine & vbNewLine & _
		"""" & emailType.Subject & """" & vbNewLine & vbNewLine & _
		"There was an error sending the email.  " & vbNewLine & _
		"Please alert the admins: " & listAdmins(), _
		vbCritical Or vbOKOnly, _
		"Email Send Error"
End Sub

Private Sub AddAttachments(ByRef Email As CEmail, ByVal files As Collection)
    If files Is Nothing Then Exit Sub
    
    Dim varFilePath As Variant
    For Each varFilePath In files
        Email.addAttachment varFilePath
    Next varFilePath
End Sub
