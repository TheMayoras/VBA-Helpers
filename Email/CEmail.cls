VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' A wrapper around an Outlook MailItem.
Option Compare Database
Option Explicit

'@Folder("Email.Wrapper")

Private m_objMailItem As Object
Private m_objFso  As FileSystemObject
Private m_sCatalogNumber As String
Private m_sRecipients As String
'@Ignore UseMeaningfulName
Private m_sCC     As String


'''''''''''''''''''''''''''''''''''
'CONSTRUCTOR'
'''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    ' the fso is used to get any attachments that are added
    Set m_objFso = New FileSystemObject
End Sub


''''''''''''''''''''''''''''''''
'DESTRUCTOR'
''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    If Not (m_objMailItem Is Nothing) Then        ' delete the mailitem if it hasn't been sent
        m_objMailItem.Delete
        Set m_objMailItem = Nothing
    End If
    Set m_objFso = Nothing
End Sub

Public Property Get Attachments() As Collection
    Dim objItem   As Variant
    Dim objItemColl As Collection: Set objItemColl = New Collection

    For Each objItem In m_objMailItem.Attachments
        objItemColl.Add objItem.PathName & objItem.FileName
    Next objItem
    
    Set Attachments = objItemColl
End Property

Public Property Get Body() As String
    On Error Resume Next
    Body = m_objMailItem.HtmlBody
End Property


Public Property Let Body(ByVal Body As String)
    On Error Resume Next
    m_objMailItem.HtmlBody = Replace(Body, vbNewLine, "<br>")
End Property


Public Property Let HtmlBody(ByVal Body As String)
    m_objMailItem.HtmlBody = Body
End Property

Public Property Let CatalogNumber(ByVal CatalogNumber As String)
    m_sCatalogNumber = CatalogNumber
End Property

Public Property Get cc() As String
    On Error Resume Next
    cc = m_sCC
End Property

'@Ignore UseMeaningfulName
Public Property Let cc(ByVal cc As String)
    On Error Resume Next
    m_sCC = cc
End Property

Public Property Get IsNothing() As Boolean
    On Error Resume Next
    IsNothing = (m_objMailItem Is Nothing)
End Property


' set the MailItem we are wrapping around
Public Property Set MailItem(ByVal item As Object)
    If IsNothing Then
        Set m_objMailItem = item
    Else        ' Cannot overwrite a MailItem
        Err.Raise vbObjectError + 515, source:="public property set MailItem", _
                  Description:="The current MailItem is already set."
    End If
End Property

Public Property Get Recipients() As String
    On Error Resume Next
    Recipients = m_sRecipients
End Property

Private Property Let Recipients(ByVal Recipients As String)
    On Error Resume Next
    m_sRecipients = Recipients
End Property

Public Property Get Subject() As String
    On Error Resume Next
    Subject = m_objMailItem.Subject
End Property

Public Property Let Subject(ByVal Subject As String)
    On Error Resume Next
    m_objMailItem.Subject = Subject
End Property


''''''''''''''''''''
'METHODS'
''''''''''''''''''''
Public Function addAttachment(ByVal filePath As String, Optional ByVal displayName As String = vbNullString) As Boolean
    addAttachment = False
    If (m_objFso.FileExists(filePath)) Then
        addAttachment = True
        If displayName = vbNullString Then
            m_objMailItem.Attachments.Add filePath
        Else
            m_objMailItem.Attachments.Add filePath, displayName:=displayName
        End If
    End If
End Function

Public Sub AddRecipient(ByVal recipient As String)
    Recipients = Recipients & recipient & ";"
End Sub

Public Function AttachmentNames() As String
    
    Dim sItem     As Variant
    Dim sAttachments As String

    For Each sItem In m_objMailItem.sAttachments
        sAttachments = sAttachments & ", " & sItem.PathName & sItem.FileName
    Next sItem
    
    If sAttachments = vbNullString Then
        attachmentNames = vbNullString
        
    Else
        attachmentNames = Right$(sAttachments, Len(sAttachments) - 2)
    End If
End Function

Public Sub RemoveRecipient(ByVal recipient As String)
    Recipients = Replace(Recipients, recipient, vbNullString, compare:=vbTextCompare)
End Sub

Public Sub Send()
    If m_objMailItem Is Nothing Then        ' Can't send a mail item that hasn't been set
        Err.Raise vbObjectError + 516, source:="CEmail.send", _
                  Description:="MailItem cannot be sent because it is not referencing anything"
                  
    Else
        Me.Subject = "MOC: " & Me.Subject
        
        m_objMailItem.To = m_sRecipients        ' set the recipients of the email to the recipients added
        m_objMailItem.cc = m_sCC        ' set the CC of the email to the CC's added
        
        On Error GoTo EMAIL_SEND_ERR
        
Debug.Print "Seding....  " & vbNewLine & vbTab & "To: " & m_objMailItem.To & vbNewLine & vbTab & "    " & m_objMailItem.cc & vbNewLine & vbNewLine & vbTab & m_objMailItem.Subject & vbNewLine & vbNewLine & vbTab & m_objMailItem.Body     
         m_objMailItem.send
        
        On Error GoTo 0
        Set m_objMailItem = Nothing
    End If
    
    Exit Sub
    
EMAIL_SEND_ERR:
    MsgBox "Hmmm...." & vbNewLine & vbNewLine & "There seems to have been an issue sending an email.  " & vbNewLine, "Error Sending Email"
    Set m_objMailItem = Nothing
End Sub



