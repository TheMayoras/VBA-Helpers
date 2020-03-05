VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEmailType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
'@Folder("Email.Interface")

Public Property Get Description() As String
End Property

Public Property Get Recipients() As String
End Property

Public Property Get Cc() As String
End Property

Public Property Get Subject() As String
End Property

Public Property Get Attachments() As Collection
End Property



