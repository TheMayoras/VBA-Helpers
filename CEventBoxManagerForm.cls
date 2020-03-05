VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEventBoxManagerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name: CEventBoxManager
' Kind: Class Module
' Purpose: Mimics a "Dialog Box".  Unlike a dialog box, this class fires an event when the form is closed.  NOTE: It does not pause execution in the main routine.
' Author: MayoraB
' Date: 5/16/2019
' ------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Forms.SupportCode")
Private Type TEventBoxManager
    fieldName As String
    FieldValue As String
    setModal As Boolean
End Type


'@MemberAttribute VB_VarHelpID, -1
Private WithEvents m_objDialogForm As Form
Attribute m_objDialogForm.VB_VarHelpID = -1
Private this      As TEventBoxManager

Private Const EVENT_PROC As String = "[Event Procedure]"


'''''''''''''''''
'EVENT'
'''''''''''''''''
Public Event FormClosed(ByVal objDialogForm As CEventBoxManagerForm, ByRef Cancel As Boolean)
Public Event FormAfterClose()


''''''''''''''''''''''''''''''''''
'CONSTRUCTOR'
''''''''''''''''''''''''''''''''''
Public Sub initializeProperties(ByRef formClass As Form, Optional ByVal fieldName As String, Optional ByVal setModal As Boolean = True)
    this.fieldName = fieldName
    
    Set m_objDialogForm = formClass
    m_objDialogForm.Visible = False

    this.setModal = setModal

End Sub


''''''''''''''''''''''''''''''''''
'DESTRUCTOR'
''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    On Error Resume Next
    Me.closeForm
    DoCmd.Close acForm, m_objDialogForm.name
    
    On Error GoTo 0
End Sub


' ----------------------------------------------------------------
' Procedure Name: fieldName
' Purpose: the name of the field that FieldValue uses
' Procedure Kind: Property (Let)
' Procedure Access: Public
' Parameter field (String): the field name
' Author: MayoraB
' Date: 5/16/2019
' ----------------------------------------------------------------
Public Property Let fieldName(ByVal field As String)
    this.fieldName = field
End Property


' ----------------------------------------------------------------
' Procedure Name: FieldValue
' Purpose: the value of the field specified in the constructor
' Procedure Kind: Property (Get)
' Procedure Access: Public
' Return Type: Variant
' Author: MayoraB
' Date: 5/16/2019
' ----------------------------------------------------------------
Public Property Get FieldValue() As Variant
    FieldValue = this.FieldValue
End Property


' ----------------------------------------------------------------
' Procedure Name: objDialogForm
' Purpose: returns the form
' Procedure Kind: Property (Get)
' Procedure Access: Public
' Return Type: Form
' Author: MayoraB
' Date: 5/16/2019
' ----------------------------------------------------------------
Public Property Get DialogForm() As Form
    Set DialogForm = m_objDialogForm
End Property


'''''''''''''''''''''''''''''''''
'''PROPERTIES'
'''''''''''''''''''''''''''''''''
Public Property Set DialogForm(ByVal DialogForm As Form)
    Set m_objDialogForm = DialogForm
End Property


Public Property Let Modal(ByVal setModal As Boolean)
    this.setModal = setModal
    
    If m_objDialogForm.Visible Then
        m_objDialogForm.Modal = setModal
    End If
End Property


Public Property Let Visible(ByVal isVisible As Boolean)
    m_objDialogForm.Visible = isVisible
    
    
    If Not isVisible Then m_objDialogForm.Modal = False        ' make sure the form isn't modal.  If this is hidden with Modal = True it can cause Access to appear as though it froze
    If isVisible Then m_objDialogForm.Modal = this.setModal
End Property


'''''''''''''''''''''''
'METHODS'
'''''''''''''''''''''''
Public Sub launch(Optional ByVal openArgs As String)
    
    On Error Resume Next        ' Some of the forms passed into this class will not have the necessary parts below.  On Error Resume Next skips those steps
    With m_objDialogForm
        
        '@Ignore MemberNotOnInterface
        .txtOpenArgs = openArgs        ' gives the supplied arguments to the text box
        '@Ignore MemberNotOnInterface
        .OpenArguments = openArgs
        '@Ignore MemberNotOnInterface
        .ParseArgs        ' calls parseArgs on the form.  This is useful when the form must extract information from the opening arguments
        
        .OnUnload = EVENT_PROC        ' Set the OnUnload event to say "[Event Procedure]".  If it does not say this, we cannot actually capture the Unload Event
        .OnClose = EVENT_PROC        ' Set the OnClose event to say "[Event Procedure]".  If it does not say this, we cannot actually capture the Close Event
        
        If this.setModal Then .Modal = True        ' changes the Modal property to true if it was specified.  If not supplied, it will default to the form's modal property
        
        .Visible = True        ' make the form visible
        .SetFocus        ' make it the current form
        
    End With
    
    On Error GoTo 0
        
End Sub


' ----------------------------------------------------------------
' Procedure Name: m_objDialogForm_Close
' Purpose: Raises the FormClosed event and passes a ByRef cancel and a copy of this object
' Procedure Kind: Sub
' Procedure Access: Private
' Author: MayoraB
' Date: 5/17/2019
' ----------------------------------------------------------------
Private Sub m_objDialogForm_Close()
    RaiseEvent FormAfterClose

    Set m_objDialogForm = Nothing

End Sub


'''''''''''''''''''''''''''''''''
'FORM EVENTS
'''''''''''''''''''''''''''''''''
Private Sub m_objDialogForm_Unload(Cancel As Integer)
    On Error Resume Next
    this.FieldValue = m_objDialogForm.controls(this.fieldName).value        ' get the desired fieldvalue from the form
    
    On Error GoTo 0
    Dim bEventCancel As Boolean: bEventCancel = False        ' initialize the event's cancel option
    RaiseEvent FormClosed(Me, bEventCancel)        ' raise the event
    
    Cancel = bEventCancel        ' pass the event's cancel to the actual form's cancel
    
End Sub


Public Sub closeForm()
    On Error Resume Next
    Me.Visible = False        ' hide the form
    Me.Modal = False
    
    On Error GoTo 0
End Sub


