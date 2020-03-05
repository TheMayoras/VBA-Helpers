VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEventBoxManagerReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Forms.SupportCode")

' ------------------------------------------------------
' Name: clsEventBoxManager
' Kind: Class Module
' Purpose: Mimics a "Dialog Box".  Unlike a dialog box, this class fires an event when the form is closed.  NOTE: It does not pause execution in the main routine.
' Author: MayoraB
' Date: 5/16/2019
' ------------------------------------------------------
Option Compare Database
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents m_objDialogReport As Report
Attribute m_objDialogReport.VB_VarHelpID = -1
Private m_bSetModal As Boolean

Private Const EVENT_PROC As String = "[Event Procedure]"


'''''''''''''''''
'EVENT'
'''''''''''''''''
'@Ignore ParameterCanBeByVal
Public Event ReportClosed(ByVal objDialogForm As CEventBoxManagerReport, ByRef Cancel As Boolean)
Public Event ReportAfterClose()


''''''''''''''''''''''''''''''''''
'CONSTRUCTOR'
''''''''''''''''''''''''''''''''''
Public Sub initializeProperties(ByRef reportClass As Report, Optional ByVal setModal As Boolean = True)

    Set m_objDialogReport = reportClass
    reportClass.Requery
    m_objDialogReport.Visible = False

    m_bSetModal = setModal
End Sub


''''''''''''''''''''''''''''''''''
'DESTRUCTOR'
''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    On Error Resume Next
    Me.closeReport
    DoCmd.Close acReport, m_objDialogReport.name
    
    On Error GoTo 0
End Sub


' ----------------------------------------------------------------
' Procedure Name: objDialogForm
' Purpose: returns the form
' Procedure Kind: Property (Get)
' Procedure Access: Public
' Return Type: Form
' Author: MayoraB
' Date: 5/16/2019
' ----------------------------------------------------------------
Public Property Get dialogReport() As Report
    Set dialogReport = m_objDialogReport
End Property


'''''''''''''''''''''''''''''''''
'''PROPERTIES'
'''''''''''''''''''''''''''''''''
Public Property Set dialogReport(ByVal objDialogReport As Report)
    Set m_objDialogReport = dialogReport
End Property


Public Property Let Modal(ByVal setModal As Boolean)
    m_bSetModal = setModal
    
    If m_objDialogReport.Visible Then
        m_objDialogReport.Modal = setModal
    End If
End Property


Public Property Let Visible(ByVal isVisible As Boolean)
    m_objDialogReport.Visible = isVisible
    
    If isVisible = False Then m_objDialogReport.Modal = False        ' make sure the form isn't modal.  If this is hidden with Modal = True it can cause Access to appear as though it froze
End Property


'''''''''''''''''''''''
'METHODS'
'''''''''''''''''''''''
Public Sub launch(Optional ByVal openArgs As String)
    
    On Error Resume Next        ' Some of the forms passed into this class will not have the necessary parts below.  On Error Resume Next skips those steps
    With m_objDialogReport
        
        .txtOpenArgs = openArgs        ' gives the supplied arguments to the text box
        .OpenArguments = openArgs
        .ParseArgs        ' calls parseArgs on the form.  This is useful when the form must extract information from the opening arguments
        
        .OnUnload = EVENT_PROC        ' Set the OnUnload event to say "[Event Procedure]".  If it does not say this, we cannot actually capture the Unload Event
        .OnClose = EVENT_PROC        ' Set the OnClose event to say "[Event Procedure]".  If it does not say this, we cannot actually capture the Close Event
        
        .Modal = m_bSetModal        ' changes the Modal property to true if it was specified.  If not supplied, it will default to the form's modal property
        
        .Visible = True        ' make the form visible
        .SetFocus        ' make it the current form
        
    End With
    
    On Error GoTo 0
        
End Sub


' ----------------------------------------------------------------
' Procedure Name: m_objDialogReport_Close
' Purpose: Raises the FormClosed event and passes a ByRef cancel and a copy of this object
' Procedure Kind: Sub
' Procedure Access: Private
' Author: MayoraB
' Date: 5/17/2019
' ----------------------------------------------------------------
Private Sub m_objDialogReport_Close()
    RaiseEvent ReportAfterClose

    Set m_objDialogReport = Nothing

End Sub


'''''''''''''''''''''''''''''''''
'FORM EVENTS
'''''''''''''''''''''''''''''''''
Private Sub m_objDialogReport_Unload(Cancel As Integer)
    
    On Error GoTo 0
    Dim bEventCancel As Boolean: bEventCancel = False        ' initialize the event's cancel option
    RaiseEvent ReportClosed(Me, bEventCancel)        ' raise the event
    
    Cancel = bEventCancel        ' pass the event's cancel to the actual form's cancel
    
End Sub


Public Sub closeReport()
    On Error Resume Next        ' will fail if the form was closed by something else
    Me.Visible = False        ' hide the form
    Me.Modal = False
    
    On Error GoTo 0
End Sub

Public Sub Requery()
    m_objDialogReport.Requery
End Sub

