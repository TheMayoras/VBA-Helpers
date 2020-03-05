Attribute VB_Name = "PublishVersion"
'@Folder("Misc")

Option Compare Database
Option Explicit

#If IS_DEV Then

Const BACKEND As String = "B:\GLOBAL\2127-CHMUS\JACKSON\MOC\Automated MOC - User\Backend\Backend.accdb"
Const vbext_pk_Proc As Integer = 0

Public Sub CreatePublishedVersion()
    Dim objDbCopy As Object
Debug.Print "Starting to publish..."
    Set objDbCopy = CopyDatabase()
    SetSettings objDbCopy
    PublishDatabase objDbCopy
Debug.Print "Done..."

    Set objDbCopy = Nothing
End Sub

Private Function CopyDatabase() As Object
    Dim objFso    As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Dim sCopyTo   As String
    Dim objDb     As Object
    Set objDb = objFso.GetFile(CurrentDb.name)
    
    sCopyTo = Environ("TEMP") & "\" & objFso.GetFile(CurrentDb.name).name
    objFso.CopyFile CurrentDb.name, sCopyTo, True
    
    Set CopyDatabase = objFso.GetFile(sCopyTo)
    
Debug.Print "-- Copied database to " & sCopyTo
       
End Function

Private Sub SetSettings(dbCopy As Object)
    Dim objApp    As Application
    Set objApp = New Application
    
    With objApp
        .OpenCurrentDatabase dbCopy.path, True
        
        ' Compact on Close
        .SetOption "Auto Compact", True
        ' Disable dev environment
        .SetOption "Conditional Compilation Arguments", "IS_DEV=0"
        
        ' prevent design view and checking truncated fields
        .SetOption "DesignWithData", False
        .SetOption "CheckTruncatedNumFields", False
        
        ' set form to show when opening db
        SetProperty .CurrentDb, "StartupForm", "frm_Login", dbText
        
        ' hide toolbars and navigation pane
        SetProperty .CurrentDb, "StartupShowDBWindow", False, dbBoolean
        SetProperty .CurrentDb, "AllowFullMenus", False, dbBoolean
        SetProperty .CurrentDb, "AllowBuiltinToolbars", False, dbBoolean
        .DoCmd.SelectObject acModule, , True
        .DoCmd.RunCommand acCmdWindowHide
        
        PublishForms .CurrentProject
        PublishReports .CurrentProject
        AddMove00 .vbe.ActiveVBProject
        RelinkTables objApp
        
        .CloseCurrentDatabase
        .Quit
    End With
    Set objApp = Nothing
End Sub

Private Sub SetProperty(db As Database, key As String, value As Variant, propType As DataTypeEnum)
    On Error GoTo PROP_NOTEXIST

    db.Properties(key) = value
    
EXIT_SUB:
    On Error GoTo 0
Debug.Print "-- " & key & " set to " & value
    Exit Sub
    
PROP_NOTEXIST:
    If err.number = 3270 Then        ' Prop does not exist error
        db.Properties.Append db.CreateProperty(key, propType, value)
    End If
    Resume EXIT_SUB

End Sub


Private Sub PublishDatabase(dbCopy As Object)
    Dim objFso    As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Dim sAccdeName As String
    sAccdeName = objFso.GetFile(CurrentDb.name).ParentFolder.path & "\" & objFso.GetBaseName(dbCopy.name) & ".accde"
    
    If objFso.FileExists(sAccdeName) Then objFso.DeleteFile sAccdeName
    
    With New Access.Application        ' it seems to be necessary to use syscmd 603 with a separate access application
        .AutomationSecurity = 1        ' msoAutomationSecurityLow
        ' SysCmd 603 from -> https://codekabinett.com/rdumps.php?Lang=2&targetDoc=make-access-accde-vb-script
        .SysCmd 603, dbCopy.path, sAccdeName
    End With
    
Debug.Print "-- Database converted to " & sAccdeName

    Set objFso = Nothing
End Sub

Private Sub PublishForms(proj As CurrentProject)
    Dim i         As Integer
    For i = 0 To proj.AllForms.count - 1
        If proj.AllForms(i).IsLoaded Then
            proj.Application.DoCmd.Close acForm, proj.AllForms(i).name
        End If
        
        Dim f     As Form
        proj.Application.DoCmd.OpenForm proj.AllForms(i).name, acDesign
        Set f = proj.Application.Forms(proj.AllForms(i).name)
        With f
            .autoCenter = True
            .popUp = True
            .Move 0, 0
            .RecordSelectors = False
            .NavigationButtons = False
            .CloseButton = False
            .ControlBox = False
            .MinMaxButtons = False
            .ShortcutMenu = False
        End With
        
        proj.Application.DoCmd.Close acForm, f.name, acSaveYes
    Next i
    
Debug.Print "-- Published all forms"
End Sub


Private Sub PublishReports(proj As CurrentProject)
    Dim i         As Integer
    For i = 0 To proj.allReports.count - 1
        If proj.allReports(i).IsLoaded Then
            proj.Application.DoCmd.Close acReport, CurrentProject.allReports(i).name
        End If
        
        Dim r     As Report
        proj.Application.DoCmd.OpenReport proj.allReports(i).name, acDesign
        Set r = proj.Application.Reports(proj.allReports(i).name)
        With r
            .autoCenter = True
            .popUp = True
            .Move 0, 0
            .CloseButton = False
            .ControlBox = False
            .MinMaxButtons = False
        End With
        
        proj.Application.DoCmd.Close acReport, r.name, acSaveYes
    Next i
Debug.Print "-- Published all forms"
End Sub

Private Sub AddMove00(vbe As Object)
    Dim objModule As Object
    Dim sPrefix As String
    For Each objModule In vbe.VBComponents
        With objModule
            ' Load event needs to be prefixed with Report_ if report and likewise for forms
            If .name Like "Form_*" Then
                sPrefix = "Form"
            ElseIf .name Like "Report_*" Then
                sPrefix = "Report"
            Else
                GoTo NOT_FORM_REPORT
            End If
            
            Dim lProcLine As Long
            On Error Resume Next
            lProcLine = .CodeModule.ProcBodyLine(sPrefix & "_Load", vbext_pk_Proc)
            
            ' Proc doesn't exist, make it
            If err.number <> 0 Then
                lProcLine = .CodeModule.CreateEventProc("Load", sPrefix)
            End If
            On Error GoTo 0
            
            Dim lProcLineCount As Long
            lProcLineCount = .CodeModule.ProcCountLines(sPrefix & "_Load", vbext_pk_Proc)
            
            ' need to add Me.Move 0, 0
            If Not (.CodeModule.Lines(lProcLine, lProcLineCount) Like "*Move?0,?0*") Then
                .CodeModule.InsertLines lProcLine + 1, "Me.Move 0, 0:MSGBOX""ADDED IT"""
                Debug.Print "    -- Adding Me.Move 0, 0 to " & .name
            End If
            
            
        End With
NOT_FORM_REPORT:
    Next objModule
    Debug.Print "-- Added Move 0,0 to all Load"
End Sub

Private Sub RelinkTables(app As Application)
    Dim objTableDef As tableDef
    For Each objTableDef In app.CurrentDb.TableDefs
        With objTableDef
            If Not app.GetHiddenAttribute(acTable, .name) Then
                If .Connect Like "*DATABASE*" Then
                    .Connect = ";DATABASE=" & BACKEND
                    Debug.Print "    -- Linked " & .name
                'Else
                '    app.DoCmd.TransferDatabase acLink, "Microsoft Access", BACKEND, acTable, .name, .name, False
                End If
            End If
        End With
    Next objTableDef
End Sub

Public Function TotalLinesInProject() As Long
    Dim VBP As Object
    Dim VBComp As Object
    Dim LineCount As Long
    
    Set VBP = Application.vbe.ActiveVBProject
    
'    If VBP.Protection = vbext_pp_locked Then
'        TotalLinesInProject = -1
'        Exit Function
'    End If
    
    For Each VBComp In VBP.VBComponents
        LineCount = LineCount + VBComp.CodeModule.CountOfLines
    Next VBComp
    
    TotalLinesInProject = LineCount
End Function
    
#End If

