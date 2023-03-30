Attribute VB_Name = "FPMod"
Option Explicit

Public Const constFPout = 1     '[INTEGER FOR FILE ACCESS TYPE INPUT/OUTPUT]
Public Const constFPIn = 0

Public strGuardRow              '[STRING - GUARD ROW VARIETY]
Public strOrientation           '[STRING - ORIENTATION OF THE FIELD PLAN]
Public strDesignFileOrder       '[STRING - DESIGN FILE ORDER]
Public RegUser As String        '[STRING - REGISTERED USER]
Public flagSaved As Boolean     '[BOOLEAN - FLAG FOR FILE SAVE STATUS]
Public flagChanged As Boolean   '[BOOLEAN - FLAG FOR FILE CHANGED STATUS]
Public flagLoaded As Boolean    '[BOOLEAN - FILE HAS BEEN LOADED]


Public Sub GEN_ErrorHandler(ErrorNumber As Integer, Routine As String, Descriptor As String, Description As String, source As String)

    '[GENERIC ERROR HANDLING ROUTINE FOR VISUAL BASIC 4.0]
    Dim Result
    Dim Message, title, Options, Default
    
    Err.Clear
    
    Select Case ErrorNumber
    Case 52, 62, 76         '[ERROR OPENING FILE FROM COMMAND LINE]
        Message = "An error has occurred while accessing the file " & Descriptor & " in process " & Routine & Chr$(10) & Chr$(10) & "Error Description : " & Description & " (" & ErrorNumber & ")" & Chr$(10) & Chr$(10) & "This action will be cancelled and a clean field plan workspace will be created.  Click on OK to continue."
        title = "Error Opening File"
        Options = vbCritical + vbOKOnly
        Result = MsgBox(Message, Options, title)
        frmMain.GridFieldPlan.Rows = 1
        frmMain.GridFieldPlan.Cols = 1
        Exit Sub
    Case 32755              '[CANCEL WAS SELECTED]
        Exit Sub
    Case Else               '[UNEXPECTED ERROR]
        Message = "An unexpected error has occurred. Source descriptor =  " & Descriptor & " in process " & Routine & Chr$(10) & Chr$(10) & "Error Description : " & Description & " (" & ErrorNumber & ")" & Chr$(10) & Chr$(10) & "This action will be cancelled.  Please contact the Author with the details of this error message.  Click on OK to exit."
        title = "Unexpected Error"
        Options = vbCritical + vbOKOnly
        Result = MsgBox(Message, Options, title)
        Close
        Unload frmMain
        
    End Select
    
End Sub

Public Sub ResizeForm()

    Dim intColCounter

    '[STRETCH LINE TO FORM WIDTH]
    frmMain.LineMenu.X1 = 0
    frmMain.LineMenu.X2 = frmMain.Width

    '[RESIZE GRID TO MATCH FORM SIZE]
    If frmMain.WindowState <> 1 Then
        frmHelp.Visible = True
        frmMain.GridFieldPlan.Width = frmMain.Width - 400
        frmMain.GridFieldPlan.Height = (frmMain.Height - 600) - frmMain.GridFieldPlan.Top
        frmMain.GridFieldPlan.Left = (frmMain.Width - frmMain.GridFieldPlan.Width) / 2 * 0.75
        frmMain.lblDescription.Left = frmMain.GridFieldPlan.Left
        frmMain.lblDescription.Width = frmMain.GridFieldPlan.Width
        frmMain.lblDescription.Top = frmMain.GridFieldPlan.Top - frmMain.lblDescription.Height - 5
        '[RESIZE GRID COLUMNS TO MATCH FORM SIZE]
        For intColCounter = 0 To (frmMain.GridFieldPlan.Cols - 1)
            frmMain.GridFieldPlan.ColWidth(intColCounter) = (frmMain.GridFieldPlan.Width / frmMain.GridFieldPlan.Cols) * 0.95
        Next intColCounter
'        frmHelp.Top = Screen.Height - frmHelp.Height - (Screen.Height * 0.05)
'        frmHelp.Left = Screen.Width - frmHelp.Width - (Screen.Width * 0.05)
    ElseIf frmMain.WindowState = 1 Then
        frmHelp.Visible = False
    End If


End Sub

Sub FileRead(ReadFileName As String)

    '[CALL ERROR HANDLING ROUTINE]
    On Error GoTo ErrorHandler

    Dim intRowCounter As Integer
    Dim intColCounter As Integer
    Dim FileHandle As Integer
    Dim txtDummy
    Dim varDummy, intRows, intCols
    '[ALLOCATE FILEHANDLE AND READ SELECTED FILE DETAILS]
    FileHandle = OpenFile(ReadFileName, constFPIn)
    Input #FileHandle, txtDummy: frmDesc.TextSiteDesc = txtDummy
    Input #FileHandle, txtDummy: frmDesc.TextDate = txtDummy
    Input #FileHandle, txtDummy: frmDesc.TextCoOperator = txtDummy
    Input #FileHandle, txtDummy: frmDesc.TextPlantingRate = txtDummy
    Input #FileHandle, txtDummy: frmDesc.TextFertiliser = txtDummy
    Input #FileHandle, txtDummy: frmDesc.TextHerbicide = txtDummy
    '[GUARD ROW]
    Input #FileHandle, strGuardRow
    '[DESIGN FILE ORDER]
    Input #FileHandle, strDesignFileOrder
    '[GRID SIZE]
    Input #FileHandle, intRows, intCols

    ShowProgress "Opening file " & ReadFileName, "Reading field plan details", 0

    If intRows = 0 Then intRows = 1
    If intCols = 0 Then intCols = 1

    frmMain.GridFieldPlan.Rows = intRows
    frmMain.GridFieldPlan.Cols = intCols
    varDummy = 0
    '[WRITE GRID]
    For intRowCounter = 0 To (frmMain.GridFieldPlan.Rows - 1)
        frmMain.GridFieldPlan.Row = intRowCounter   '[SET ROW POSITION]
        For intColCounter = 0 To (frmMain.GridFieldPlan.Cols - 1)
            varDummy = varDummy + 1
            frmMain.GridFieldPlan.Col = intColCounter   '[SET COL POSITION]
            Input #FileHandle, txtDummy: frmMain.GridFieldPlan.Text = txtDummy
            UpdateProgress "", "Reading cell " & Str(varDummy) & " of " & Str(intRows * intCols), (varDummy / (intRows * intCols)) * 100
        Next intColCounter
    Next intRowCounter
    Close #FileHandle

    HideProgress

ErrorHandler:
    If Err.Number > 0 Then
        Call GEN_ErrorHandler(Err.Number, "FileRead", ReadFileName, Err.Description, Err.source)
        HideProgress
    End If
    
End Sub

Public Sub FileSave(SaveFileName As String)
    
    '[CALL ERROR HANDLING ROUTINE]
    On Error GoTo ErrorHandler
    Dim varDummy
    ShowProgress "Saving file " & SaveFileName & ". Please wait ...", "Writing field plan details", 0
    
    '[SAVE DATA AS PASSED SaveFileName]
    Dim FileHandle As Integer
    Dim intRowCounter As Integer
    Dim intColCounter As Integer
    FileHandle = OpenFile(SaveFileName, constFPout)

    '[WRITE DETAILS TO OUTPUT FILE]
    Write #FileHandle, frmDesc.TextSiteDesc
    Write #FileHandle, frmDesc.TextDate
    Write #FileHandle, frmDesc.TextCoOperator
    Write #FileHandle, frmDesc.TextPlantingRate
    Write #FileHandle, frmDesc.TextFertiliser
    Write #FileHandle, frmDesc.TextHerbicide
    '[GUARD ROW]
    Write #FileHandle, strGuardRow
    '[DESIGN FILE ORDER]
    Write #FileHandle, strDesignFileOrder
    '[GRID SIZE]
    Write #FileHandle, frmMain.GridFieldPlan.Rows, frmMain.GridFieldPlan.Cols
    '[WRITE GRID]
    varDummy = 0
    For intRowCounter = 0 To (frmMain.GridFieldPlan.Rows - 1)
        frmMain.GridFieldPlan.Row = intRowCounter   '[SET ROW POSITION]
        For intColCounter = 0 To (frmMain.GridFieldPlan.Cols - 1)
            varDummy = varDummy + 1
            frmMain.GridFieldPlan.Col = intColCounter   '[SET COL POSITION]
            UpdateProgress "", "Writing cell " & Str(varDummy) & " of " & Str(frmMain.GridFieldPlan.Rows * frmMain.GridFieldPlan.Cols), (varDummy / (frmMain.GridFieldPlan.Rows * frmMain.GridFieldPlan.Cols)) * 100
            If intColCounter < (frmMain.GridFieldPlan.Cols - 1) Then
                Write #FileHandle, frmMain.GridFieldPlan.Text;
            Else
                Write #FileHandle, frmMain.GridFieldPlan.Text
            End If
        Next intColCounter
    Next intRowCounter
    Close #FileHandle
    HideProgress
        
ErrorHandler:
    If Err.Number > 0 Then
        Call GEN_ErrorHandler(Err.Number, "FileSave", SaveFileName, Err.Description, Err.source)
        HideProgress
    End If

End Sub

Public Function SaveCheck()

    '[FUNCTION TO CHECK IF FILE HAS BEEN CHANGED/SAVED]
    Dim Result

    If flagChanged = True Then  '[FILE HAS BEEN CHANGED AND NEEDS TO BE SAVED]
        Result = Confirm("You have made changes to this file.  Do you wish to save these changes ?", "Warning - changes not saved.", vbYesNoCancel)
        If Result = vbYes Then
            '[USER WISHES TO SAVE CHANGES TO THIS FILE]
            If flagLoaded Then '[FILE HAS A NAME, JUST SAVE]
                FileSave (frmMain.CommonDialog.FileName)
            Else    '[FILE HAS NO NAME, USE SAVE AS]
                FileSaveAs ("fldplan.fpl")
            End If
            flagChanged = False
            flagSaved = True
        ElseIf Result = vbNo Then
            '[USER DOES NOT WISH TO SAVE THE CHANGES TO THIS FILE]
            flagChanged = False
            flagSaved = False
        ElseIf Result = vbCancel Then
            '[USER CANCELLED OPERATION]
            flagSaved = False
        End If
    End If

    SaveCheck = Result


End Function
Public Function FileSaveAs(NewFileName As String) As Boolean
    
    '[CALL ERROR HANDLING ROUTINE]
    On Error GoTo ErrorHandler

FileSaveAsDialog:
    '[FUNCTION TO SAVE FILE AS A NEW NAME AND RETURN CODE INDICATING WHETHER CANCEL WAS PRESSED]
    Dim Result

    '[SAVE CURRENT DATA AS A NEW FILE]
    frmMain.CommonDialog.DialogTitle = "Save as a new field plan file"
    frmMain.CommonDialog.FileName = NewFileName
    frmMain.CommonDialog.ShowSave

    '[PROCESS COMMON DIALOG SAVE FORM]
    FileSave (frmMain.CommonDialog.FileName)

    FileSaveAs = True

ErrorHandler:
    If Err.Number > 0 Then
        Call GEN_ErrorHandler(Err.Number, "FileSaveAs", NewFileName, Err.Description, Err.source)
        FileSaveAs = False
    End If

End Function

Public Function Confirm(Msg, title, ButtonStyle)

    Rem frmHelp.Hide

    Dim Style
    '[FUNCTION TO CONFIRM SOME ACTION]
    If ButtonStyle = vbOKOnly Then
        Style = ButtonStyle + vbCritical + vbDefaultButton1 + vbApplicationModal  ' Define buttons.
    Else
        Style = ButtonStyle + vbQuestion + vbDefaultButton2 + vbApplicationModal  ' Define buttons.
    End If
    Confirm = MsgBox(Msg, Style, title)

    Rem frmHelp.Show


End Function

Public Function OpenFile(FileName As String, FileMode As Integer) As Integer

    '[FUNCTION TO OPEN A FILE AND RETURN THE FILE HANDLE ASSOCIATED WITH THE FILE]
    '[FILEMODE 0=INPUT]
    '[FILEMODE 1=OUTPUT]
    Dim FileHandle As Integer '[Next Free File Handle]
    Dim Result
    FileHandle = FreeFile(0)  '[Allocate free handle 1-255]

    Select Case FileMode
        Case 0
            Open FileName For Input As #FileHandle    '[OPEN FILE FOR INPUT]
        Case 1
            Open FileName For Output As #FileHandle '[OPEN FILE FOR OUTPUT]
        Case Else
    End Select

    OpenFile = FileHandle         '[RETURN FILE HANDLE]

End Function
Public Sub SetFormTitles()

    '[ALIGN FORM TITLES WITH THAT OF THE DESCRIPTION FORM]
    frmMain.lblDescription = frmDesc.TextSiteDesc.Text & " - " & frmDesc.TextDate.Text & " - " & frmDesc.TextCoOperator.Text


End Sub

Public Sub FormHelp(Helptext As String, CallingForm)

    '[PASS HELP TEXT TO HELP FORM]
    If frmHelp.lblText = Helptext Then Exit Sub

    frmHelp.lblText = Helptext
    '    frmHelp.ZOrder 0
    '    If CallingForm.GotFocus Then CallingForm.SetFocus

End Sub

Public Sub ShowProgress(CaptionTXT, CommandTXT, ProgressValue)
    
    ' [LOAD AND DISPLAY THE PROGRESS FORM DIALOG]
    frmProgr.Show 0
    frmProgr.Refresh
    frmProgr.gaugeProgress.Value = ProgressValue
    frmProgr.lblCaption = CaptionTXT
    frmProgr.lblCommand = CommandTXT
    frmProgr.Refresh
    
End Sub

Public Sub HideProgress()

    '[UNLOAD FILER FORMS AND RETURN TO MAIN FORM]
    Unload frmProgr

End Sub

Public Sub UpdateProgress(CaptionTXT, CommandTXT, ProgressValue)

    '[UPDATE PROGRESS DISPLAY WITH LATEST VALUES]
    If CaptionTXT > "" Then
        frmProgr.lblCaption = CaptionTXT
        frmProgr.lblCaption.Refresh
    End If
    If CommandTXT > "" Then
        frmProgr.lblCommand = CommandTXT
        frmProgr.lblCommand.Refresh
    End If
    If ProgressValue > 0 Then
        frmProgr.gaugeProgress.Value = ProgressValue
        frmProgr.gaugeProgress.Refresh
    End If
    
End Sub
