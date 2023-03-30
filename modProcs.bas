Attribute VB_Name = "modProcs"
Option Explicit
'[----------------------------------------------]
'[modProcs.bas      Basic Sub Modules           ]
'[----------------------------------------------]
'[GSR                                           ]
'[generic staff roster system for Windows       ]
'[(C) softEdge & David Gilbert, 1997            ]
'[----------------------------------------------]
'[version                           2.1         ]
'[initial date                      21/07/1997  ]
'[----------------------------------------------]

Public DsStaff         As Dynaset
Public DsClass         As Dynaset
Public DsRoster        As Dynaset
Public DsDefault       As Dynaset
Public DsReport        As Dynaset
Public DBMain          As Database

'[FLAG FOR BETA OR FULL VERSION]
Public flagBeta        As Boolean
Public Const constDelay = 1        'delay for timer
Public Const constEmail = "gilberd@dpi.qld.gov.au"
Public Const constFileOut = 1     '[INTEGER FOR FILE ACCESS TYPE INPUT/OUTPUT]
Public Const constFileIn = 0

'[WARNING LEVELS FOR EXCEPTION REPORT]
Public Const constCritical = 2
Public Const constSerious = 1
Public Const constWarning = 0

Type WeekType
    ShortDay            As String
    LongDay             As String
End Type
    
'[STAFF TYPE FOR STAFF DETAILS]
Type StaffType
    DayDate         As Date
    Roster          As String
    Minutes         As Single
    Amount          As Single
    StartTime       As Date
    EndTime         As Date
End Type
    
Public ArrayWeek(7)         As WeekType
Public flagDeleteConfirm    As Boolean      'flag for deletion confirmation
Public intRosterClass       As Integer      'roster class id number (1 - 10)
Public gsrReturn            As Double       'return code from gsr msg box
Public gsrNote              As String       'string for roster note
Public strClip              As String       'string for holding clip
Public strFileName          As String       'string for filename
Public intDaysUsed          As Integer      'number of days program has been installed for
Public sinModifier          As Single       'modifier for validation code
Function Validate(strValidate) As String

    '[FUNCTION TO VALIDATE A STRING AND RETURN THE VALIDATION CODE]
    Dim strRegCode      As String       '[RETURNED HEX VALIDATION CODE]
    Dim sinValue        As Single       '[ACCUMULATED VALUE]
    Dim intCounter      As Integer      '[COUNTER FOR LENGTH]

    If IsNull(strValidate) Or Len(strValidate) = 0 Then
        '[NO CODE TO VALIDATE SO RETURN BLANK]
        strRegCode = ""
    Else
        For intCounter = 1 To Len(strValidate)
        '[CYCLE THROUGH VALIDATION STRING AND ACCUMULATE VALUES]
            sinValue = sinValue + (Asc(Mid$(strValidate, intCounter, 1)) * sinModifier)
        Next intCounter
        strRegCode = Hex(sinValue)
    End If
    
    '[RETURN CODE VALUE]
    Validate = strRegCode

End Function

Sub exCheckNameInStaffList(strCellText, intDayCount)

    '[DECLARE VARIABLES REQUIRED]
    Dim strBookmark         As String
    Dim strFullname         As String
    Dim strLastName         As String
    Dim strFirstName        As String
    Dim intStart            As Integer
    Dim intFinish           As Integer
    Dim intBreak            As Integer
    Dim SQLStmt             As String
    Dim flagNameFound       As Boolean
    Dim intClass            As Integer
    Dim strTime             As String
    
    '[SAVE BOOKMARK]
    strBookmark = DsStaff.Bookmark
    
    '[SET FLAG]
    flagNameFound = True
    If Len(Trim(strCellText)) = 0 Or IsNull(strCellText) Then Exit Sub
    
    Do While flagNameFound
        '[EXTRACT NAME FROM CELL TEXT AND CHECK FOR IT IN THE STAFF LIST]
        intBreak = InStr(strCellText, Chr$(vbKeyReturn))
        Select Case intBreak
        Case 0          '[NO BREAK FOUND]
            If Len(Trim(strCellText)) = 0 Then
                '[RESTORE BOOKMARK]
                DsStaff.Bookmark = strBookmark
                Exit Sub
            End If
            strFullname = Trim$(strCellText)
            strCellText = ""
        Case Else       '[BREAK FOUND]
            strFullname = Trim$(Left$(strCellText, intBreak - 1))
            strCellText = Trim$(Mid$(strCellText, intBreak + 1))
        End Select
        
        strLastName = Trim(Left$(strFullname, InStr(strFullname, ",") - 1))
        strFirstName = Trim(Mid$(strFullname, InStr(strFullname, ",") + 1))
        
        '[LOCATE LASTNAME, FIRSTNAME IN DYNASET]
        SQLStmt = "LastName = '" & strLastName & "' AND FirstName = '" & strFirstName & "'"
        DsStaff.FindFirst SQLStmt
        
        If DsStaff.NoMatch Then
            '[NAME NOT FOUND IN STAFF LIST, ADD TO EXCEPTION REPORT]
            intClass = DsReport("Class")
            strTime = DsReport("ShiftStart")
            Call exAddNewException(5, intClass, intDayCount, strTime, strFullname, "Name not found in staff list")
        End If
    
    Loop
    
    '[RESTORE BOOKMARK]
    DsStaff.Bookmark = strBookmark

End Sub

Sub FileRead(ReadFileName As String)

    '[CALL ERROR HANDLING ROUTINE]
    On Error GoTo ErrorHandler
    
    Dim intRows             As Integer
    Dim intCols             As Integer
    Dim intRowCounter       As Integer
    Dim intColCounter       As Integer
    Dim FileHandle          As Integer
    Dim strClassDesc        As String
    Dim strClassID          As String
    Dim strStartDate        As String
    Dim txtDummy            As String
    Dim varDummy
    
    '[SHOW PROGRESS REPORT]
    Call ReportInfo("Reading from " & ReadFileName, 0)
    
    '[ALLOCATE FILEHANDLE AND READ SELECTED FILE DETAILS]
    FileHandle = OpenFile(ReadFileName, constFileIn)
        
    '[READ DETAILS FROM INPUT FILE]
    Input #FileHandle, strClassDesc                     '[CLASS DESCRIPTION]
    Input #FileHandle, strClassID                       '[CLASS CODE]
    Input #FileHandle, strStartDate                     '[STARTING DATE OF ROSTER]
    
    '[===================================================================================]
    '[NOW SEE IF WE CAN CHANGE TO THIS ROSTER, OTHERWISE JUST PLACE IN THE CURRENT ROSTER]
    If LocateClass(strClassDesc) = True Then
        '[DESCRIPTION FOUND]
        frmRoster.ComboClass = strClassDesc '[SET DESCRIPTION]
        frmControl.MaskDate = strStartDate '[SET DATE]
    Else
        '[DESCRIPTION NOT FOUND - REPLACE THIS ROSTER]
        
    End If
    '[===================================================================================]
    
    '[GRID SIZE]
    Input #FileHandle, intRows, intCols

    If intRows < 2 Then intRows = 2
    If intCols < 9 Then intCols = 9

    frmRoster.GridRoster.Rows = intRows
    frmRoster.GridRoster.Cols = intCols
    varDummy = 0
    
    '[WRITE GRID]
    For intRowCounter = 1 To (frmRoster.GridRoster.Rows - 1)
        
        '[PROGRESS BAR]
        Call ReportProgressBar((frmRoster.GridRoster.Row / (frmRoster.GridRoster.Rows - 1)) * 100)

        frmRoster.GridRoster.Row = intRowCounter   '[SET ROW POSITION]
        For intColCounter = 0 To (frmRoster.GridRoster.Cols - 1)
            varDummy = varDummy + 1
            frmRoster.GridRoster.Col = intColCounter   '[SET COL POSITION]
            Input #FileHandle, txtDummy: frmRoster.GridRoster.Text = txtDummy
            '[RESIZE THE ROSTER CELL]
            ResizeRosterCell (txtDummy)
        Next intColCounter
    Next intRowCounter
    
    '[CLOSE FILE]
    Close #FileHandle
    
    '[CLEAR REPORT INFO]
    Call ReportInfo("", 0)
    '[CLEAR PROGRESS BAR]
    Call ReportProgressBar(0)
        

ErrorHandler:
    If Err.Number > 0 Then
        '[CANCEL WAS PRESSED ON THE SAVE FORM - NO PROCESSING REQUIRED]
        If Err.Number = cdlCancel Then Debug.Print "Cancel Pressed in File Save Module"
        '[DEBUG]
        Debug.Print Err.Number & " - Error in File Read Module"
    End If
    
End Sub

Public Sub FileSave(SaveFileName As String)
    
    '[CALL ERROR HANDLING ROUTINE]
    On Error GoTo ErrorHandler
    Dim varDummy
    
    '[SHOW PROGRESS REPORT]
    Call ReportInfo("Saving to " & SaveFileName, 0)
    
    '[SAVE DATA AS PASSED SaveFileName]
    Dim FileHandle As Integer
    Dim intRowCounter As Integer
    Dim intColCounter As Integer
    FileHandle = OpenFile(SaveFileName, constFileOut)

    '[WRITE DETAILS TO OUTPUT FILE]
    Write #FileHandle, DsClass("Description")           '[CLASS DESCRIPTION]
    Write #FileHandle, DsClass("Code")                  '[CLASS CODE]
    Write #FileHandle, Str$(DsClass("StartDate"))       '[STARTING DATE OF ROSTER]
    '[GRID SIZE]
    Write #FileHandle, frmRoster.GridRoster.Rows, frmRoster.GridRoster.Cols
    '[WRITE GRID]
    varDummy = 0
    For intRowCounter = 1 To (frmRoster.GridRoster.Rows - 1)
    
        '[PROGRESS BAR]
        Call ReportProgressBar((frmRoster.GridRoster.Row / (frmRoster.GridRoster.Rows - 1)) * 100)
    
        frmRoster.GridRoster.Row = intRowCounter   '[SET ROW POSITION]
        For intColCounter = 0 To (frmRoster.GridRoster.Cols - 1)
            varDummy = varDummy + 1
            frmRoster.GridRoster.Col = intColCounter   '[SET COL POSITION]
            If intColCounter < (frmRoster.GridRoster.Cols - 1) Then
                Write #FileHandle, frmRoster.GridRoster.Text;
            Else
                Write #FileHandle, frmRoster.GridRoster.Text
            End If
        Next intColCounter
    Next intRowCounter
    
    '[CLOSE FILE]
    Close #FileHandle
    
    '[CLEAR REPORT INFO]
    Call ReportInfo("", 0)
    '[CLEAR PROGRESS BAR]
    Call ReportProgressBar(0)
        
ErrorHandler:
    If Err.Number > 0 Then
        '[DEBUG]
        Debug.Print Err.Number & " - Error in File Save Module"
    End If

End Sub

Sub FileSetFilter()

    Dim intClassCounter         As Integer
    Dim strBookmark             As String

    '[default file extension]
    mdiMain.CommonDialog.DefaultExt = Trim(DsClass("Code"))
    
    '[save class bookmark]
    strBookmark = DsClass.Bookmark
    
    '[Set Filter Property]
    mdiMain.CommonDialog.Filter = ""
    For intClassCounter = 0 To 9
        DsClass.AbsolutePosition = intClassCounter
        '[ONLY ALLOW SAVE/LOAD TO ACTIVE ROSTERS]
        If DsClass("Active") = vbChecked Then
            '[ADD ACTIVE CLASSES TO FILTER LIST}
            mdiMain.CommonDialog.Filter = mdiMain.CommonDialog.Filter & Trim(DsClass("Description")) & " (*." & Trim(DsClass("Code")) & ")"
            mdiMain.CommonDialog.Filter = mdiMain.CommonDialog.Filter & "|*." & Trim(DsClass("Code")) & "|"
            '[SET FILTER INDEX IF A MATCH IS FOUND]
            If mdiMain.CommonDialog.DefaultExt = Trim(DsClass("Code")) Then mdiMain.CommonDialog.FilterIndex = (intClassCounter + 1)
        End If
    Next intClassCounter
    
    '[add standard file types to end of filter list]
    mdiMain.CommonDialog.Filter = mdiMain.CommonDialog.Filter & "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    '[restore class bookmark]
    DsClass.Bookmark = strBookmark

End Sub

Function LineCount(strText) As Integer

    '[FUNCTION TO COUNT NUMBER OF RETURN CHARACTERS WITHIN A GIVEN STRING]
    Dim intLineCount            As Integer
    Dim intPos                  As Integer
    
    intPos = 1
    '[STOP NULL ERROR]
    strText = Trim(strText & "")
    If Len(strText) = 0 Then intLineCount = 0

    Do While intPos > 0
        intPos = InStr(intPos + 1, strText, Chr$(vbKeyReturn))
        intLineCount = intLineCount + 1
    Loop
    
    '[RETURN VALUE]
    LineCount = intLineCount
    
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
            Open FileName For Input As #FileHandle      '[OPEN FILE FOR INPUT]
        Case 1
            Open FileName For Output As #FileHandle     '[OPEN FILE FOR OUTPUT]
        Case Else
    End Select

    OpenFile = FileHandle         '[RETURN FILE HANDLE]

End Function

Public Function FileSaveAs(NewFileName As String) As Boolean
    
    '[CALL ERROR HANDLING ROUTINE]
    On Error GoTo ErrorHandler

FileSaveAsDialog:
    '[FUNCTION TO SAVE FILE AS A NEW NAME AND RETURN CODE INDICATING WHETHER CANCEL WAS PRESSED]
    Dim Result

    '[SAVE CURRENT DATA AS A NEW FILE]
    mdiMain.CommonDialog.DialogTitle = "Save Roster As"
    mdiMain.CommonDialog.FileName = NewFileName
    mdiMain.CommonDialog.CancelError = True
    mdiMain.CommonDialog.ShowSave
    
    '[SET FILE DIALOG FLAGS]
    mdiMain.CommonDialog.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
    '[PROCESS COMMON DIALOG SAVE FORM]
    FileSave (mdiMain.CommonDialog.FileName)

    FileSaveAs = True

ErrorHandler:
    If Err.Number > 0 Then
        '[CANCEL WAS PRESSED ON THE SAVE FORM - NO PROCESSING REQUIRED]
        If Err.Number = cdlCancel Then Debug.Print "Cancel Pressed in File Save Module"
        '[DEBUG]
        Debug.Print Err.Number & " - Error in File Save As Module"
    End If

End Function


Sub procAllStaffRosters()
    
    '[THIS ROUTINE WILL CALL THE APPROPRIATE SUBROUTINE FOR ALL STAFF MEMBERS]
    '[AND DISPLAY ANY WARNING MESSAGES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    Dim strFullname         As String
    Dim strBookmark         As String
    Dim intCounter          As Integer

    
    '[CHECK FOR STAFF RECORD]
    If DsStaff.BOF And DsStaff.EOF Then Exit Sub
    '[SAVE STAFF DYNASET POSITION]
    strBookmark = DsStaff.Bookmark
    
    '[MOVE TO FIRST RECORD]
    DsStaff.MoveFirst
    
    '[SHOW WARNING FORM]
    Msg = "The full staff roster report prints weekly roster details for all staff members (" & DsStaff.RecordCount & " records).  Because the output is sent directly to your printer, please ensure it is connected and switched on." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to continue and print the reports ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
        
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now processing your roster files and producing the required staff roster for " & strFullname & "." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "This may take a few minutes."
        Style = vbInformation            ' Define buttons.
        Title = "Staff Roster"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.GaugeProgress.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[SET UP LOOP]
        Do While Not DsStaff.EOF
            '[INCREMENT COUNTER]
            intCounter = intCounter + 1
            '[APPLY STAFF NAME TO FULL STRING]
            strFullname = DsStaff("LastName") & ", " & DsStaff("FirstName")
            '[SHOW PROGRESS REPORT]
            Call ReportInfo(strFullname, 0)
            '[PROGRESS BAR]
            Call ReportProgressBar((intCounter / DsStaff.RecordCount) * 100)
            
            '[CALL STAFF ROSTER SUBROUTINE FOR THE HIGHLIGHTED STAFF MEMBER]
            prnStaffRoster (strFullname)
            
            '[MOVE TO NEXT RECORD]
            DsStaff.MoveNext
        Loop
        
        '[END PRINTING OF ROSTER FORM]
        Printer.EndDoc
    
        '[HIDE GSR MESSAGE FORM]
        Unload frmMsg
    
    End If

    '[RESTORE STAFF DYNASET POSITION]
    DsStaff.Bookmark = strBookmark

End Sub

Sub procExceptionReport()
    
    '[COMMAND TO CHECK ROSTER FOR -ANY- IRREGULARITIES AND REPORT THEM TO THE REPORT FORM]
    '[USE THE REPORT FORM TO ADDRESS ANY OF THESE PROBLEMS]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response

    '[SHOW WARNING FORM]
    Msg = "The exception report details any problems which may occur within your currently defined and active rosters." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Because this routine has to perform multiple comparisions and searches, it may take a few minutes to complete, depending upon the number of rosters, staff and the speed of your computer." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to continue and produce the report ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
        '[SHOW PLEASE WAIT MESSAGE]
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "... Please Wait ..."
        '[---------------------------------------------------------------------------------]
        mdiMain.panelStatusBar.Refresh
        
        '[QUICKEST WAY TO REINITIALISE GRID AND FORM IS TO UNLOAD IT THEN RELOAD IT]
        If frmReport.Visible Then Unload frmReport
        frmReport.Show
        frmReport.GridReport.Cols = 6
        
        '[SET REPORT FORM FONT]
        Set frmReport.GridReport.Font = frmRoster.GridRoster.Font
        '[SET GRID REPORT TITLES AND SIZE]
        frmReport.GridReport.Row = 0
        frmReport.GridReport.Col = 0: frmReport.GridReport.Text = ""
        frmReport.GridReport.Col = 1: frmReport.GridReport.Text = "Roster"
        frmReport.GridReport.Col = 2: frmReport.GridReport.Text = "Day"
        frmReport.GridReport.Col = 3: frmReport.GridReport.Text = "Time"
        frmReport.GridReport.Col = 4: frmReport.GridReport.Text = "Staff"
        frmReport.GridReport.Col = 5: frmReport.GridReport.Text = "Exception"
        
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now processing your roster files and producing an exception report." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "The report will detail which problems (if any) have been found with your rosters.  Double-clicking on the error listed in the report will take you to the source of the problem." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "This may take a few minutes."
        Style = vbInformation            ' Define buttons.
        Title = "Exception Report"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.GaugeProgress.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[CALL EXCEPTION REPORT SUBROUTINE]
        ExceptionReport
        
        '[HIDE GSR MESSAGE FORM]
        Unload frmMsg
        '[SET REPORT TITLE]
        frmReport.LabelTitle = "GSR Exception Report"
        '[RESIZE FORM TO MAXIMUM SIZE]
        frmReport.WindowState = 2

        '[CLOSE FORM IF NO ERRORS FOUND AND POPUP MESSAGE]
        If frmReport.GridReport.Rows <= 2 Then
            frmMsg.GaugeProgress.Visible = False
            frmMsg.labelInfo.Visible = False
            Unload frmReport
            Msg = "GSR has found no immediate problems with your roster design." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Running the Exception Report command regularly will enable you to design your rosters with a minimum of effort." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Remember that the exception report is only produced for rosters which are 'enabled' on the Control form."
            Style = vbOKOnly ' Define buttons.
            Title = "All Clear !"
            Response = gsrMsg(Msg, Style, Title)
        End If
    End If

End Sub

Sub procCurrentGrid()
    
    '[USE THE REPORT FORM TO ADDRESS ANY OF THESE PROBLEMS]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    
    '[ERROR CHECK FOR IF NO FORMS ARE LOADED]
    On Error Resume Next
    
    Select Case mdiMain.ActiveForm.Name
    Case "frmRoster"
        '[SHOW WARNING FORM]
        Msg = "Continue and print the " & frmRoster.ComboClass.Text & " roster ?"
        Style = vbYesNo                     ' Define buttons.
        Title = "Confirmation Required"
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then
            '[SHOW GSR MESSAGE FORM]
            Msg = "GSR is now printing your roster." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Please wait."
            Style = vbInformation            ' Define buttons.
            Title = "Printing Roster"
            Response = gsrMsg(Msg, Style, Title)
            frmMsg.GaugeProgress.Visible = True
            frmMsg.labelInfo.Visible = True
            frmMsg.ZOrder
            frmMsg.Refresh
            
            '[CALL GRID PRINT PROCEEDURE]
            Call prnGrid(frmRoster.ComboClass.Text & " Roster", frmRoster.GridRoster, 0, 0, "Week Starting : " & frmControl.MaskDate.Text, True)
            
            '[HIDE GSR MESSAGE FORM]
            Unload frmMsg
            
        End If
        
    Case "frmReport"
        '[SHOW WARNING FORM]
        Msg = "Continue and print this report form ?"
        Style = vbYesNo                     ' Define buttons.
        Title = "Confirmation Required"
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then
            '[SHOW GSR MESSAGE FORM]
            Msg = "GSR is now printing the report form." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Please wait."
            Style = vbInformation            ' Define buttons.
            Title = "Printing Report Form"
            Response = gsrMsg(Msg, Style, Title)
            frmMsg.GaugeProgress.Visible = True
            frmMsg.labelInfo.Visible = True
            frmMsg.ZOrder
            frmMsg.Refresh
            
            '[CALL GRID PRINT PROCEEDURE]
            Call prnGrid(frmReport.LabelTitle, frmReport.GridReport, 1, 1, "", False)
            
            '[HIDE GSR MESSAGE FORM]
            Unload frmMsg
            
        End If
        
        
    Case Else
    
    End Select

End Sub

Sub procSelectedStaffRoster()

    '[THIS ROUTINE WILL CALL THE APPROPRIATE SUBROUTINE FOR A SINGLE STAFF MEMBER]
    '[AND DISPLAY ANY WARNING MESSAGES]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    Dim strFullname         As String
    
    '[CHECK FOR STAFF RECORD]
    If DsStaff.BOF And DsStaff.EOF Then Exit Sub
    '[APPLY STAFF NAME TO FULL STRING]
    strFullname = DsStaff("LastName") & ", " & DsStaff("FirstName")
    
    '[SHOW WARNING FORM]
    Msg = "The staff roster report prints weekly roster details for the selected staff member (" & strFullname & ").  Because the output is sent directly to your printer, please ensure it is connected and switched on." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to continue and print the report ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
        
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now processing your roster files and producing the required staff roster for " & strFullname & "." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "This may take a few minutes."
        Style = vbInformation            ' Define buttons.
        Title = "Staff Roster"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.GaugeProgress.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[CALL STAFF ROSTER SUBROUTINE FOR THE HIGHLIGHTED STAFF MEMBER]
        prnStaffRoster (strFullname)
        
        '[END PRINTING OF ROSTER FORM]
        Printer.EndDoc
        
        '[HIDE GSR MESSAGE FORM]
        Unload frmMsg
    
    End If



End Sub

Sub prnStaffRoster(strFullname)

    '[THIS IS THE STAFF ROSTER SUBROUTINE. IT WILL CREATE A TEMP DYNASET        ]
    '[TO CONTAIN ALL OF THE ROSTER RECORDS AND CYCLE THROUGH USING SQL          ]
    '[STATEMENTS (HOPEFULLY).                                                   ]
    
    '[BUILD ROSTER DYNASET WITH ALL RECORDS]
    Dim SQLStmt         As String
    Dim strDayKey       As String
    Dim intCounter      As Integer
    Dim intDayCount     As Integer
    Dim strClassDesc    As String
    Dim sinMinsWorked   As Single
    Dim sinIncrement    As Single
    Dim intInterval     As Integer
    Dim intRosterCount  As Integer
    Dim dateStart       As Date
    Dim dateEnd         As Date
    Dim dateDay         As Date
    Dim strStaffId      As String
    Dim strName         As String
    Dim strRoster       As String
    Dim sinMinutes      As Single
    Dim sinTotal        As Single
    Dim sinTotalAmount  As Single
    Dim sinTotalMinutes As Single
    Dim strShortDay     As String
    Dim strLastDay      As String
    Dim sinX            As Single
    Dim sinY            As Single
    Dim strNote         As String
    Dim intPlace        As Integer
    Dim flagFound       As Boolean
    
    '[SET UP ARRAY FOR HOLDING STAFF ROSTER DATA]
    flagFound = False
    '[(DAY * SHIFT)]
    Dim arrayRoster(7, 20)  As StaffType
    '[ARRAY FOR COUNTING SHIFTS ON THIS DAY]
    Dim arrayCount(7)       As Integer
    
    '[SELECT APPROPRIATE RECORDS]
    SQLStmt = "SELECT * FROM Roster ORDER BY CLASS, SHIFTSTART"
    Set DsReport = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
            
    '[CHECK CURRENT DYNASET AND PREPARE]
    If DsRoster.EOF And DsRoster.BOF Then
        DsReport.Close
        Exit Sub
    End If
    
    '[*STAFF LOOP**************************************************************]
    '[RESET MINUTES WORKED]
    Erase arrayRoster
    Erase arrayCount
    '[SHOW PROGRESS REPORT]
    Call ReportInfo(strFullname, 0)
                    
    '[=ROSTER LOOP=========================================================]
    '[MOVE TO FIRST RECORD]
    DsReport.MoveFirst
    '[CYCLE THROUGH ROSTER]
    Do While Not DsReport.EOF
        
        '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
        DsClass.AbsolutePosition = (DsReport("Class") - 1)
        If DsClass("Active") = vbChecked Then
            For intDayCount = 1 To 7
                '[-=-NOW CHECK EACH DAY TO SEE IF STAFF MEMBER IS INCLUDED IN ANY DAY-=-]
                strDayKey = "Day_" & Trim(Str(intDayCount))
                
                If InStr(DsReport(strDayKey), strFullname) > 0 Then
                    '[SET FLAG TO TRUE]
                    flagFound = True
                    '[INCREMENT COUNTER FOR THIS DAY]
                    arrayCount(intDayCount) = arrayCount(intDayCount) + 1
                    '[RESET INCREMENT - ALLOWS FOR ROSTERS WITH DIFFERING INCREMENTS]
                    dateStart = DsReport("ShiftStart")
                    dateEnd = DsReport("ShiftEnd")
                    '[ALLOW FOR NEXT DAY TIMES]
                    If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
                    '[CALCULATE INCREMENT]
                    sinIncrement = (dateEnd - dateStart) * (24 * 60)
                    '[ADD MINUTES TO STAFF ROSTER ARRAY]
                    arrayRoster(intDayCount, arrayCount(intDayCount)).Minutes = arrayRoster(intDayCount, arrayCount(intDayCount)).Minutes + sinIncrement
                    '[ADD START TIME, END TIME TO STAFF ROSTER ARRAY]
                    arrayRoster(intDayCount, arrayCount(intDayCount)).StartTime = DsReport("ShiftStart")
                    arrayRoster(intDayCount, arrayCount(intDayCount)).EndTime = DsReport("ShiftEnd")
                    '[ADD ROSTER DESCRIPTION]
                    arrayRoster(intDayCount, arrayCount(intDayCount)).Roster = DsClass("Description")
                    '[ADD STARTING DATE]
                    arrayRoster(intDayCount, arrayCount(intDayCount)).DayDate = DsDefault("StartDate") + (intDayCount - 1)
                End If
                
            Next intDayCount
        End If
        
        '[MOVE TO NEXT ROSTER RECORD]
        DsReport.MoveNext
        
    Loop
    '[=====================================================================]
    '[IF NO DATA FOUND, EXIT THE ROUTINE]
    If flagFound = False Then Exit Sub
    '[PRINT STATEMENTS HERE]
    strStaffId = DsStaff("StaffID")
    
    '[SET PRINT FONT]
    Printer.FontSize = 18
    Printer.FontName = frmRoster.GridRoster.FontName
    Printer.FontBold = True
    
    '[PRINT STAFF NAME AND ID HERE]
    Printer.Print strStaffId & " - " & strFullname
    
    '[PRINT TITLES AND COLUMN HEADINGS]
    Printer.FontName = frmRoster.GridRoster.FontName
    Printer.FontSize = frmRoster.GridRoster.FontSize
    Printer.FontBold = frmRoster.GridRoster.FontBold
    Printer.FontItalic = frmRoster.GridRoster.FontItalic
    Printer.Print "Weekly Roster Starting : " & DsDefault("StartDate")
    Printer.Print
    '[PRINT HEADINGS HERE]
    sinX = Printer.CurrentX
    sinY = Printer.CurrentY
    '[DAY AND DATE]
    Printer.CurrentX = 0
    Printer.CurrentY = sinY
    Printer.Print "Date"
    '[ROSTER NAME]
    Printer.CurrentX = Printer.Width * 0.2      '[20%]
    Printer.CurrentY = sinY
    Printer.Print "Roster"
    '[START DATE]
    Printer.CurrentX = Printer.Width * 0.4      '[40%]
    Printer.CurrentY = sinY
    Printer.Print "Start Time"
    '[END DATE]
    Printer.CurrentX = Printer.Width * 0.6      '[60%]
    Printer.CurrentY = sinY
    Printer.Print "End Time"
    '[HOURS]
    Printer.CurrentX = Printer.Width * 0.8      '[80%]
    Printer.CurrentY = sinY
    Printer.Print "Hours"
    
    Printer.Print
            
    '[CYCLE THROUGH DAYS]
    For intDayCount = 1 To 7
        '[CYCLE THROUGH SHIFTS]
        sinMinutes = 0
        For intCounter = 1 To arrayCount(intDayCount)
            '[ASSIGN ARRAY DATA TO SET VARIABLES]
            sinMinutes = arrayRoster(intDayCount, intCounter).Minutes
            strShortDay = ArrayWeek(DayOfWeek(intDayCount)).ShortDay
            strRoster = arrayRoster(intDayCount, intCounter).Roster
            dateStart = arrayRoster(intDayCount, intCounter).StartTime
            dateEnd = arrayRoster(intDayCount, intCounter).EndTime
            dateDay = arrayRoster(intDayCount, intCounter).DayDate
            
            '[PRINT DAY AND SHIFT RECORD HERE]
            sinX = Printer.CurrentX
            sinY = Printer.CurrentY
            Printer.CurrentX = 0
            Printer.CurrentY = sinY
            '[DAY AND DATE]
            If strLastDay = strShortDay Then
                Printer.Print ""
            Else
                Printer.Print strShortDay
                Printer.CurrentX = Printer.Width * 0.1  '[10%]
                Printer.CurrentY = sinY
                Printer.Print Format(dateDay, "Short Date")
            End If
            '[ROSTER NAME]
            Printer.CurrentX = Printer.Width * 0.2      '[20%]
            Printer.CurrentY = sinY
            Printer.Print strRoster
            '[START DATE]
            Printer.CurrentX = Printer.Width * 0.4      '[40%]
            Printer.CurrentY = sinY
            Printer.Print Format(dateStart, "Medium Time")
            '[OLD TIME FORMAT] -> "hh:mm AMPM"
            '[END DATE]
            Printer.CurrentX = Printer.Width * 0.6      '[60%]
            Printer.CurrentY = sinY
            Printer.Print Format(dateEnd, "Medium Time")
            '[OLD TIME FORMAT] -> "hh:mm AMPM"
            '[HOURS]
            Printer.CurrentX = Printer.Width * 0.8      '[80%]
            Printer.CurrentY = sinY
            Printer.Print Format(sinMinutes / 60, "###0.00")
            '[ADD TOTALS]
            sinTotalMinutes = sinTotalMinutes + sinMinutes
            strLastDay = strShortDay
        Next intCounter
        If sinMinutes > 0 Then Printer.Print
    Next intDayCount
    '[=====================================================================]
    '[=====================================================================]
    
    '[PRINT CLOSING TOTALS AND END OF PAGE]
    sinX = Printer.CurrentX
    sinY = Printer.CurrentY
    '[TOTALS HEADING]
    Printer.CurrentX = Printer.Width * 0.6      '[60%]
    Printer.CurrentY = sinY
    Printer.Print "Total Hours"
    '[HOURS]
    Printer.CurrentX = Printer.Width * 0.8      '[80%]
    Printer.CurrentY = sinY
    Printer.Print Format(sinTotalMinutes / 60, "###0.00")
    
    Printer.Print
    Printer.Print "NOTES"
    Printer.CurrentX = 0
    Printer.Line -((Printer.Width * 0.95), Printer.CurrentY), vbBlack
    Printer.CurrentX = 0

    '[ROUTINE TO CUT NOTE MEMO FIELD UP INTO SEGMENTS SMALL ENOUGH TO PRINT]
    intPlace = 1
    If Printer.TextWidth(DsStaff("Note")) < Printer.Width * 0.95 Then
            Printer.Print DsStaff("Note")
    Else
        Do While (intPlace + intCounter) <= Len(DsStaff("Note"))
            intCounter = intCounter + 1
            strNote = Mid$(DsStaff("Note"), intPlace, intCounter)
            If Printer.TextWidth(strNote) > Printer.Width * 0.95 Or Len(strNote) > 254 Then
                '[TEXT HAS MAXED AT LIMIT, PRINT AND RESET]
                Printer.Print strNote
                intPlace = intCounter + 1
                intCounter = 1
            ElseIf (intPlace + intCounter) > Len(DsStaff("Note")) Then
                Printer.Print strNote
            End If
        Loop
    End If
    Printer.CurrentX = 0
    Printer.Line -((Printer.Width * 0.95), Printer.CurrentY), vbBlack
    
    '[CLOSING LINE]
    Printer.Print
    Printer.CurrentX = 0
    Printer.Print "Printed : "; Format(Date, "Long Date"); " at "; Format(Time, "Long Time")    '[PRINT CURRENT DATE/TIME ON END OF REPORT - LONG DATE FORMAT]

    '[END PRINTING OF DOCUMENT WITH NEWPAGE]
    Printer.NewPage
    
    '[------------------------------MAIN PRINTING ROUTINE------------------------------]
    
    '[CLOSE TEMPORARY DYNASET]
    DsReport.Close
    
End Sub

Sub procStaffReport()
    
    '[COMMAND TO PROCESS EACH STAFF MEMBER AND COLLECT INFO ABOUT EACH, SENDING IT IT THE REPORT FORM]
    '[USE THE REPORT FORM TO ADDRESS ANY OF THESE PROBLEMS]
    Dim Msg
    Dim Style
    Dim Title
    Dim Response
    
    '[SHOW WARNING FORM]
    Msg = "The staff report lists roster, hour and cost figures for each staff member in your staff list." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Because this routine has to perform multiple comparisions and searches, it may take a few minutes to complete, depending upon the number of rosters, staff and the speed of your computer." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to continue and produce the report ?"
    Style = vbYesNo                     ' Define buttons.
    Title = "Confirmation Required"
    Response = gsrMsg(Msg, Style, Title)
    If Response = vbYes Then
        '[SHOW PLEASE WAIT MESSAGE]
        '[SET STATUS BAR MESSAGE-----------------------------------------------------------]
        StatusBar "... Please Wait ..."
        '[---------------------------------------------------------------------------------]
        mdiMain.panelStatusBar.Refresh
        
        '[QUICKEST WAY TO REINITIALISE GRID AND FORM IS TO UNLOAD IT THEN RELOAD IT]
        If frmReport.Visible Then Unload frmReport
        frmReport.Show
        frmReport.GridReport.Cols = 5
        
        '[SET REPORT FORM FONT]
        Set frmReport.GridReport.Font = frmRoster.GridRoster.Font
        '[SET GRID REPORT TITLES AND SIZE]
        frmReport.GridReport.Row = 0
        frmReport.GridReport.Col = 0: frmReport.GridReport.Text = "Staff ID"
        frmReport.GridReport.Col = 1: frmReport.GridReport.Text = "Name"
        frmReport.GridReport.Col = 2: frmReport.GridReport.Text = "Roster"
        frmReport.GridReport.Col = 3: frmReport.GridReport.Text = "Hours"
        frmReport.GridReport.Col = 4: frmReport.GridReport.Text = "Amount"
        
        '[SHOW GSR MESSAGE FORM]
        Msg = "GSR is now processing your roster files and producing the staff detail report." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "The report will detail which problems (if any) have been found with your rosters.  Double-clicking on the error listed in the report will take you to the source of the problem." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "This may take a few minutes."
        Style = vbInformation            ' Define buttons.
        Title = "Staff Report"
        Response = gsrMsg(Msg, Style, Title)
        frmMsg.GaugeProgress.Visible = True
        frmMsg.labelInfo.Visible = True
        frmMsg.ZOrder
        frmMsg.Refresh
        
        '[CALL STAFF REPORT SUBROUTINE]
        StaffReport
        
        '[HIDE GSR MESSAGE FORM]
        Unload frmMsg
        '[SET REPORT TITLE]
        frmReport.LabelTitle = "GSR Staff Report"
        '[RESIZE FORM TO MAXIMUM SIZE]
        frmReport.WindowState = 2
        
        '[CLOSE FORM IF NO ERRORS FOUND AND POPUP MESSAGE]
        If frmReport.GridReport.Rows <= 2 Then
            frmMsg.GaugeProgress.Visible = False
            frmMsg.labelInfo.Visible = False
            Unload frmReport
            Msg = "GSR has found no staff members or rosters to process." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Running the Staff Report command regularly will enable you to keep track of general roster costs." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Remember that the exception report is only produced for rosters which are 'enabled' on the Control form."
            Style = vbOKOnly ' Define buttons.
            Title = "All Clear !"
            Response = gsrMsg(Msg, Style, Title)
        End If
    End If

End Sub


Sub resReportForm()

    '[TEMP VARIABLES SO WE CAN CATCH ILLEGAL WIDTHS]
    Dim sinWidth        As Single
    Dim sinHeight       As Single
    
    '[IF FORM IS MINIMISED THEN EXIT THIS ROUTINE]
    If frmReport.WindowState = 1 Then Exit Sub
    '[RESIZE GRID AND LIST BOXES AND ARRANGE CONTROLS ON FORM]
    
    '[REPORT GRID]
    frmReport.GridReport.Left = 150
    frmReport.GridReport.Top = frmReport.PanelToolBar.Height + 100
    sinWidth = frmReport.Width - 400
    If sinWidth > 0 Then frmReport.GridReport.Width = sinWidth
    sinHeight = frmReport.Height - frmReport.PanelToolBar.Height - frmReport.GridReport.Top - 50
    If sinHeight > 0 Then frmReport.GridReport.Height = sinHeight
   
    '[RESIZE GRID AND COLS ACCORDING TO REPORT TYPE]
    Select Case frmReport.GridReport.Cols
    Case 5
        '[STAFF ID]
        frmReport.GridReport.ColWidth(0) = (frmReport.GridReport.Width) * 0.15     '[15 %]
        '[STAFF NAME]
        frmReport.GridReport.ColWidth(1) = (frmReport.GridReport.Width) * 0.25     '[25 %]
        '[ROSTER NAME]
        frmReport.GridReport.ColWidth(2) = (frmReport.GridReport.Width) * 0.25     '[25 %]
        '[HOURS]
        frmReport.GridReport.ColWidth(3) = (frmReport.GridReport.Width) * 0.15     '[15 %]
        frmReport.GridReport.ColAlignment(3) = 1
        '[AMOUNT]
        frmReport.GridReport.ColWidth(4) = (frmReport.GridReport.Width) * 0.15     '[15 %]
        frmReport.GridReport.ColAlignment(4) = 1
    Case 6
        frmReport.GridReport.ColWidth(0) = frmReport.ImageWarning(0).Width
        '[ROSTER NAME]
        frmReport.GridReport.ColWidth(1) = (frmReport.GridReport.Width - frmReport.GridReport.ColWidth(0)) * 0.1     '[10 %]
        '[DATE]
        frmReport.GridReport.ColWidth(2) = (frmReport.GridReport.Width - frmReport.GridReport.ColWidth(0)) * 0.1     '[10 %]
        '[TIME]
        frmReport.GridReport.ColWidth(3) = (frmReport.GridReport.Width - frmReport.GridReport.ColWidth(0)) * 0.1     '[10 %]
        '[STAFF NAME]
        frmReport.GridReport.ColWidth(4) = (frmReport.GridReport.Width - frmReport.GridReport.ColWidth(0)) * 0.2     '[20 %]
        '[EXCEPTION]
        frmReport.GridReport.ColWidth(5) = (frmReport.GridReport.Width - frmReport.GridReport.ColWidth(0)) * 0.45    '[45 %]
    Case Else
    End Select

End Sub

Sub ShadeForm(Frm As Form)

' Description
'     Draws the "install-type" shaded background on a form
'
' Paramaters
'     Name                 Type     Value
'     -----------------------------------------------------------
'     Frm                  Form     The form to draw the shade on
'
' Returns
'     Nothing
'
' Last modified by Gord MacLeod 05.02.96

Dim i%
Dim NumberOfRects As Integer
Dim GradColor As Long
Dim GradValue As Integer

   Frm.ScaleMode = 3
   Frm.DrawStyle = 6
   Frm.DrawWidth = 2
   Frm.AutoRedraw = True
   
   NumberOfRects = 64
    
   For i% = 1 To 64
   
      GradValue = 255 - (i% * 4 - 1)
       
      ' Put GradValue in Red and Green for a different look
      GradColor = RGB(0, 0, GradValue)
       
      ' Draw the line
      Frm.Line (0, Frm.ScaleHeight * (i% - 1) / 64)-(Frm.ScaleWidth, Frm.ScaleHeight * i% / 64), GradColor, BF
    
   Next i%

   Frm.Refresh

End Sub


Function LocateClass(strClassDesc) As Boolean

    '[LOCATE PASSED DESC IN CLASS DYNASET]
    Dim SQLStmt
    SQLStmt = "[Description]='" & strClassDesc & "'"

    DsClass.FindFirst SQLStmt
    
    If DsClass.NoMatch Then LocateClass = False Else LocateClass = True
    
End Function
Function CheckStaffDay(strFullString, intCol)

    '[CHECK THAT THE STAFF MEMBER PASSED IS AVAILABLE ON THE SELECTED DAY]
    Dim SQLStmt         As String
    Dim strSurname      As String
    Dim strFirstName    As String
    Dim intDelimiter    As Integer
    Dim strBookmark     As String
    Dim strDayKey       As String
    Dim flagResult      As Boolean
    
    '[STORE CURRENT STAFF LOCATION]
    strBookmark = DsStaff.Bookmark
    
    intDelimiter = InStr(strFullString, ",")
    strSurname = Trim(Left(strFullString, intDelimiter - 1))
    strFirstName = Trim(Mid(strFullString, intDelimiter + 1))
    
    '[LOCATE LASTNAME, FIRSTNAME IN DYNASET]
    SQLStmt = "LastName = '" & strSurname & "' AND FirstName = '" & strFirstName & "'"
    DsStaff.FindFirst SQLStmt
    
    '[CHECK FOR RESULT OF SEARCH]
    strDayKey = "Day_" & Trim(Str$(intCol - 1))
    If DsStaff(strDayKey) = vbChecked Then flagResult = True Else flagResult = False

    '[RESTORE CURRENT STAFF LOCATION]
    DsStaff.Bookmark = strBookmark

    CheckStaffDay = flagResult

End Function

Public Sub prnReportHeadings(sngLineHeight, intLineCount, sngPageWidth, intPage, strReportTitle, GridPrint, intStartCol, intBoxState, strRightText)

    '[--------------> X <--------------]
    '[  .
    '[  .
    '[ Y
    '[  .
    '[  .

    '[RESET LINE COUNT FOR START OF NEW PAGE]
    intLineCount = 0

    '[PRINT PAGE HEADINGS, FIRST LINE WHITE ON BLACK, REMAINING LINES STANDARD BLACK]
    Printer.FillStyle = 1
    Printer.FillColor = vbWhite
    
    '[SET PRINT FONT]
    Printer.FontSize = 18
    Printer.FontName = GridPrint.FontName
    Printer.FontBold = True
    Printer.Line (0, 0)-(sngPageWidth, Printer.TextHeight("Dummy String") * 2), vbBlack, B
    Printer.FillStyle = 1
    Printer.CurrentX = sngPageWidth * 0.01      '[LEFT MARGIN]
    Printer.CurrentY = sngLineHeight * 0.05     '[TOP MARGIN]
    Printer.ForeColor = vbBlack
    '[REPORT TITLE]
    Printer.Print strReportTitle
    Printer.FontSize = 14
    Printer.CurrentY = sngLineHeight * 0.05     '[TOP MARGIN]
    Printer.CurrentX = (Printer.Width * 0.89) - Printer.TextWidth(strRightText)
    '[RIGHT JUSTIFIED TEXT]
    Printer.Print strRightText
    '[PRINT TITLES AND COLUMN HEADINGS]
    Printer.FontName = GridPrint.FontName
    Printer.FontSize = GridPrint.FontSize
    Printer.FontBold = GridPrint.FontBold
    Printer.FontItalic = GridPrint.FontItalic

    '[INCREMENT LINE COUNTER AND SET X POSITION]
    intLineCount = intLineCount + 3
    '[TITLE DESCRIPTION LINES]
    '[-----REGISTERED USER AND PAGE NUMBER-----]
    Printer.CurrentX = sngPageWidth * 0.01:   Printer.CurrentY = (sngLineHeight * intLineCount):    Printer.Print DsDefault("RegUser")
    Printer.CurrentX = sngPageWidth * 0.9: Printer.CurrentY = (sngLineHeight * intLineCount):     Printer.Print "Page: " & intPage
    intLineCount = intLineCount + 1

    '[-----COLUMN HEADINGS]
    Call prnGridRow(0, sngLineHeight, intLineCount, sngPageWidth, GridPrint, intStartCol, intBoxState)


End Sub


Sub ExceptionReport()

    '[THIS IS THE EXCEPTION REPORT SUBROUTINE. IT WILL CREATE A TEMP DYNASET    ]
    '[TO CONTAIN ALL OF THE ROSTER RECORDS AND CYCLE THROUGH USING SQL          ]
    '[STATEMENTS (HOPEFULLY).                                                   ]
    
    '[BUILD ROSTER DYNASET WITH ALL RECORDS]
    Dim SQLStmt         As String
    Dim strBookmark     As String
    Dim strFullname     As String
    Dim strDayKey       As String
    Dim intCounter      As Integer
    Dim intDayCount     As Integer
    Dim strClassDesc    As String
    Dim sinMinsWorked   As Single
    Dim sinIncrement    As Single
    Dim intInterval     As Integer
    Dim intTotal        As Integer
    '[STORE CURRENT STAFF LOCATION]
    strBookmark = DsStaff.Bookmark
    
    '[SELECT APPROPRIATE RECORDS]
    SQLStmt = "SELECT * FROM Roster"
    Set DsReport = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
            
    '[CHECK CURRENT DYNASET AND PREPARE]
    If (DsRoster.EOF And DsRoster.BOF) Or (DsReport.EOF And DsReport.EOF) Then
        DsReport.Close
        Exit Sub
    End If
    
    '[*STAFF LOOP**************************************************************]
    DsStaff.MoveLast
    DsReport.MoveLast
    
    '[MOVE TO FIRST STAFF RECORD]
    DsStaff.MoveFirst
    intTotal = DsStaff.RecordCount * DsReport.RecordCount
    intCounter = 1
    
    '[CYCLE THROUGH STAFF LIST]
    Do While Not DsStaff.EOF
        '[RESET MINUTES WORKED]
        sinMinsWorked = 0
        '[DETERMINE FULL NAME]
        strFullname = DsStaff("LastName") & ", " & DsStaff("FirstName")
        '[SHOW PROGRESS REPORT]
        Call ReportInfo(strFullname, 0)
                        
        '[=ROSTER LOOP=========================================================]
        '[MOVE TO FIRST RECORD]
        DsReport.MoveFirst
      
      
        '[CYCLE THROUGH ROSTER]
        Do While Not DsReport.EOF
            
            '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
            DsClass.AbsolutePosition = DsReport("Class") - 1
            
            '[PROGRESS BAR]
            Call ReportProgressBar((intCounter / intTotal) * 100)

            If DsClass("Active") = vbChecked Then
                
                '[-=-FIRST CHECK - STAFF NOT AVAIABLE-=-]
                Call exCheckStaffAvailable(sinMinsWorked, sinIncrement)
                
            End If
            
            DsReport.MoveNext
            intCounter = intCounter + 1
            
        Loop
        '[=====================================================================]
        '[-=-THIRD CHECKCHECK MINUTES WORKED-=-]
        If ((sinMinsWorked / 60) > DsStaff("MaxHours")) And DsStaff("maxHours") > 0 Then Call exAddNewException(constWarning, 0, 0, 0, strFullname, Str$(Format(sinMinsWorked / 60, "#0.00")) & " hours allocated/" & Str$(DsStaff("MaxHours")) & " hours allowed.")
        If ((sinMinsWorked / 60) < DsStaff("MinHours")) Then Call exAddNewException(constWarning, 0, 0, 0, strFullname, Str$(Format(sinMinsWorked / 60, "#0.00")) & " hours allocated/" & Str$(DsStaff("MinHours")) & " hours required.")
        
        DsStaff.MoveNext
    Loop
    '[*************************************************************************]
    
    
    '[CLEAR REPORT INFO]
    Call ReportInfo("", 0)
    '[CLEAR PROGRESS BAR]
    Call ReportProgressBar(0)
    
    '[RETURN TO STAFF BOOKMARK]
    DsStaff.Bookmark = strBookmark
    
    '[CLOSE TEMPORARY DYNASET]
    DsReport.Close

    '[MOVE TO FIRST ROW IN REPORT]
    frmReport.GridReport.Row = 1
    frmReport.GridReport.Col = 0
    
    '[CALL SUBROUTINE TO RESIZE FORM]
    Call resReportForm

End Sub


Sub exCheckRosterConflict(strFullname, intDayCount, sinIncrement)

    '[SUBROUTINE TO CHECK FOR CONFLICTS IN THE CURRENT ROSTER]
    '[THIS SUBROUTINE WILL TAKE A LONG TIME TO PROCESS AS IT HAS TO CHECK EVERY RECORD]
    Dim SQLStmt         As String
    Dim xSQLStmt        As String
    Dim strDayKey       As String
    Dim intCounter      As Integer
    Dim strClassDesc    As String
    Dim intClass        As Integer
    Dim strTime         As String
    Dim dateTemp        As Date
    Dim strBookmark     As String
    Dim dateStart       As Date
    Dim dateEnd         As Date
    
    '[EXIT IF RECORD IS NULL]
    If IsNull(DsReport("ShiftStart")) Or IsNull(DsReport("ShiftEnd")) Then Exit Sub
    
    '[SAVE RECORD POSITION]
    strBookmark = DsReport.Bookmark
    '[SET DAY KEY STRING]
    strDayKey = "Day_" & Trim(Str(intDayCount))
    '[SET TIME TO CHECK AGAINST]
    strTime = DsReport("ShiftStart")
    dateStart = CDate(DsReport("ShiftStart"))
    dateEnd = CDate(DsReport("ShiftEnd"))
    
    '[SET CLASS ID]
    intClass = DsReport("Class")
    '[FIND CLASS DESCRIPTION]
    xSQLStmt = "ID = " & intClass
    DsClass.FindFirst xSQLStmt
    strClassDesc = DsClass("Description")
    
    '[SET SQL STATEMENT FOR LOCATING RECORDS]
    SQLStmt = "[Class] <> " & DsReport("Class") & " AND [ShiftStart] <= TIMEVALUE('" & dateEnd & "') AND [ShiftEnd] >= TIMEVALUE('" & dateStart & "')"
    '[MOVE TO FIRST RECORD TO MATCH THIS CONDITION]
    DsReport.FindFirst SQLStmt
    
    '[DO WHILE LOOP TO KEEP GOING WHILE RECORDS MATCH]
    Do While Not DsReport.NoMatch
        '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
        DsClass.AbsolutePosition = DsReport("Class") - 1
        '[DEBUG]
        Debug.Print strFullname, DsReport("Class"), intDayCount, DsReport("ShiftStart"), DsReport("ShiftEnd")
        '[CHECK FOR CONTENTS OF NEW RECORD]
        If Not (IsNull(DsReport("ShiftStart")) Or IsNull(DsReport("ShiftEnd"))) Then
            If DsClass("Active") = vbChecked Then
                '[FIND OUT IF THE TIME MATCHES OR FALLS WITHIN THE INCREMENT ALLOWED]
                If CDate(DsReport("ShiftStart")) >= CDate(strTime) And CDate(strTime) <= DsReport("ShiftEnd") Then
                    '[CHECK TO SEE IF THE NAME IS IN THE CELL]
                    If InStr(DsReport(strDayKey), strFullname) > 0 Then
                        '[NAME IS FOUND IN CELL, PRODUCE EXCEPTION REPORT WARNING]
                        intClass = DsReport("Class")
                        strTime = DsReport("ShiftStart")
                        '[CALL EXCEPTION REPORT]
                        Call exAddNewException(constCritical + 1, intClass, intDayCount, strTime, strFullname, "Conflict with " & strClassDesc & " roster.")
                    End If
                End If
                '[FIND NEXT RECORD]
            End If
        End If
        DsReport.FindNext SQLStmt
    Loop
    
    '[RESTORE RECORD POSITION]
    DsReport.Bookmark = strBookmark

End Sub

Sub exCheckStaffAvailable(sinMinutesWorked, sinIncrement)

    '[SUBROUTINE TO CHECK THAT THE STAFF MEMBER LISTED IN THE ROSTER IS AVAILABLE]
    Dim strFullname     As String
    Dim strDayKey       As String
    Dim intCounter      As Integer
    Dim intDayCount     As Integer
    Dim strClassDesc    As String
    Dim intClass        As Integer
    Dim strTime         As String
    Dim dateTemp        As Date
    Dim intBookmark     As Integer
    Dim strLastSearchName       As String
    Dim strBookmark     As String
    Dim dateStart       As Date
    Dim dateEnd         As Date
    
    '[DETERMINE FULL NAME]
    strFullname = DsStaff("LastName") & ", " & DsStaff("FirstName")
    '[SAVE BOOKMARK POSITION]
    strBookmark = DsReport.Bookmark
    
    '[RESET INCREMENT - ALLOWS FOR ROSTERS WITH DIFFERING INCREMENTS]
    If IsNull(DsReport("ShiftStart")) Or IsNull(DsReport("ShiftEnd")) Then
        dateStart = Format(Now, "Short Date")
        dateEnd = Format(Now, "Short Date")
    Else
        dateStart = DsReport("ShiftStart")
        dateEnd = DsReport("ShiftEnd")
    End If
    
    '[ALLOW FOR NEXT DAY TIMES]
    If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
    
    '[CALCULATE INCREMENT]
    sinIncrement = (dateEnd - dateStart) * (24 * 60)

    '[DEBUG]
    Debug.Print strFullname

    '[CYCLE ACROSS DAYS]
    For intDayCount = 1 To 7
        '[ASSIGN DAY KEY]
        strDayKey = "Day_" & Trim(Str$(intDayCount))
        
        '[CHECK ALL NAMES IN CELL AGAINST NAMES IN STAFF LIST]
        If Len(Trim(DsReport(strDayKey))) > 0 Then Call exCheckNameInStaffList(DsReport(strDayKey), intDayCount)
        
        If DsStaff(strDayKey) = vbUnchecked And InStr(DsReport(strDayKey), strFullname) > 0 Then
            '[NAME FOUND, STAFF NOT AVAILABLE THOUGH, ADD TO EXCEPTION REPORT]
            intClass = DsReport("Class")
            strTime = DsReport("ShiftStart")
            Call exAddNewException(constCritical, intClass, intDayCount, strTime, strFullname, "Unavailable staff member in roster")
            '[MINUTES WORKED]
            sinMinutesWorked = sinMinutesWorked + sinIncrement
            Call exCheckStaffInClass(strFullname, intDayCount)
            Call exCheckRosterConflict(strFullname, intDayCount, sinIncrement)
        ElseIf InStr(DsReport(strDayKey), strFullname) > 0 Then
            '[MINUTES WORKED]
            sinMinutesWorked = sinMinutesWorked + sinIncrement
            Call exCheckStaffInClass(strFullname, intDayCount)
            Call exCheckRosterConflict(strFullname, intDayCount, sinIncrement)
        End If
    Next intDayCount

End Sub


Sub exAddNewException(intWarningLevel, intClass, intDayCount, strTime, strFullname, strMessage)

    '[SUBROUTINE TO PLACE NEW MESSAGE IN REPORT GRID]
    Dim strClassDescription     As String
    Dim strDayName              As String
    Dim intStartDay             As Integer
    Dim intCounter              As Integer
    Dim SQLStmt                 As String
    
    intStartDay = DsDefault("StartDay")
        
    '[SET DAY NAME - LONG FORMAT]
    If intDayCount = 0 Then
        strDayName = "-"
    Else
        If intStartDay + (intDayCount - 1) > 7 Then
            strDayName = ArrayWeek(intDayCount - 7 + intStartDay - 1).ShortDay
        Else
            strDayName = ArrayWeek(intStartDay + (intDayCount - 1)).ShortDay
        End If
    End If
    
    '[SET ROSTER NAME]
    If intClass > 0 Then
        SQLStmt = "ID = " & intClass
        DsClass.FindFirst SQLStmt
        strClassDescription = DsClass("Description")
    Else
        strClassDescription = "-"
    End If
    
    '[TIME FORMAT]
    If strTime = "0" Then
        strTime = "-"
    Else
        strTime = Format(strTime, "Medium Time")
        '[OLD TIME FORMAT] -> "hh:mm AMPM"
    End If
    
    '[TRIM MESSAGE]
    strMessage = Trim(strMessage)
    
    frmReport.GridReport.AddItem intWarningLevel & Chr$(vbKeyTab) & strClassDescription & Chr$(vbKeyTab) & strDayName & Chr$(vbKeyTab) & strTime & Chr$(vbKeyTab) & strFullname & Chr$(vbKeyTab) & strMessage, (frmReport.GridReport.Rows - 1)
    frmReport.GridReport.Row = (frmReport.GridReport.Rows - 2)
    frmReport.GridReport.Col = 0
    
    '[SET WARNING ICON/PICTURE - ADJUST IF > 2]
    If intWarningLevel = 5 Then intWarningLevel = 1
    If intWarningLevel > 2 Then intWarningLevel = 2
    frmReport.GridReport.Picture = frmReport.ImageWarning(intWarningLevel).Picture

End Sub

Sub exCheckStaffInClass(strFullname, intDayCount)
    
    Dim strClassKey     As String
    Dim intClass        As Integer
    Dim strTime         As String
    
    '[CHECK NULL START/END TIME]
    If IsNull(DsReport("ShiftStart")) Or IsNull(DsReport("ShiftEnd")) Then Exit Sub
    
    '[-=-SECOND CHECK - STAFF NOT IN THIS CLASS-=-]
    strClassKey = "Class_" & Trim(Str$(DsReport("Class")))
    intClass = DsReport("Class")
    strTime = DsReport("ShiftStart")
    If DsStaff(strClassKey) = vbUnchecked Then
        '[STAFF MEMBER NOT AVAILABLE IN THIS CLASS]
        Call exAddNewException(constSerious, intClass, intDayCount, strTime, strFullname, "Staff member not in this class")
    End If

End Sub

Function gsrMsg(strMsg, Style, strTitle)

    '[PROCESS PASSED VARIABLES AND SETUP MSG FORM]
    Load frmMsg
    
    '[SET TITLE AND MESSAGE]
    frmMsg.LabelTitle = strTitle
    frmMsg.LabelMessage = Trim(strMsg & " ")
    
    '[SET STYLE OF BUTTONS]
    Select Case Style
    Case vbYesNo
        frmMsg.cmdYes.Visible = True
        frmMsg.cmdNo.Visible = True
        frmMsg.cmdCancel.Visible = False
        frmMsg.cmdOK.Visible = False
        
        frmMsg.cmdNo.Default = True
        
        frmMsg.cmdYes.Left = (frmMsg.Width / 3) - (frmMsg.cmdYes.Width / 2)
        frmMsg.cmdNo.Left = ((frmMsg.Width / 3) * 2) - (frmMsg.cmdNo.Width / 2)
        
    Case vbYesNoCancel
        frmMsg.cmdYes.Visible = True
        frmMsg.cmdNo.Visible = True
        frmMsg.cmdCancel.Visible = True
        frmMsg.cmdOK.Visible = False
    
        frmMsg.cmdCancel.Default = True
        
        frmMsg.cmdYes.Left = (frmMsg.Width / 4) - (frmMsg.cmdYes.Width / 2)
        frmMsg.cmdNo.Left = ((frmMsg.Width / 4) * 2) - (frmMsg.cmdNo.Width / 2)
        frmMsg.cmdCancel.Left = ((frmMsg.Width / 4) * 3) - (frmMsg.cmdCancel.Width / 2)
    
    Case vbOKOnly
        frmMsg.cmdYes.Visible = False
        frmMsg.cmdNo.Visible = False
        frmMsg.cmdCancel.Visible = False
        frmMsg.cmdOK.Visible = True
        
        frmMsg.cmdOK.Default = True
        
        frmMsg.cmdOK.Left = (frmMsg.Width / 2) - (frmMsg.cmdOK.Width / 2)
    
    Case vbInformation
        frmMsg.cmdYes.Visible = False
        frmMsg.cmdNo.Visible = False
        frmMsg.cmdCancel.Visible = False
        frmMsg.cmdOK.Visible = False
    
    Case vbQuestion
        frmMsg.cmdYes.Visible = False
        frmMsg.cmdNo.Visible = False
        frmMsg.cmdOK.Visible = True
        frmMsg.cmdCancel.Visible = True
        frmMsg.TextNote.Visible = True
        frmMsg.TextNote.Text = Trim(strMsg & " ")
    
        frmMsg.cmdOK.Left = (frmMsg.Width / 3) - (frmMsg.cmdOK.Width / 2)
        frmMsg.cmdCancel.Left = ((frmMsg.Width / 3) * 2) - (frmMsg.cmdCancel.Width / 2)
    
    Case Else
        frmMsg.cmdYes.Visible = True
        frmMsg.cmdNo.Visible = True
        frmMsg.cmdCancel.Visible = False
        frmMsg.cmdOK.Visible = False
        
        frmMsg.cmdNo.Default = True
        
        frmMsg.cmdYes.Left = (frmMsg.Width / 3) - (frmMsg.cmdYes.Width / 2)
        frmMsg.cmdNo.Left = ((frmMsg.Width / 3) * 2) - (frmMsg.cmdNo.Width / 2)
    
    End Select
    
    '[BEEP]
    Beep
    '[SHOW FORM MODAL]
    If Style = vbInformation Then
        frmMsg.Show
    Else
        frmMsg.Show 1
    End If
    '[RETURN VALUE IN gsrReturn]
    gsrMsg = gsrReturn

End Function

Sub prnGrid(strReportTitle, GridPrint As Grid, intStartCol, intBoxState, strRightText, flagNote)

    Dim sngLineHeight           '[Height of each printed line]
    Dim intLinesPerPage         '[Number of Lines which can be printed per page]
    Dim sngPageWidth            '[Width of Printable Page Area]
    Dim intLineCount            '[Number of lines already printed on this page]
    Dim intRowToPrint           '[Number of Row to Print]
    Dim intListCounter          '[File List Counter]
    Dim intPage                 '[Page Number]
    Dim intPlace                '[Place within string]
    Dim strNote                 '[Text string to hold note contents]
    Dim intCounter
    
    '[SET PRINT FONT]
    Printer.FontName = GridPrint.FontName
    Printer.FontSize = GridPrint.FontSize
    Printer.FontBold = GridPrint.FontBold
    Printer.FontItalic = GridPrint.FontItalic

    '[PRINTER COLOUR AND TEXT HEIGHT]
    sngLineHeight = Printer.TextHeight("Dummy String") * 1.05    '[GIVE SMALL MARGIN FOR LINE DRAWING]
    sngPageWidth = Printer.ScaleWidth * 0.95
    intLineCount = 0
    intLinesPerPage = Int(Printer.ScaleHeight / sngLineHeight) - 4  '[NUMBER OF LINES TO PRINT PER PAGE]
    Printer.FillStyle = 1
    Printer.ForeColor = vbBlack
    intPage = 1

    '[CLEAR PAGE AND PRINT PAGE HEADINGS]
    Call prnReportHeadings(sngLineHeight, intLineCount, sngPageWidth, intPage, strReportTitle, GridPrint, intStartCol, intBoxState, strRightText)

    '[------------------------------MAIN PRINTING ROUTINE------------------------------]
    For intRowToPrint = 1 To (GridPrint.Rows - 1)
        Call prnGridRow(intRowToPrint, sngLineHeight, intLineCount, sngPageWidth, GridPrint, intStartCol, intBoxState)
        Call ReportProgressBar(intRowToPrint / (GridPrint.Rows - 1) * 100)
        Call ReportInfo("Printing " & strReportTitle, 0)
        
        '[CHECK FOR END OF PAGE]
        If intLineCount >= intLinesPerPage Then
            '[CLOSING LINE]
            Printer.Line (0, Printer.CurrentY)-(sngPageWidth, Printer.CurrentY), vbBlack
            '[MUST USE END DOC BECAUSE OF BUG? IN NEWPAGE PROC.]
            'Printer.EndDoc
            '[TRY NEWPAGE--------------------------------------------------------------]
            Printer.NewPage
            '[-------------------------------------------------------------------------]
            Printer.FontName = GridPrint.FontName
            Printer.FontSize = GridPrint.FontSize
            Printer.FontBold = GridPrint.FontBold
            Printer.FontItalic = GridPrint.FontItalic
            Printer.FillStyle = 1
            Printer.ForeColor = vbBlack
            intPage = intPage + 1
            Call prnReportHeadings(sngLineHeight, intLineCount, sngPageWidth, intPage, strReportTitle, GridPrint, intStartCol, intBoxState, strRightText)
        End If
    Next intRowToPrint


    '[ROUTINE TO CUT NOTE MEMO FIELD UP INTO SEGMENTS SMALL ENOUGH TO PRINT]
    Printer.Print ""
    
    If flagNote = True And Not IsNull(DsClass("Note")) Then
        '[SET STARTING PRINTER POSITION]
        intLineCount = intLineCount + 2
        Printer.CurrentY = intLineCount * sngLineHeight
                
        '[COUNT RETURN CHARS IN THIS STRING]
        intLineCount = intLineCount + LineCount(DsClass("Note") + " ")
        intPlace = 1
        
        If Printer.TextWidth(DsClass("Note")) < Printer.Width * 0.95 Then
            Printer.Print DsClass("Note")
        Else
            Do While (intPlace + intCounter) <= Len(DsClass("Note"))
                intCounter = intCounter + 1
                strNote = Mid$(DsClass("Note"), intPlace, intCounter)
                If Printer.TextWidth(strNote) > Printer.Width * 0.95 Or Len(strNote) > 254 Then
                    '[TEXT HAS MAXED AT LIMIT, PRINT AND RESET]
                    Printer.Print strNote
                    intPlace = intCounter + 1
                    intCounter = 1
                    intLineCount = intLineCount + 1
                ElseIf (intPlace + intCounter) > Len(DsClass("Note")) Then
                    Printer.Print strNote
                    intLineCount = intLineCount + 1
                End If
            Loop
        End If
    End If

    '[CLOSING LINE]
    Printer.Line (0, sngLineHeight * (intLineCount + 1))-(sngPageWidth, sngLineHeight * (intLineCount + 1)), vbBlack
    Printer.CurrentX = 0
    Printer.Print "Printed : "; Format(Date, "Long Date"); " at "; Format(Time, "Long Time")    '[PRINT CURRENT DATE/TIME ON END OF REPORT - LONG DATE FORMAT]

    '[END PRINTING OF DOCUMENT]
    Printer.EndDoc
    '[------------------------------MAIN PRINTING ROUTINE------------------------------]

End Sub

Sub prnGridRow(intRowToPrint, sngLineHeight, intLineCount, sngPageWidth, GridPrint, intStartCol, intBoxState)

    Dim sngColumnWidth              '[WIDTH OF EACH COLUMN]
    Dim intColumn                   '[COLUMN COUNTER]
    Dim strCellText                 '[CELL TEXT TO PRINT AND CHECK FOR LENGTH]
    Dim BoxState                    '[FLAG FOR STATE OF PRINTING BOXES]
    Dim sinBoxHeight                '[HEIGHT OF EACH BOUNDING BOX (ALLOWS FOR MULTIPLE ROWS]
    Dim intBoxRows                  '[NUMBER OF ROWS IN THE BOX]
    Dim intCounter                  '[COUNTER FOR NUMBER OF NAMES IN CELL]
    Dim strCellName     As String   '[SINGLE NAME EXTRACTED FROM CELL]
    Dim intStart        As Integer  '[STARTING POSITION OF vbKeyEnter IN CELL}
    Dim sinLinePos      As Single   '[ACCUMULATED POSITION ON THE LINE]
    Dim sinGridWidth    As Single   '[ACTUAL WIDTH OF ALL COLUMNS IN GRID]
    Dim flagDateNeeded  As Boolean  '[FLAG FOR FIRST ROW DATES]
    Dim dateStartDate   As Date     '[DATE FOR DISPLAY BESIDE SHORT DAY FORMAT]
    
    '[PRINT THE GIVEN ROW FROM THE ROSTER GRID]
    intLineCount = intLineCount + 1

    '[SET ROSTER GRID ROW TO CORRESPOND]
    GridPrint.Row = intRowToPrint
    GridPrint.Col = 0
    sinLinePos = 0
    flagDateNeeded = False
    
    '[DETERMINE IF DATES ARE NEEDED]
    If mdiMain.ActiveForm.Name = "frmRoster" And intRowToPrint = 0 Then
        flagDateNeeded = True
        '[CHECK TO SEE IF THE MASK DATE ON THE ROSTER FORM IS A VALID DATE]
        If IsDate(frmControl.MaskDate) Then
            dateStartDate = CDate(frmControl.MaskDate)
        Else
            dateStartDate = Date
        End If
    End If
    
    '[CYCLE THROUGH COLUMNS AND DETERMINE BOX HEIGHT]
    For intColumn = intStartCol To (GridPrint.Cols - 1)
        '[SET COLUMN]
        GridPrint.Col = intColumn
        '[ACCULMULATE GRID COLUMN WIDTH]
        sinGridWidth = sinGridWidth + GridPrint.ColWidth(intColumn)
        '[COUNT NUMBERS OF NAMES IN CELL]
        intStart = 1
        intCounter = 0
        Do While intStart > 0
            '[SEARCH FOR ENTER KEY CHARACTER (CR)]
            intStart = InStr(intStart + 1, GridPrint.Text, Chr$(vbKeyReturn))
            If intStart > 0 Then intCounter = intCounter + 1
        Loop
        
        '[IF NO CR'S FOUND, MAKE BOX 1 LINE ELSE ALLOW SPACE AT END OF BOX]
        If intCounter = 0 Then
            If mdiMain.ActiveForm.Name = "frmRoster" Then intCounter = 2 Else intCounter = 1
        Else
            If intRowToPrint = 0 Then intCounter = intCounter + 1 Else intCounter = intCounter + 2
        End If
        
        Rem If flagDateNeeded = True Then intCounter = intCounter + 1
        
        '[CHECK BOX SIZE TO GET MAXIMUM]
        If sinBoxHeight < intCounter * sngLineHeight Then
            sinBoxHeight = (intCounter * sngLineHeight)
            intBoxRows = intCounter
        End If
    Next intColumn
    
    '[CYCLE THROUGH COLUMNS AND PRINT]
    For intColumn = intStartCol To (GridPrint.Cols - 1)

        '[DETERMINE PERCENTAGE WIDTH OF GRID FOR EACH COLUMN (90% of PRINTER PAGE WIDTH)]
        sinLinePos = sinLinePos + sngColumnWidth            '[ACCUMULATE PREVIOUS POSITIONS ON THE LINE]
        sngColumnWidth = (sngPageWidth * (GridPrint.ColWidth(intColumn) / sinGridWidth))
        
        '[DEBUG]
        ' Debug.Print intColumn, sngColumnWidth, sngPageWidth, sinLinePos, sngColumnWidth + sinLinePos, Printer.Width, GridPrint.ColWidth(intColumn), sinGridWidth
        
        '[BOUNDING BOX]
        '[SET TO BLACK ON WHITE PRINT FOR STANDARD OR WHITE ON BLACK FOR HEADINGS]
        If intRowToPrint < 0 Then '[CHANGE BACK TO =0 IF WHITE ON BLACK TITLES ARE NEEDED]
            Printer.FillStyle = 0
            Printer.FillColor = vbBlack
            Printer.Line ((sinLinePos), (sngLineHeight * intLineCount))-Step(sngColumnWidth, sngLineHeight), vbBlack, BF
            Printer.FillStyle = 1
            Printer.ForeColor = vbWhite
        Else
            Printer.FillStyle = 1
            Printer.FillColor = vbWhite
            
            Select Case intBoxState
            Case 0  '[PRINT ALL BOXES]
                Printer.Line ((sinLinePos), (sngLineHeight * intLineCount))-Step(sngColumnWidth, sinBoxHeight), vbBlack, B
            Case 1  '[PRINT SOME BOXES]
                If intRowToPrint = 0 Then   '[BOX FIRST ROW]
                    Printer.Line ((sinLinePos), (sngLineHeight * intLineCount))-Step(sngColumnWidth, sinBoxHeight), vbBlack, B
                Else
                    Printer.Line ((sinLinePos), (sngLineHeight * intLineCount))-Step(0, sinBoxHeight), vbBlack, B
                End If
            Case 2  '[PRINT NO BOXES]
            
            End Select

            Printer.ForeColor = vbBlack
        End If
        
        '[SET COLUMN]
        GridPrint.Col = intColumn
        '[ALLOCATE CELL TEXT]
        If flagDateNeeded And intColumn > 1 Then
            '[ADD CR AND DATE IF REQUIRED]
            strCellText = "[" & (dateStartDate + (intColumn - 2)) & "]" & Chr$(13) & GridPrint.Text
        Else
            strCellText = GridPrint.Text
        End If
        
        intStart = 1
        intCounter = 0
        
        Do While intStart > 0
            '[CYCLE THROUGH ALL NAMES IN CELL]
            intStart = InStr(intStart, strCellText, Chr$(13))
            If intStart > 0 Then
                intCounter = intCounter + 1
                '[EXTRACT NAME FROM LEFT OF CELL TEXT]
                strCellName = Left(strCellText, intStart - 1)
                '[TRIM CELL TEXT BY EXTRACTED STRING LENGTH]
                strCellText = Mid(strCellText, (intStart + 1))
                intStart = 1
                '[CHECK STRING LENGTH TO MAKE SURE IT FITS INTO THE CELL.  IF NOT, TRUNCATE BY ONE CHARACTER AND TRY AGAIN]
                Do While Printer.TextWidth(strCellName) > (sngColumnWidth * 0.95)
                    strCellName = Left$(strCellName, Len(strCellName) - 1)
                Loop
                '[PRINT CELL TEXT]
                Printer.CurrentX = sinLinePos + (sngColumnWidth * 0.05): Printer.CurrentY = (sngLineHeight * (intLineCount + (intCounter - 1))) + (sngLineHeight * 0.05): Printer.Print strCellName
                Printer.FillStyle = 1
                Printer.ForeColor = vbBlack
            ElseIf Len(Trim(strCellText)) > 0 Then
                intCounter = intCounter + 1
                '[ONLY ONE NAME IN CELL]
                '[CHECK STRING LENGTH TO MAKE SURE IT FITS INTO THE CELL.  IF NOT, TRUNCATE BY ONE CHARACTER AND TRY AGAIN]
                Do While Printer.TextWidth(strCellText) > (sngColumnWidth * 0.95)
                    strCellText = Left$(strCellText, Len(strCellText) - 1)
                Loop
                '[PRINT CELL TEXT]
                Printer.CurrentX = sinLinePos + (sngColumnWidth * 0.05): Printer.CurrentY = (sngLineHeight * (intLineCount + (intCounter - 1))) + (sngLineHeight * 0.05): Printer.Print strCellText
                Printer.FillStyle = 1
                Printer.ForeColor = vbBlack
            End If
        Loop

    Next intColumn
    
    '[RIGHT HAND LINE]
    If intBoxState = 1 Then Printer.Line ((sinLinePos + sngColumnWidth), (sngLineHeight * intLineCount))-Step(0, sinBoxHeight), vbBlack, B
    
    '[INCREMENT COUNTER FOR EACH BOX LINE DRAWN]
    intLineCount = intLineCount + (intBoxRows - 1)
    
End Sub


Sub RosterInfo(strDisplayText, sinForeColor)

    '[SET CAPTION ON INFO LABEL TO PASSED TEXT AND COLOR TO PASSED COLOR]
    If sinForeColor <= 0 Then sinForeColor = frmRoster.ForeColor
    frmRoster.labelInfo.ForeColor = sinForeColor
    frmRoster.labelInfo.Caption = strDisplayText
    frmRoster.labelInfo.Refresh

End Sub
Sub ReportInfo(strDisplayText, sinForeColor)

    '[SET CAPTION ON INFO LABEL TO PASSED TEXT AND COLOR TO PASSED COLOR]
    If sinForeColor <= 0 Then sinForeColor = vbBlack
    frmMsg.labelInfo.ForeColor = sinForeColor
    frmMsg.labelInfo.Caption = strDisplayText
    frmMsg.labelInfo.Refresh

End Sub

Sub BuildRosterDynaset(intClass)

    '[REBUILD ROSTER DYNASET BASED UPON THE CLASS PASSED]
    Dim SQLStmt         As String
    
    '[SELECT APPROPRIATE RECORDS]
    SQLStmt = "SELECT * FROM Roster WHERE Class = " & Str(intClass) & " ORDER BY ID, CLASS, SHIFTSTART"
    Set DsRoster = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
    
End Sub

Sub FillRosterGrid()

    '[PLACE LIST ITEMS IN ROSTER DYNASET]
    Dim intCol          As Integer
    Dim intRow          As Integer
    Dim intCounter      As Integer
        
    '[ROSTER LABEL INFO]
    Call RosterInfo("Loading Roster", 0)
        
    '[CLEAR CURRENT GRID AND PREPARE]
    frmRoster.GridRoster.Rows = 2
    
    '[NEED TO CLEAR LAST ROW OF GRID]
    frmRoster.GridRoster.Row = 1
    For intCol = 0 To (frmRoster.GridRoster.Cols - 1)
        frmRoster.GridRoster.Col = intCol
        frmRoster.GridRoster.Text = ""
    Next intCol
    
    intCounter = 1
    
    If DsRoster.RecordCount > 0 Then
    
        '[CYLCE THROUGH DYNASET AND PLACE IN GRID]
        DsRoster.MoveFirst
        
        Do While Not DsRoster.EOF
            frmRoster.GridRoster.AddItem Format(DsRoster("ShiftStart"), "Medium Time") & Chr$(vbKeyTab) & Format(DsRoster("ShiftEnd"), "Medium Time") & Chr$(vbKeyTab) & DsRoster("Day_1") & Chr$(vbKeyTab) & DsRoster("Day_2") & Chr$(vbKeyTab) & DsRoster("Day_3") & Chr$(vbKeyTab) & DsRoster("Day_4") & Chr$(vbKeyTab) & DsRoster("Day_5") & Chr$(vbKeyTab) & DsRoster("Day_6") & Chr$(vbKeyTab) & DsRoster("Day_7"), intCounter
            '[OLD TIME FORMAT] -> "hh:mm AMPM"
            frmRoster.GridRoster.Row = intCounter
            ProgressBar ((intCounter / DsRoster.RecordCount) * 100)
            intCounter = intCounter + 1
            '[CYCLE THROUGH COLUMNS AND SIZE]
            For intCol = 2 To 8
                frmRoster.GridRoster.Col = intCol
                ResizeRosterCell (frmRoster.GridRoster.Text)
            Next intCol
            DsRoster.MoveNext
        Loop
    
    End If
    
    '[REMOVE LAST ROW]
    If frmRoster.GridRoster.Rows > 2 Then frmRoster.GridRoster.Rows = frmRoster.GridRoster.Rows - 1
    '[CLEAR PROGRESS BAR]
    ProgressBar (0)
    
    '[ROSTER LABEL INFO]
    Call RosterInfo("", 0)
    
End Sub

Sub ProgressBar(intPercent)

    '[CATCH EXTREME VALUES]
    If intPercent < 0 Then intPercent = 0
    If intPercent > 100 Then intPercent = 100
    '[SHOW PROGRESS IN MENU BAR]
    frmRoster.GaugeProgress.Value = intPercent
    frmRoster.GaugeProgress.Refresh
    

End Sub

Sub ReportProgressBar(intPercent)

    '[CATCH EXTREME VALUES]
    If intPercent < 0 Then intPercent = 0
    If intPercent > 100 Then intPercent = 100
    '[SHOW PROGRESS IN MENU BAR]
    frmMsg.GaugeProgress.Value = intPercent
    frmMsg.GaugeProgress.Refresh

End Sub
Sub PutNameInCell(strFullString)

    Dim sinColWidth         As Single
    Dim sinRowHeight        As Single
    Dim intEnterCount       As Single
    Dim intStartPos         As Single
    
    intStartPos = 1
    intEnterCount = 0
    
    '[CHECK FOR PLACEMENTS INTO TITLE ROWS/COLS]
    If frmRoster.GridRoster.Col <= 1 Or frmRoster.GridRoster.Row = 0 Then Exit Sub
    
    '[ROUTINE TO PLACE THE PASSED NAME INTO THE ROSTER CELL]
    '[CHECKING FOR OTHER NAMES AND PREVIOUS INSTANCE OF NAME]
    If Len(Trim(frmRoster.GridRoster.Text)) = 0 Then
        '[CELL IS EMPTY - PLACE NAME]
        frmRoster.GridRoster.Text = Trim(strFullString)
    Else
        '[CELL IS NOT EMPTY - CHECK FOR NAME]
        If InStr(frmRoster.GridRoster.Text, strFullString) > 0 Then
            '[NAME IS ALREADY IN CELL, EXIT ROUTINE]
            Exit Sub
        Else
            '[NAME NOT IN CELL, ADD AT END OF CELL]
            frmRoster.GridRoster.Text = Trim(frmRoster.GridRoster.Text) & Chr$(vbKeyReturn) & Trim(strFullString)
        End If
    End If

    ResizeRosterCell (strFullString)

End Sub
Sub RemoveFromRoster(strFullString)
    
    '[REMOVE THE PASSED STAFF NAME FROM ALL SELECTED CELLS IN THE ROSTER GRID]
    Dim intCounter          As Integer
    Dim strLastName         As String
    Dim strFirstName        As String
    Dim intDelimiter        As Integer
    Dim intCol              As Integer
    Dim intRow              As Integer
    
    '[Delimit list item, break into Surname, FirstName]
    intDelimiter = InStr(strFullString, ",")
    strLastName = Trim(Left(strFullString, intDelimiter - 1))
    strFirstName = Trim(Mid(strFullString, intDelimiter + 1))

    '[DETERMINE WHETHER IT IS A MULTISELECT OR SINGLE CELL FILL]
    If frmRoster.GridRoster.SelStartCol = -1 Or frmRoster.GridRoster.SelEndCol = -1 Then
        '[SINGLE CELL REMOVE]
        '[CALL ROUTINE TO REMOVE NAME FROM CELL]
        RemoveNameFromCell (strFullString)
    Else
        '[MULTI CELL REMOVE]
        For intCol = frmRoster.GridRoster.SelStartCol To frmRoster.GridRoster.SelEndCol
            frmRoster.GridRoster.Col = intCol
            For intRow = frmRoster.GridRoster.SelStartRow To frmRoster.GridRoster.SelEndRow
                frmRoster.GridRoster.Row = intRow
                '[CALL ROUTINE TO REMOVE NAME FROM CELL]
                RemoveNameFromCell (strFullString)
            Next intRow
        Next intCol
    End If

End Sub

Sub RemoveNameFromCell(strFullString)
    
    Dim sinColWidth         As Single
    Dim sinRowHeight        As Single
    Dim intEnterCount       As Single
    Dim intStartPos         As Single
    Dim strLeft             As String
    Dim strRight            As String
    Dim intLeftPos          As Integer
    Dim intRightPos         As Integer
    Dim intLength           As Integer
    
    intStartPos = 1
    intEnterCount = 0
    
    '[ROUTINE TO REMOVE THE PASSED NAME FROM THE ROSTER CELL]
    '[CHECKING FOR OTHER NAMES AND PREVIOUS INSTANCE OF NAME]
    If frmRoster.GridRoster.Text = "" Then
        '[CELL IS EMPTY - EXIT ROUTINE]
        Exit Sub
    Else
        '[CELL IS NOT EMPTY - CHECK FOR NAME]
        If InStr(frmRoster.GridRoster.Text, strFullString) > 0 Then
            '[NAME IS IN CELL, REMOVE FROM CELL]
            '[FIND STRING LEFT OF PASSED NAME]
            intLeftPos = InStr(frmRoster.GridRoster.Text, strFullString)
            If intLeftPos = 1 Then
                strLeft = ""
            Else
                strLeft = Left(frmRoster.GridRoster.Text, intLeftPos - 2)
            End If
            '[FIND STRING RIGHT OF PASSED NAME]
            strRight = Trim(Mid$(frmRoster.GridRoster.Text, Len(strLeft) + Len(strFullString) + 2))
            '[APPLY NEW CONCATENATED STRING TO CELL]
            frmRoster.GridRoster.Text = strLeft & strRight
        Else
            '[NAME NOT IN CELL, EXIT ROUTINE]
            Exit Sub
        End If
    End If

    '[RESIZE ROW AND COL TO SUIT]
    sinRowHeight = frmRoster.GridRoster.RowHeight(frmRoster.GridRoster.Row)
    sinColWidth = frmRoster.GridRoster.ColWidth(frmRoster.GridRoster.Col)
    
    If sinColWidth < (frmRoster.TextWidth(strFullString) * 1.25) Then frmRoster.GridRoster.ColWidth(frmRoster.GridRoster.Col) = (frmRoster.TextWidth(strFullString) * 1.25)
            
    '[LOOK FOR ENTER CHARACTERS IN THE CURRENT CELL AND ADJUST SIZE]
    Do While InStr(intStartPos, frmRoster.GridRoster.Text, Chr$(vbKeyReturn)) > 0
        intEnterCount = intEnterCount + 1
        intStartPos = InStr(intStartPos, frmRoster.GridRoster.Text, Chr$(vbKeyReturn)) + 1
    Loop

    If sinRowHeight < (frmRoster.TextHeight("A") * (intEnterCount + 1) * 1.25) Then frmRoster.GridRoster.RowHeight(frmRoster.GridRoster.Row) = (frmRoster.TextHeight("A") * (intEnterCount + 1) * 1.25)

End Sub

Sub ResizeRosterCell(strFullString)
    
    '[DECLARATIONS]
    Dim sinRowHeight        As Single
    Dim sinColWidth         As Single
    Dim intEnterCount       As Integer
    Dim intStartPos         As Integer
    
    '[SET START POSITION]
    intStartPos = 1
    intEnterCount = 0
    
    '[RETURN IF CELL IS EMPTY]
    If Len(frmRoster.GridRoster.Text) <= 0 Then Exit Sub
    
    '[RESIZE ROW AND COL TO SUIT]
    sinRowHeight = frmRoster.GridRoster.RowHeight(frmRoster.GridRoster.Row)
    sinColWidth = frmRoster.GridRoster.ColWidth(frmRoster.GridRoster.Col)
    
    If sinColWidth < (frmRoster.TextWidth(strFullString) * 1.1) Then frmRoster.GridRoster.ColWidth(frmRoster.GridRoster.Col) = (frmRoster.TextWidth(strFullString) * 1.1)
            
    '[LOOK FOR ENTER CHARACTERS IN THE CURRENT CELL AND ADJUST SIZE]
    Do While InStr(intStartPos, frmRoster.GridRoster.Text, Chr$(vbKeyReturn)) > 0
        intEnterCount = intEnterCount + 1
        intStartPos = InStr(intStartPos, frmRoster.GridRoster.Text, Chr$(vbKeyReturn)) + 1
    Loop

    If sinRowHeight < (frmRoster.TextHeight("A") * (intEnterCount + 1) * 1.1) Then frmRoster.GridRoster.RowHeight(frmRoster.GridRoster.Row) = (frmRoster.TextHeight("A") * (intEnterCount + 1) * 1.1)

End Sub

Sub SaveRosterGrid()

    '[PLACE LIST ITEMS IN ROSTER DYNASET]
    Dim intCol          As Integer
    Dim intRow          As Integer
    Dim intCounter      As Integer
    Dim strDayKey       As String
        
    '[ROSTER LABEL INFO]
    Call RosterInfo("Saving Roster", 0)
        
    '[CLEAR CURRENT DYNASET AND PREPARE]
    Do While DsRoster.RecordCount > 0
        DsRoster.MoveFirst
        DsRoster.Delete
    Loop
    intCounter = 1
    
    For intRow = 1 To (frmRoster.GridRoster.Rows - 1)
        frmRoster.GridRoster.Row = intRow
        frmRoster.GridRoster.Col = 0
        '[ADD NEW ITEM TO DYNASET]
        DsRoster.AddNew
            DsRoster("Class") = intRosterClass
            DsRoster("Time") = frmRoster.GridRoster.Text
            DsRoster("ShiftStart") = frmRoster.GridRoster.Text
            frmRoster.GridRoster.Col = 1
            DsRoster("ShiftEnd") = frmRoster.GridRoster.Text
            DsRoster("Row") = intRow
        
        For intCol = 2 To (frmRoster.GridRoster.Cols - 1)
            frmRoster.GridRoster.Col = intCol
            '[CREATE DAY KEY]
            strDayKey = "Day_" & Trim(Str(intCol - 1))
            '[TRIM AND REMOVE LEADING CR'S FROM TEXT]
            frmRoster.GridRoster.Text = Trim(frmRoster.GridRoster.Text)
            If Left$(frmRoster.GridRoster.Text, 1) = Chr$(vbKeyReturn) Then frmRoster.GridRoster.Text = Mid(frmRoster.GridRoster.Text, 2)
            DsRoster(strDayKey) = (frmRoster.GridRoster.Text & " ")
        Next intCol
        
        '[UPDATE DYNASET ITEM]
        DsRoster.Update
        '[PROGRESS BAR]
        ProgressBar (intRow / (frmRoster.GridRoster.Rows - 1) * 100)
        
    Next intRow
    '[ZERO PROGRESS BAR]
    ProgressBar (0)
    
    '[ROSTER LABEL INFO]
    Call RosterInfo("", 0)

End Sub

Sub StaffReport()

    '[THIS IS THE STAFF REPORT SUBROUTINE. IT WILL CREATE A TEMP DYNASET        ]
    '[TO CONTAIN ALL OF THE ROSTER RECORDS AND CYCLE THROUGH USING SQL          ]
    '[STATEMENTS (HOPEFULLY).                                                   ]
    
    '[BUILD ROSTER DYNASET WITH ALL RECORDS]
    Dim SQLStmt         As String
    Dim strBookmark     As String
    Dim strFullname     As String
    Dim strDayKey       As String
    Dim intCounter      As Integer
    Dim intDayCount     As Integer
    Dim strClassDesc    As String
    Dim sinMinsWorked   As Single
    Dim sinIncrement    As Single
    Dim intInterval     As Integer
    Dim intRosterCount  As Integer
    Dim dateStart       As Date
    Dim dateEnd         As Date
    Dim flagStart       As Boolean
    Dim strStaffId      As String
    Dim strName         As String
    Dim strRoster       As String
    Dim sinMinutes      As Single
    Dim sinTotal        As Single
    Dim sinTotalAmount  As Single
    Dim sinTotalMinutes As Single
    
    '[SET UP ARRAY FOR HOLDING STAFF WAGE DATA]
    Dim arrayStaff(10) As StaffType
    Dim arrayRoster(10) As StaffType
    
    '[STORE CURRENT STAFF LOCATION]
    strBookmark = DsStaff.Bookmark
    
    '[SELECT APPROPRIATE RECORDS]
    SQLStmt = "SELECT * FROM Roster ORDER BY CLASS, SHIFTSTART"
    Set DsReport = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
            
    '[CHECK CURRENT DYNASET AND PREPARE]
    If DsRoster.EOF And DsRoster.BOF Then
        DsReport.Close
        Exit Sub
    End If
    
    '[*STAFF LOOP**************************************************************]
    '[MOVE TO FIRST STAFF RECORD]
    DsStaff.MoveFirst
    '[CYCLE THROUGH STAFF LIST]
    Do While Not DsStaff.EOF
        
        '[RESET MINUTES WORKED]
        Erase arrayStaff
        '[DETERMINE FULL NAME]
        strFullname = DsStaff("LastName") & ", " & DsStaff("FirstName")
        '[SHOW PROGRESS REPORT]
        Call ReportInfo(strFullname, 0)
        '[PROGRESS BAR]
        Call ReportProgressBar(((DsStaff.AbsolutePosition + 1) / DsStaff.RecordCount) * 100)
                        
        '[=ROSTER LOOP=========================================================]
        '[MOVE TO FIRST RECORD]
        DsReport.MoveFirst
        '[CYCLE THROUGH ROSTER]
        Do While Not DsReport.EOF
            
            '[CHECK TO SEE IF THE ROSTER IS ACTIVE]
            DsClass.AbsolutePosition = (DsReport("Class") - 1)
            
            If DsClass("Active") = vbChecked And Not IsNull(DsReport("ShiftStart")) And Not IsNull(DsReport("ShiftEnd")) Then
                For intDayCount = 1 To 7
                    '[-=-NOW CHECK EACH DAY TO SEE IF STAFF MEMBER IS INCLUDED IN ANY DAY-=-]
                    strDayKey = "Day_" & Trim(Str(intDayCount))
                    
                    If InStr(DsReport(strDayKey), strFullname) > 0 Then
                        '[RESET INCREMENT - ALLOWS FOR ROSTERS WITH DIFFERING INCREMENTS]
                        dateStart = DsReport("ShiftStart")
                        dateEnd = DsReport("ShiftEnd")
                        '[ALLOW FOR NEXT DAY TIMES]
                        If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
                        '[CALCULATE INCREMENT]
                        sinIncrement = (dateEnd - dateStart) * (24 * 60)
                        '[ADD MINUTES TO STAFF ARRAY]
                        arrayStaff(DsReport("Class")).Minutes = arrayStaff(DsReport("Class")).Minutes + sinIncrement
                    End If
                    
                Next intDayCount
            End If
            
            '[MOVE TO NEXT ROSTER RECORD]
            DsReport.MoveNext
            
        Loop
        '[=====================================================================]
        '[DISPLAY STAFF RECORD HERE - PLACE DETAILS IN GRIDREPORT ON FRMREPORT ]
        flagStart = False
        sinMinsWorked = 0
        sinTotalMinutes = 0
        sinTotalAmount = 0
        
        For intCounter = 1 To 10
            '[CHECK TO SEE IF THERE ARE HOURS ALLOCATED FOR THIS RECORD]
            If arrayStaff(intCounter).Minutes > 0 Then
                '[CHECK FOR NAME PRINTED]
                If flagStart = False Then
                    '[ASSIGN STAFF ID AND NAMES]
                    strStaffId = DsStaff("StaffID")
                    strName = strFullname
                    flagStart = True
                Else
                    strStaffId = ""
                    strName = ""
                End If
                '[ASSIGN OTHER VALUES - ROSTER NAME, HOURS AND TOTAL]
                sinMinutes = arrayStaff(intCounter).Minutes / 60
                If IsNull(DsStaff("HourRate")) Then sinTotal = 0 Else sinTotal = DsStaff("HourRate") * sinMinutes
                '[SET ROSTER NAME]
                DsClass.AbsolutePosition = (intCounter - 1)
                strRoster = DsClass("Description")
                '[ADD MINUTES TO TOTAL]
                sinTotalMinutes = sinTotalMinutes + sinMinutes
                sinTotalAmount = sinTotalAmount + sinTotal
                '[ADD MINUTES TO TOTAL ROSTER ARRAY]
                arrayRoster(intCounter).Minutes = arrayRoster(intCounter).Minutes + sinMinutes
                '[ADD CURRENCY AMOUNT TO TOTAL ROSTER ARRAY]
                arrayRoster(intCounter).Amount = arrayRoster(intCounter).Amount + sinTotal
                '[ADD LINE TO REPORT GRID]
                frmReport.GridReport.AddItem strStaffId & Chr$(vbKeyTab) & strName & Chr$(vbKeyTab) & strRoster & Chr$(vbKeyTab) & Format(sinMinutes, ("##0.00")) & Chr$(vbKeyTab) & Format(sinTotal, ("Currency")), (frmReport.GridReport.Rows - 1)
            End If
        Next intCounter
        
        If sinTotalMinutes > 0 Then
            '[ADD TOTALS]
            frmReport.GridReport.AddItem "" & Chr$(vbKeyTab) & "" & Chr$(vbKeyTab) & "TOTALS" & Chr$(vbKeyTab) & Format(sinTotalMinutes, ("##0.00")) & Chr$(vbKeyTab) & Format(sinTotalAmount, ("Currency")), (frmReport.GridReport.Rows - 1)
            '[AND BLANK ROW]
            frmReport.GridReport.AddItem "", (frmReport.GridReport.Rows - 1)
        End If
        '[=====================================================================]
        '[MOVE TO NEXT STAFF RECORD]
        DsStaff.MoveNext
    Loop
    '[*************************************************************************]
    '[PLACE STAFF ROSTER TOTALS]
    sinTotalMinutes = 0
    sinTotalAmount = 0

    For intCounter = 1 To 10
        '[CHECK TO SEE IF THERE ARE HOURS ALLOCATED FOR THIS RECORD]
        If arrayRoster(intCounter).Minutes > 0 Then
            strStaffId = ""
            strName = ""
            '[ASSIGN OTHER VALUES - ROSTER NAME, HOURS AND TOTAL]
            sinMinutes = arrayRoster(intCounter).Minutes
            sinTotal = arrayRoster(intCounter).Amount
            '[SET ROSTER NAME]
            DsClass.AbsolutePosition = (intCounter - 1)
            strRoster = DsClass("Description")
            '[ADD MINUTES TO TOTAL]
            sinTotalMinutes = sinTotalMinutes + sinMinutes
            sinTotalAmount = sinTotalAmount + sinTotal
            '[ADD LINE TO REPORT GRID]
            frmReport.GridReport.AddItem strStaffId & Chr$(vbKeyTab) & strName & Chr$(vbKeyTab) & strRoster & Chr$(vbKeyTab) & Format(sinMinutes, ("##0.00")) & Chr$(vbKeyTab) & Format(sinTotal, ("Currency")), (frmReport.GridReport.Rows - 1)
        End If
    Next intCounter
    '[*************************************************************************]
    If sinTotalMinutes > 0 Then
        '[ADD TOTALS]
        frmReport.GridReport.AddItem "" & Chr$(vbKeyTab) & "" & Chr$(vbKeyTab) & "TOTALS" & Chr$(vbKeyTab) & Format(sinTotalMinutes, ("##0.00")) & Chr$(vbKeyTab) & Format(sinTotalAmount, ("Currency")), (frmReport.GridReport.Rows - 1)
        '[AND BLANK ROW]
        frmReport.GridReport.AddItem "", (frmReport.GridReport.Rows - 1)
    End If
        
    
    '[CLEAR REPORT INFO]
    Call ReportInfo("", 0)
    '[CLEAR PROGRESS BAR]
    Call ReportProgressBar(0)
    
    '[RETURN TO STAFF BOOKMARK]
    DsStaff.Bookmark = strBookmark
    
    '[CLOSE TEMPORARY DYNASET]
    DsReport.Close

    '[MOVE TO FIRST ROW IN REPORT]
    frmReport.GridReport.Row = 1
    frmReport.GridReport.Col = 0
    
    '[CALL SUBROUTINE TO RESIZE FORM]
    Call resReportForm


End Sub

Sub StatusBar(strMessage)

    '[PLACE PASSED MESSAGE ON THE STATUS BAR]
    If mdiMain.panelStatusBar.Caption <> strMessage Then mdiMain.panelStatusBar.Caption = strMessage

End Sub

Sub Terminate()

    '[CHECK FOR PRESENCE OF SAVE BUTTONS]
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title
    
    '[CHECK FOR CHANGED DATA AND NOTIFY]
    If frmRoster.cmdSave.Visible = True Then
        '[ROSTER IS UNSAVED, POPUP YES/NO DIALOG]
        Msg = "You have made changes to this roster (" & frmRoster.ComboClass.Text & ") but have not saved these changes." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "If you choose not to save now, any changes you have made since your last save will be lost." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to save these changes before you exit ?"
        Style = vbYesNoCancel ' Define buttons.
        Title = "Roster Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            SaveRosterGrid
        ElseIf Response = vbCancel Then
            Exit Sub
        End If
    End If
    
    '[CHECK TO SEE IF WE ARE MOVING FROM AN UNSAVED RECORD]
    If frmStaff.cmdSave.Visible = True Then
        '[RECORD IS UNSAVED, POPUP YES/NO DIALOG]
        Msg = "You have made changes to this staff record (" & DsStaff("Lastname") & ", " & DsStaff("FirstName") & ") but have not saved these changes." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "If you choose not to save now, any changes you have made since your last save will be lost." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to save these changes before you exit ?"
        Style = vbYesNoCancel ' Define buttons.
        Title = "Staff Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            '[COMMIT RECORD CHANGES TO THE DYNASET]
            SaveStaffDetails
        ElseIf Response = vbCancel Then
            Exit Sub
        End If
    End If


    '[place closing/termination statements here]
    Dim intStaffState       As Integer
    Dim intRosterState      As Integer
    Dim intControlState     As Integer
    Dim intMainState        As Integer
    
    '[DEFAULT 1=NORMAL]
    intStaffState = 1
    intRosterState = 1
    intControlState = 1
    intMainState = 1
    
    If frmStaff.Visible = False Then intStaffState = 0
    If frmRoster.Visible = False Then intRosterState = 0
    If frmControl.Visible = False Then intControlState = 0
    If mdiMain.Visible = False Then intMainState = 0
            
    If frmStaff.WindowState = 2 Then intStaffState = 2
    If frmRoster.WindowState = 2 Then intRosterState = 2
    If frmControl.WindowState = 2 Then intControlState = 2
    If mdiMain.WindowState = 2 Then intMainState = 2
    
    '[SAVE WINDOW STATES HERE]
    DsDefault.Edit
    
        DsDefault("StaffState") = intStaffState
        DsDefault("RosterState") = intRosterState
        DsDefault("ControlState") = intControlState
        DsDefault("MainState") = intMainState
        DsDefault("DeleteConfirm") = frmControl.CheckDelete.Value
    
        '[SAVE ROSTER GRID FONTS]
        DsDefault("RosterFontName") = frmRoster.GridRoster.Font.Name
        DsDefault("RosterFontBold") = frmRoster.GridRoster.Font.Bold
        DsDefault("RosterFontItalic") = frmRoster.GridRoster.Font.Italic
        DsDefault("RosterFontSize") = frmRoster.GridRoster.Font.Size
        
        '[SAVE LAST ROSTER ID]
        DsDefault("RosterID") = frmRoster.ComboClass.ListIndex
        
        '[SET TOOLBAR STATE]
        If mdiMain.PanelToolBar.Visible = True Then
            DsDefault("ToolBarState") = 1
        Else
            DsDefault("ToolBarState") = 0
        End If
        
        '[SET STATUSBAR STATE]
        If mdiMain.panelStatusBar.Visible = True Then
            DsDefault("StatusBarState") = 1
        Else
            DsDefault("StatusBarState") = 0
        End If
        
        '[SAVE LOCKED COLUMNS]
        DsDefault("RosterLocked") = frmRoster.GridRoster.FixedCols
        
    DsDefault.Update

    Close
    End


End Sub

Sub TransferToRoster(strFullString)

    '[TRANSFER THE PASSED STAFF NAME TO ALL SELECTED CELLS IN THE ROSTER GRID]
    Dim intCounter          As Integer
    Dim strLastName         As String
    Dim strFirstName        As String
    Dim intDelimiter        As Integer
    Dim intCol              As Integer
    Dim intRow              As Integer
    Dim Msg As String
    Dim Style
    Dim Response
    Dim Title

    '[Delimit list item, break into Surname, FirstName]
    intDelimiter = InStr(strFullString, ",")
    strLastName = Trim(Left(strFullString, intDelimiter - 1))
    strFirstName = Trim(Mid(strFullString, intDelimiter + 1))

    '[DETERMINE WHETHER IT IS A MULTISELECT OR SINGLE CELL FILL]
    If frmRoster.GridRoster.SelStartCol = -1 Or frmRoster.GridRoster.SelEndCol = -1 Then
        '[SINGLE CELL FILL]
        '[CALL ROUTINE TO PLACE NAME IN CELL]
        If frmRoster.GridRoster.Col <= 1 Then Exit Sub
        If CheckStaffDay(strFullString, frmRoster.GridRoster.Col) Then
            PutNameInCell (strFullString)
        Else
            '[SHOW MESSAGE]
            Msg = "This staff member (" & strFullString & ") is marked as not being available on " & ArrayWeek(DayOfWeek(frmRoster.GridRoster.Col - 1)).LongDay & "s." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "If you require this staff member to be available on this day, you should click the appropriate check box on the staff details form."
            Style = vbOKOnly                    ' Define buttons.
            Title = "Staff Member Not Available"     ' Define title.
            Response = gsrMsg(Msg, Style, Title)
        End If
    Else
        '[MULTI CELL FILL]
        For intCol = frmRoster.GridRoster.SelStartCol To frmRoster.GridRoster.SelEndCol
            If intCol <= 1 Then Exit For
            frmRoster.GridRoster.Col = intCol
            For intRow = frmRoster.GridRoster.SelStartRow To frmRoster.GridRoster.SelEndRow
                frmRoster.GridRoster.Row = intRow
                '[CHECK STAFF MEMBER IS AVAILABLE FOR THIS DAY]
                If CheckStaffDay(strFullString, intCol) Then
                    '[CALL ROUTINE TO PLACE NAME IN CELL]
                    PutNameInCell (strFullString)
                Else
                    '[SHOW MESSAGE]
                    Msg = "This staff member (" & strFullString & ") is marked as not being available on " & ArrayWeek(DayOfWeek(intCol - 1)).LongDay & "s." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "If you require this staff member to be available on this day, you should click the appropriate check box on the staff details form."
                    Style = vbOKOnly                    ' Define buttons.
                    Title = "Staff Member Not Available"     ' Define title.
                    Response = gsrMsg(Msg, Style, Title)
                    '[EXIT THE ROUTINE]
                    Exit For
                End If
            Next intRow
        Next intCol
    End If
    

End Sub
Sub AddNewStaff()

    Dim strDisplayname      As String
    Dim intCounter          As Integer
    Dim SQLStmt             As String
    
    '[ADD A NEW STAFF MEMBER TO THE LIST]
    DsStaff.AddNew
        DsStaff("StaffID") = "*999999"
        DsStaff("LastName") = "*LastName"
        DsStaff("FirstName") = "*FirstName"
        DsStaff("Day_1") = 1
        DsStaff("Day_2") = 1
        DsStaff("Day_3") = 1
        DsStaff("Day_4") = 1
        DsStaff("Day_5") = 1
        DsStaff("Day_6") = 1
        DsStaff("Day_7") = 1
    DsStaff.Update
    
    '[MOVE TO FIRST RECORD IF NO RECORD LOCATED]
    DsStaff.Bookmark = DsStaff.LastModified
    
    If DsStaff.EOF Or DsStaff.BOF Then DsStaff.MoveFirst
    
    strDisplayname = DsStaff("LastName") & ", " & DsStaff("FirstName")
    
    '[REFILL STAFF LIST FOR ORDER]
    FillStaffList
                    
    '[RELOCATE STAFF NAME]
    For intCounter = 0 To (frmStaff.ListStaff.ListCount - 1)
        If frmStaff.ListStaff.List(intCounter) = strDisplayname Then frmStaff.ListStaff.ListIndex = intCounter
    Next intCounter

End Sub

Sub DropStaffOnRoster(X, Y)
    
    '[DROP THE SELECTED STAFF MEMBER INTO THE SELECTED GRID CELL]
    '[FIRST STEP - LOCATE X,Y POS IN GRID]
    Dim sinRowHeight        As Single
    Dim sinColWidth         As Single
    Dim intRow              As Integer
    Dim intCol              As Integer
    Dim intColCounter       As Integer
    Dim intRowCounter       As Integer
    
    For intRowCounter = 1 To (frmRoster.GridRoster.Rows - 1)
        sinRowHeight = sinRowHeight + frmRoster.GridRoster.RowHeight(intRowCounter)
        If Y > sinRowHeight And Y < sinRowHeight + frmRoster.GridRoster.RowHeight(intRowCounter) Then
            intRow = intRowCounter
        End If
    Next intRowCounter
    
    For intColCounter = 1 To (frmRoster.GridRoster.Cols - 1)
        sinColWidth = sinColWidth + frmRoster.GridRoster.ColWidth(intColCounter)
        If X > sinColWidth And X < sinColWidth + frmRoster.GridRoster.ColWidth(intColCounter) Then
            intCol = intColCounter
        End If
    Next intColCounter
    
    If intCol > 0 And intRow > 0 Then
        frmRoster.GridRoster.Row = intRow
        frmRoster.GridRoster.Col = intCol
        frmRoster.GridRoster.Text = frmRoster.ListStaff.List(frmRoster.ListStaff.ListIndex)
    End If
    

End Sub

Sub FillStaffList()
    
    '[PLACE VALUES FROM THE STAFF DYNASET INTO THE STAFF FORM LIST]
    Dim intCounter As Integer
    If DsStaff.RecordCount = 0 Then AddNewStaff
    DsStaff.MoveFirst
    
    '[CLEAR ALL CURRENT CONTENTS IN THE LIST]
    frmStaff.ListStaff.Clear
    
    '[MOVE THROUGH STAFF DYNASET AND FILL LIST]
    Do While Not DsStaff.EOF
        frmStaff.ListStaff.AddItem DsStaff("LastName") & ", " & DsStaff("FirstName")
        DsStaff.MoveNext
    Loop
    
End Sub

Sub FillStaffRosterList()
    
    '[PLACE VALUES FROM THE STAFF DYNASET INTO THE STAFF FORM LIST]
    Dim intCounter          As Integer
    Dim intClassChoice      As Integer
    Dim strClassKey         As String
    Dim DsBookmark
    
    '[SAVE CURRENT LOCATION IN DYNASET]
    If DsStaff.EOF Or DsStaff.BOF Then
        '[MOVE TO FIRST POSITION OFF BEGINNING/END OF TABLE]
        DsStaff.MoveFirst
    End If
    DsBookmark = DsStaff.Bookmark
        
    If frmRoster.ComboClass.ListIndex = -1 Then Exit Sub
    
    '[MAKE CLASS CHOICE]
    intClassChoice = frmRoster.ComboClass.ItemData(frmRoster.ComboClass.ListIndex)
    strClassKey = "Class_" & Trim(Str(intClassChoice))
    
    If DsStaff.RecordCount = 0 Then Exit Sub
    DsStaff.MoveFirst
    
    '[CLEAR ALL CURRENT CONTENTS IN THE LIST]
    frmRoster.ListStaff.Clear
    
    '[MOVE THROUGH STAFF DYNASET AND FILL LIST]
    Do While Not DsStaff.EOF
        '[ADD STAFF IF CLASS MATCHES AND STAFF MEMBER IS AVAILABLE]
        If (DsStaff(strClassKey) = vbChecked) Then frmRoster.ListStaff.AddItem DsStaff("LastName") & ", " & DsStaff("FirstName")
        DsStaff.MoveNext
    Loop

    '[RESTORE STAFF POSITION]
    DsStaff.Bookmark = DsBookmark
    
End Sub
Sub LocateStaff()

    '[FIND STAFF MEMBER IN DYNASET AND UPDATE ALL STAFF TEXT BOXES]
    Dim strFullString   As String   '[Temporary full string lastname, firstname]
    Dim strSurname      As String   '[Temporary surname string]
    Dim strFirstName    As String   '[Temporary firstname string]
    Dim intDelimiter    As Integer  '[Temporary location of surname, firstname delimiter]
    Dim SQLStmt         As String   '[Search string]
    Dim intCounter      As Integer
    Dim strField        As String   '[String representation of Class]
    
    '[IF LIST IS EMPTY THEN EXIT]
    If frmStaff.ListStaff.ListIndex < 0 Then Exit Sub
    
    '[CHECK TO SEE IF WE ARE MOVING FROM AN UNSAVED RECORD]
    If frmStaff.cmdSave.Visible = True Then
        '[RECORD IS UNSAVED, POPUP YES/NO DIALOG]
        Dim Msg As String
        Dim Style
        Dim Response
        Dim Title
        
        Msg = "You have made changes to this staff record (" & DsStaff("Lastname") & ", " & DsStaff("FirstName") & ") but have not saved these changes." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "If you choose not to save now, any changes you have made since your last save will be lost." & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & "Do you wish to save these changes before you move to another record ?"
        Style = vbYesNo ' Define buttons.
        Title = "Staff Changes Not Saved"  ' Define title.
        Response = gsrMsg(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
        
            Dim strDisplayname      As String
        
            '[COMMIT RECORD CHANGES TO THE DYNASET]
            SaveStaffDetails
            
            '[SAVE DISPLAYED NAME TO TEMPORARY STRING]
            strDisplayname = frmStaff.ListStaff.List(frmStaff.ListStaff.ListIndex)
            
            '[FILL STAFF LIST SO WE GET ORDER]
            FillStaffList
                
            '[RELOCATE STAFF NAME]
            For intCounter = 0 To (frmStaff.ListStaff.ListCount - 1)
                If frmStaff.ListStaff.List(intCounter) = strDisplayname Then frmStaff.ListStaff.ListIndex = intCounter
            Next intCounter
        
        End If
        
    End If
    
    '[Delimit list item, break into Surname, FirstName]
    strFullString = frmStaff.ListStaff.List(frmStaff.ListStaff.ListIndex)
    If strFullString = "" Then Exit Sub
    
    intDelimiter = InStr(strFullString, ",")
    strSurname = Trim(Left(strFullString, intDelimiter - 1))
    strFirstName = Trim(Mid(strFullString, intDelimiter + 1))
    
    '[LOCATE LASTNAME, FIRSTNAME IN DYNASET]
    SQLStmt = "LastName = '" & strSurname & "' AND FirstName = '" & strFirstName & "'"
    DsStaff.FindFirst SQLStmt
        
    '[IF CANNOT FIND STAFF MEMBER (?) THEN EXIT]
    If DsStaff.NoMatch Then Exit Sub
    
    '[UPDATE STAFF DETAILS ON FORM]
    ShowStaffDetails
    
End Sub

Sub ProcessDayLength()

    '[PROCESS CHANGES TO COMBO BOXES ON THE CONTROL FORM]
    Select Case frmControl.OptionTIme(0).Value
    Case True
        DsDefault.Edit
            DsDefault("StartTime") = frmControl.ComboHour.Text & ":" & frmControl.ComboMinute.Text
        DsDefault.Update
    Case False
        DsDefault.Edit
            DsDefault("EndTime") = frmControl.ComboHour.Text & ":" & frmControl.ComboMinute.Text
        DsDefault.Update
    End Select
        
    '[CALCULATE DAY LENGTH AND DISPLAY]
    If DsDefault("StartTime") > DsDefault("EndTime") Then
        frmControl.MaskDayLength.Text = CDate("12:00") + CDate("12:00") + (DsDefault("EndTime") - DsDefault("StartTime"))
    Else
        frmControl.MaskDayLength.Text = DsDefault("StartTime") - DsDefault("EndTime")
    End If
    
End Sub

Sub FillClassGrid()
    
    '[PLACE VALUES FROM THE CLASS DYNASET INTO THE MAIN FORM CLASS GRID]
    Dim intCounter As Integer
    DsClass.MoveFirst
    
    frmControl.GridClass.Row = 0
    frmControl.GridClass.Col = 0
    frmControl.GridClass.Text = "Code"
    frmControl.GridClass.Col = 1
    frmControl.GridClass.Text = "Description"
    frmControl.GridClass.Col = 2
    frmControl.GridClass.Text = ""
    
    '[Make sure only 10 editable rows are included]
    frmControl.GridClass.Rows = 11
    
    For intCounter = 1 To 10
        
        frmControl.GridClass.Row = intCounter
        frmControl.GridClass.Col = 0
        frmControl.GridClass.Text = Trim(DsClass("Code") & " ")
        frmControl.GridClass.Col = 1
        frmControl.GridClass.Text = Trim(DsClass("Description") & " ")
        frmControl.GridClass.Col = 2
        Select Case DsClass("Active")
        Case vbChecked
            frmControl.GridClass.Picture = frmControl.ImageSwitch(constWarning).Picture
        Case vbUnchecked
            frmControl.GridClass.Picture = frmControl.ImageSwitch(constCritical).Picture
        Case Else
            frmControl.GridClass.Picture = frmControl.ImageSwitch(constSerious).Picture
        End Select
        frmControl.GridClass.Text = DsClass("Active")
        
        DsClass.MoveNext
        
    Next intCounter
    
    
End Sub

Sub Initialise()
    
    '[Assign dynasets to the public variables]
    Dim SQLStmt             As String
    Dim strRegCode          As String
    
    Set DBMain = Workspaces(0).OpenDatabase("GSR.DAT")
    
    '[SQLSTMT TO ORDER STAFF DYNASET]
    SQLStmt = "SELECT * FROM Staff ORDER BY LastName, FirstName;"
        
    Set DsStaff = DBMain.OpenRecordset(SQLStmt, dbOpenDynaset)
    Set DsClass = DBMain.OpenRecordset("Class", dbOpenDynaset)
    Set DsRoster = DBMain.OpenRecordset("Roster", dbOpenDynaset)
    Set DsDefault = DBMain.OpenRecordset("Defaults", dbOpenDynaset)
    
    '[Change to first record in default dynaset]
    DsDefault.MoveFirst

    '[Fill Day of Week Array]
    ArrayWeek(1).LongDay = "Sunday":        ArrayWeek(1).ShortDay = "Sun"
    ArrayWeek(2).LongDay = "Monday":        ArrayWeek(2).ShortDay = "Mon"
    ArrayWeek(3).LongDay = "Tuesday":       ArrayWeek(3).ShortDay = "Tue"
    ArrayWeek(4).LongDay = "Wednesday":     ArrayWeek(4).ShortDay = "Wed"
    ArrayWeek(5).LongDay = "Thursday":      ArrayWeek(5).ShortDay = "Thu"
    ArrayWeek(6).LongDay = "Friday":        ArrayWeek(6).ShortDay = "Fri"
    ArrayWeek(7).LongDay = "Saturday":      ArrayWeek(7).ShortDay = "Sat"

    '[LOAD FORMS]
    Load frmStaff
    Load frmControl
    Load frmRoster

    '[ARRANGE FORMS IN A CASCADE]
    'mdiMain.Arrange vbCascade

    '[RESTORE WINDOW STATE FROM DEFAULT DYNASET]
    '[0 = HIDDEN]
    '[1 = NORMAL]
    '[2 = MAXIMISED]
    
    If DsDefault("ControlState") = 0 Then
        mdiMain.cmdControlWindow.Value = False
    End If
    If DsDefault("RosterState") = 0 Then
        mdiMain.cmdRosterWindow.Value = False
    End If
    If DsDefault("StaffState") = 0 Then
        mdiMain.cmdStaffWindow.Value = False
    End If
    
    If DsDefault("ControlState") = 2 Then
        frmControl.WindowState = 2
    End If
    If DsDefault("RosterState") = 2 Then
        frmRoster.WindowState = 2
    End If
    If DsDefault("StaffState") = 2 Then
        frmStaff.WindowState = 2
    End If
    If DsDefault("MainState") = 2 Then
        mdiMain.WindowState = 2
    End If


    '[APPLY FONT TO ROSTER GRID]
    If DsDefault("RosterFontName") > "" Then
        Dim intCounter      As Integer
        Dim flagFontFound   As Boolean
        
        flagFontFound = False
        
        '[CEHCK THAT THIS FONT EXISTS]
        For intCounter = 0 To Printer.FontCount - 1  ' Determine number of fonts.
            If DsDefault("RosterFontName") = Printer.Fonts(intCounter) Then flagFontFound = True
        Next intCounter

        If flagFontFound = True Then
            frmRoster.GridRoster.Font.Name = DsDefault("RosterFontName")
            frmRoster.GridRoster.Font.Bold = DsDefault("RosterFontBold")
            frmRoster.GridRoster.Font.Italic = DsDefault("RosterFontItalic")
            frmRoster.GridRoster.Font.Size = DsDefault("RosterFontSize")
            '[APPLY FONT TO CONTROL GRID]
            Set frmControl.GridClass.Font = frmRoster.GridRoster.Font
            '[APPLY FONT TO WHOLE ROSTER FORM]
            Set frmRoster.Font = frmRoster.GridRoster.Font
        End If
        
    End If
    
    '[CHECK FIRST INSTALLED DATE AND IF ISNULL THEN PLACED CURRENT DATE IN THE FIELD]
    If IsNull(DsDefault("InstallDate")) Or Not (IsDate(DsDefault("InstallDate"))) Then
        DsDefault.Edit
            DsDefault("InstallDate") = Format(Now, "Short Date")
        DsDefault.Update
    End If
    
    '[SET DAYS USED PUBLIC VARIABLE]
    intDaysUsed = CDate(Format(Now, "Short Date")) - CDate(Format(DsDefault("InstallDate"), "Short Date"))
    
    '[SET MODIFIER AND CALL VALIDATION ROUTINE]
    sinModifier = 129#
    '[------------------------------------------------------------------------------------------]
    '[-VALIDATION ROUTINES HERE-----------------------------------------------------------------]
    '[------------------------------------------------------------------------------------------]
    strRegCode = Validate(DsDefault("RegUser"))
    If strRegCode = "" Or Not (strRegCode = DsDefault("RegCode")) Or IsNull(DsDefault("RegCode")) Then
        '[CODE DOESN'T MATCH - REPLACE CODE IN DATABASE]
        DsDefault.Edit
            If intDaysUsed > 45 Then
                DsDefault("RegUser") = "Unregistered Version"
                DsDefault("RegCode") = ""
            Else
                DsDefault("RegUser") = "Shareware Evaluation Version"
                DsDefault("RegCode") = ""
            End If
        DsDefault.Update
    End If
    '[------------------------------------------------------------------------------------------]
    
    '[RESTORE DELETE CONFIRM FLAG]
    frmControl.CheckDelete.Value = DsDefault("DeleteConfirm")
    
    '[CHANGE PUBLIC VARIABLE DELETE CONFIRMATION]
    If frmControl.CheckDelete.Value = 1 Then
        flagDeleteConfirm = True
    Else
        flagDeleteConfirm = False
    End If
       
    '[SET ROSTER TO THAT LAST WORKED ON]
    If (frmRoster.ComboClass.ListCount - 1) < DsDefault("RosterID") Then
        frmRoster.ComboClass.ListIndex = (frmRoster.ComboClass.ListCount - 1)
    Else
        frmRoster.ComboClass.ListIndex = DsDefault("RosterID")
    End If

    '[MOVE ROSTER FORM TO THE FRONT]
    If frmRoster.Visible = True Then
        frmRoster.ZOrder
    End If
    
    
End Sub

Sub RebuildRoster()

    '[REBUILD ROSTER USING START-TIME, END-TIME AND INTERVAL]
    Dim intRow          As Integer
    Dim intCol          As Integer
    Dim dateTemp        As Date
    Dim dateStart       As Date
    Dim dateEnd         As Date
    Dim intInterval     As Integer
    Dim flagContinue    As Boolean
    Dim strInterval     As String
    Dim flagNextDay     As Boolean
    
    dateStart = DsDefault("StartTime")
    dateTemp = DsDefault("StartTime")
    dateEnd = DsDefault("EndTime")
    
    '[ADD 24 HOURS IF END TIME IS THE NEXT DAY]
    If dateEnd < dateStart Then dateEnd = dateEnd + CDate("12:00") + CDate("12:00")
    flagContinue = True
    
    '[DETERMINE INTERVAL]
    intInterval = frmControl.ComboIncrement.ItemData(DsDefault("Increment") - 1)
    
    Select Case intInterval
    Case 30
        strInterval = "00:30"
    Case 15
        strInterval = "00:15"
    Case Else
        strInterval = Str(Format(intInterval, "0#")) & ":00"
    End Select
    
    '[REDUCE ROWS TO TWO AND REBUILD]
    frmRoster.GridRoster.Rows = 2
    frmRoster.GridRoster.Row = 1
    
    '[BUILD ROWS]
    Do While flagContinue = True
        frmRoster.GridRoster.AddItem Format(dateTemp, "Medium Time") & Chr$(vbKeyTab) & Format(dateTemp + CDate(strInterval), "Medium Time"), (frmRoster.GridRoster.Rows - 1)
        '[OLD TIME FORMAT] -> "hh:mm AMPM"
        dateTemp = dateTemp + CDate(strInterval)
        If dateTemp > dateEnd Then flagContinue = False
    Loop
    
    '[REMOVE LAST BLANK ROW]
    frmRoster.GridRoster.Rows = (frmRoster.GridRoster.Rows - 1)

    

End Sub

Sub SetClassLabels()
    
    '[SET CLASS DESCRIPTIONS TO MATCH CLASS CHECK BOX LABELS]
    Dim intCounter      As Integer
    Dim intClassIndex   As Integer      '[LOCATION OF SELECTED ITEM IN COMBOCLASS LIST]
    Dim strClass        As String
    DsClass.MoveFirst
    intCounter = 0
    
    '[SAVE CURRENTLY DISPLAYED CLASS ITEM INDEX]
    strClass = frmRoster.ComboClass.Text
    
    '[IF NOTHING SELECTED, CHOOSE THE FIRST ITEM]
    If intClassIndex = -1 Then intClassIndex = 0
    
    '[CLEAR COMBO BOX]
    frmRoster.ComboClass.Clear
    
    Do While Not DsClass.EOF
        frmStaff.CheckClass(intCounter).Caption = DsClass("Description")

        '[ALSO UPDATE COMBO BOX ON ROSTER FORM]
        If DsClass("Active") = vbChecked Then
            frmRoster.ComboClass.AddItem DsClass("Description")
            frmRoster.ComboClass.ItemData(frmRoster.ComboClass.NewIndex) = (intCounter + 1)
        End If
        
        intCounter = intCounter + 1
        DsClass.MoveNext
        
    Loop
    
    '[RESTORE CURRENTLY DISPLAYED CLASS ITEM INDEX]
    For intCounter = 0 To frmRoster.ComboClass.ListCount
        If frmRoster.ComboClass.List(intCounter) = strClass And strClass > "" Then frmRoster.ComboClass.ListIndex = intCounter
    Next intCounter

End Sub


Sub SetDayLabels()

    '[SET LABELS ON STAFF FORM TO THE ORDER SPECIFIED BY STARTDAY]
    Dim intCounter      As Integer
    For intCounter = 1 To 7
        frmStaff.CheckDay(intCounter - 1).Caption = ArrayWeek(DayOfWeek(intCounter)).LongDay
    Next intCounter
    
End Sub

Sub SetGridTitles()

    '[SET GRID TITLES ON THE ROSTER GRID ACCORDING TO SELECTED START DAY]
    Dim intCounter      As Integer
    Dim intCol          As Integer
    Dim intRow          As Integer
    Dim intStartDay     As Integer
    
    '[SAVE CURRENT GRID POSITION]
    intRow = frmRoster.GridRoster.Row
    intCol = frmRoster.GridRoster.Col

    frmRoster.GridRoster.Row = 0
    
    '[CYCLE THROUGH COLUMNS AND SET TITLES]
    frmRoster.GridRoster.Col = 0
    frmRoster.GridRoster.Text = "Start"
    frmRoster.GridRoster.Col = 1
    frmRoster.GridRoster.Text = "End"
    
    For intCounter = 2 To 8
        frmRoster.GridRoster.Col = intCounter
        frmRoster.GridRoster.Text = ArrayWeek(DayOfWeek(intCounter - 1)).ShortDay
    Next intCounter

    '[RESTORE GRID POSITION]
    frmRoster.GridRoster.Row = intRow
    frmRoster.GridRoster.Col = intCol
    
End Sub

Function DayOfWeek(intDayNumber)

    '[FUNCTION TO RETURN INTEGER FOR DAY OF WEEK ARRAY]
    Dim intStartDay             As Integer
    Dim intReturnValue          As Integer
    
    intStartDay = DsDefault("StartDay")
    
    If intStartDay + (intDayNumber - 1) > 7 Then
        intReturnValue = (intDayNumber - 7) + (intStartDay - 1)
    Else
        intReturnValue = (intStartDay + (intDayNumber - 1))
    End If

    DayOfWeek = intReturnValue

End Function
Sub ShowStaffDetails()
        
    '[SHOW STAFF DETAILS FOR THE CURRENTLY SELECTED STAFF MEMBER IN THE DYNASET]
    Dim intCounter      As Integer
    Dim strField        As String
    
    '[TEXT BOXES]
    frmStaff.TextStaffID.Text = DsStaff("StaffID")
    '[THE " " IS REQUIRED IN CASE OF NULL VALUES IN THE DATABASE (ALTHOUGH THIS SHOULDN'T OCCUR)]
    frmStaff.TextLastName.Text = Trim(DsStaff("LastName") & " ")
    frmStaff.TextFirstName.Text = Trim(DsStaff("FirstName") & " ")
    frmStaff.TextMiddleName.Text = Trim(DsStaff("MiddleName") & " ")
    
    frmStaff.MaskHomePhone.Text = Trim(DsStaff("HomePhone") & " ")
    frmStaff.MaskHourRate.Text = Trim(DsStaff("HourRate") & " ")
    frmStaff.MaskBirthDate.Text = Trim(DsStaff("BirthDate") & " ")
    frmStaff.MaskDateHired.Text = Trim(DsStaff("DateHired") & " ")
    
    frmStaff.TextNote.Text = Trim(DsStaff("Note") & " ")
    
    '[MAX/MIN HOURS
    If DsStaff("MinHours") > 0 Then frmStaff.MaskMinHours.Text = DsStaff("MinHours") Else frmStaff.MaskMinHours.Text = ""
    If DsStaff("MaxHours") > 0 Then frmStaff.MaskMaxHours.Text = DsStaff("MaxHours") Else frmStaff.MaskMaxHours.Text = ""

    '[STAFF AVAILABILITY DAYS]
    For intCounter = 1 To 7
        strField = "Day_" & Trim(Str(intCounter))
        If IsNull(DsStaff(strField).Value) Then frmStaff.CheckDay(intCounter - 1).Value = vbChecked Else frmStaff.CheckDay(intCounter - 1).Value = DsStaff(strField).Value
    Next intCounter
    
    '[DAYS EMPLOYED]
    If DsStaff("DateHired") > 0 Then frmStaff.LabelDaysEmployed.Caption = Format(Date - DsStaff("DateHired"), "yy\y\r\s mm\m\t\h\s") Else frmStaff.LabelDaysEmployed.Caption = ""
    If DsStaff("BirthDate") > 0 Then frmStaff.LabelAge.Caption = Format(Date - DsStaff("BirthDate"), "yy\y\r\s mm\m\t\h\s") Else frmStaff.LabelAge.Caption = ""
    
    '[STAFF CLASSIFICATION CHECK BOXES]
    For intCounter = 1 To 10
        strField = "Class_" & Trim(Str(intCounter))
        frmStaff.CheckClass(intCounter - 1).Value = DsStaff(strField).Value
    Next intCounter

    '[SET SAVE BUTTON TO DISABLED]
    frmStaff.cmdSave.Visible = False

End Sub


Sub SaveStaffDetails()

    '[SAVE STAFF DETAILS TO THE DYNASET FOR THE CURRENTLY DISPLAYED STAFF MEMBER]
    Dim intCounter      As Integer
    Dim strField        As String
    
    '[OPEN DYNASET FOR EDITING]
    DsStaff.Edit
    
        '[TEXT BOXES]
        '[CHECK TO SEE IF THE STAFF ID IS BLANK - IF SO, REPLACE WITH OLD STAFF ID]
        If frmStaff.TextStaffID <> "" Then DsStaff("StaffID") = frmStaff.TextStaffID.Text
        If frmStaff.TextLastName.Text <> "" Then DsStaff("LastName") = frmStaff.TextLastName.Text
        If frmStaff.TextFirstName.Text <> "" Then DsStaff("FirstName") = frmStaff.TextFirstName.Text
        DsStaff("MiddleName") = frmStaff.TextMiddleName.Text
    
        DsStaff("HomePhone") = frmStaff.MaskHomePhone.Text
        DsStaff("HourRate") = frmStaff.MaskHourRate.Text
        DsStaff("BirthDate") = frmStaff.MaskBirthDate.Text
        DsStaff("DateHired") = frmStaff.MaskDateHired.Text
        DsStaff("Note") = frmStaff.TextNote.Text
                
        '[STAFF MAX/MIN HOURS]
        DsStaff("MinHours") = frmStaff.MaskMinHours.Text
        DsStaff("MaxHours") = frmStaff.MaskMaxHours.Text
                
        '[STAFF AVAILABILITY DAYS]
        For intCounter = 1 To 7
            strField = "Day_" & Trim(Str(intCounter))
            DsStaff(strField) = (frmStaff.CheckDay(intCounter - 1).Value)
        Next intCounter
                
        '[STAFF CLASSIFICATION CHECK BOXES]
        For intCounter = 1 To 10
            strField = "Class_" & Trim(Str(intCounter))
            DsStaff(strField) = (frmStaff.CheckClass(intCounter - 1).Value)
        Next intCounter
    
    '[COMMIT CHANGES TO DYNASET]
    DsStaff.Update

    '[SET SAVE BUTTON TO DISABLED]
    frmStaff.cmdSave.Visible = False

End Sub
