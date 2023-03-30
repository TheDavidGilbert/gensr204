DECLARE FUNCTION ReadCODESfile% ()
DECLARE FUNCTION CodesMenu% (ptrCodesMenu%)
DECLARE SUB CODESLIST ()
DECLARE FUNCTION OpenCodesList% ()
DECLARE SUB SetWindowButton (intWINDOW%)
DECLARE FUNCTION GetKey$ ()
DECLARE SUB DisplayStaffScreens ()
DECLARE SUB STAFFLIST ()
DECLARE SUB HelpBox (Line1$, Line2$)
DECLARE SUB STARTUP ()
DECLARE SUB SHUTDOWN ()

DECLARE FUNCTION MainMenu% (ptrMainMenu%)
DECLARE FUNCTION DisplayItemList% (MenuList$(), MenuCount%, Caption$, MenuPos%)
DECLARE SUB MAIN ()
DECLARE FUNCTION OpenStaffList% ()

DEFINT A-Z
REM $INCLUDE: 'COMMON.BI'

' [------------------------------------------------]
' [GENERIC STAFF ROSTER PROGRAM - GSR.EXE          ]
' [QUICKBASIC 4.5 COMPILER SOURCE CODE             ]
' [------------------------------------------------]
' [DAVID GILBERT, JUNE 1997                        ]
' [------------------------------------------------]
' [(A) MAINTAIN STAFF LIST ------------------------]
' [(B) PRODUCE WEEKLY STAFF ROSTER ----------------]
' [(C) PRINT INDIVIDUAL STAFF ROSTERS -------------]
' [------------------------------------------------]
' [STAFF.LST ---- STAFF MEMBERS AND DETAILS--------]
' [------------------------------------------------]
' [DATA DECLARATIONS AND TYPES---------------------]
CONST RegUser = "Coffee Club, Toowoomba"


' [------------------------------------------------]
STARTUP
MAIN
' [------------------------------------------------]

SUB CODESLIST

' [------------------------------------------------]
' [CODE PROCEDURES HERE----------------------------]
' [------------------------------------------------]
DIM tmpCODE AS STRING * 1
DIM tmpDESC AS STRING * 20
DIM tmpFLAG AS INTEGER
DIM tmpKey$

ptrCODEMenu = 1

DO WHILE ptrCODEMenu > 0

	ptrCODEMenu = CodesMenu(ptrCODEMenu)
	tmpFLAG = 1

' [------------------------------------------------]
' [PROCESS SELECTED ITEM FROM CODE MENU------------]
' [------------------------------------------------]
	SELECT CASE ptrCODEMenu
	CASE IS > 0 < 11  ' [CODE ITEM SELECTED]
		Inbuff$ = recCodes(ptrCODEMenu).CODE
		IF EditString(10, 25, 4, 30, WOR%, "Enter Code", Inbuff$, 1) > 0 THEN
			' [GET CODE LETTER/IDENTIFIER]
			tmpCODE = Inbuff$
			Inbuff$ = recCodes(ptrCODEMenu).DESC
			IF EditString(10, 25, 4, 30, WOR%, "Enter Code Description", Inbuff$, 20) > 0 THEN
				' [GET CODE DESCRIPTION]
				tmpDESC = Inbuff$
				' [CHECK TO SEE IF CODE OR DESCRIPTION ALREADY EXIST IN ARRAY]
				FOR tmpCOUNTER = 1 TO 10
					IF tmpCODE <> " " AND tmpCOUNTER <> ptrCODEMenu AND (recCodes(tmpCOUNTER).DESC = tmpDESC OR recCodes(tmpCOUNTER).CODE = tmpCODE) THEN tmpFLAG = 0
				NEXT tmpCOUNTER

				IF tmpFLAG = 1 THEN
					' [ASSIGN ENTERED DATA TO ARRAY]
					IF tmpCODE = " " THEN tmpDESC = ""
					recCodes(ptrCODEMenu).CODE = tmpCODE
					recCodes(ptrCODEMenu).DESC = tmpDESC
					' [WRITE CHANGED CODE TO FILE]
					PUT #intFILEcodes, ptrCODEMenu, recCodes(ptrCODEMenu)
				ELSE
					' [DISPLAY ERROR POPUP]
					PopUp 8, 15, 5, 55, WOR, 2, 3, 0
					FastPrint 9, 16, "       This code/description already exists.", WOR%
					FastPrint 11, 16, "            [Press any key to continue]", YOR%
					Choice = -1
					tmpKey$ = GetKey
					ShutUp (0)
				END IF

			END IF
		END IF

	CASE ELSE

	END SELECT


' [------------------------------------------------]
' [------------------------------------------------]
' [------------------------------------------------]

LOOP


END SUB

FUNCTION CodesMenu (ptrCodesMenu)

	DIM mnuCODES$(10)
 
	FOR tmpCOUNTER = 1 TO 10
		mnuCODES$(tmpCOUNTER) = recCodes(tmpCOUNTER).CODE + ". " + recCodes(tmpCOUNTER).DESC
	NEXT tmpCOUNTER

	CALL HelpBox("Use cursor keys to navigate menu and press ENTER to modify", "Press ESC to return to the main menu")

	CodesMenu = DisplayItemList(mnuCODES$(), 10, "Customise Codes", ptrCodesMenu)

END FUNCTION

FUNCTION DisplayItemList (MenuList$(), MenuCount, Caption$, MenuPos)

	REM $INCLUDE: 'LISTDISP.BI'
	DisplayItemList = MenuChoice

END FUNCTION

FUNCTION DisplayList (MenuList$(), MenuCount, Caption$)
  
	REM $INCLUDE: 'NAMEDISP.BI'
	DisplayList = MenuChoice

END FUNCTION

SUB DisplayStaffScreens

' [CREATE POPUP SCREENS AND SET LABELS-------------]
	stdTop = 3  ' Location of first line
	CALL HelpBox("Use TAB key to switch windows and press ENTER to select", "ESCAPE returns you to the main menu")

' [STAFF LIST SCREEN-------------------------------]
	PopUp stdTop, 2, 18, 31, BOWH%, 3, 0, 20              ' STAFF LIST SCREEN
	FastPrint stdTop, 11, " STAFF LIST ", ROWH%
' [------------------------------------------------]


' [STAFF DETAILS SCREEN----------------------------]
	PopUp stdTop, 33, 9, 47, WOG%, 3, 0, 20              ' STAFF DETAILS SCREEN
	FastPrint stdTop, 49, " STAFF DETAILS ", BOG%
	FastPrint stdTop + 1, 34, "ID", YOG%
	FastPrint stdTop + 1, 39, "SURNAME", YOG%
	FastPrint stdTop + 1, 55, "FIRSTNAME", YOG%
	FastPrint stdTop + 3, 34, "DOB", YOG%
	FastPrint stdTop + 3, 44, "RATE", YOG%
	FastPrint stdTop + 3, 54, "CONTACT", YOG%
	FastPrint stdTop + 5, 34, "CODE :", YOG%
' [------------------------------------------------]


' [STAFF AVAILABILITY SCREEN-----------------------]
	stdTop = stdTop + 2
	PopUp stdTop + 7, 33, 9, 47, WOBL%, 3, 0, 20          ' STAFF AVAILABILITY SCREEN
	FastPrint stdTop + 7, 47, " STAFF AVAILABILITY ", YOBL%
	FastPrint stdTop + 8, 37, "MON", YOBL%
	FastPrint stdTop + 8, 43, "TUE", YOBL%
	FastPrint stdTop + 8, 49, "WED", YOBL%
	FastPrint stdTop + 8, 55, "THU", YOBL%
	FastPrint stdTop + 8, 61, "FRI", YOBL%
	FastPrint stdTop + 8, 67, "SAT", YOBL%
	FastPrint stdTop + 8, 73, "SUN", YOBL%
	FastPrint stdTop + 9, 34, "A", WOBL%
	FastPrint stdTop + 10, 34, "M", WOBL%
	FastPrint stdTop + 12, 34, "P", WOBL%
	FastPrint stdTop + 13, 34, "M", WOBL%
' [------------------------------------------------]


END SUB

DEFSNG A-Z
FUNCTION EDITSTRING (Row AS INTEGER, Col AS INTEGER, Hght AS INTEGER, Wdth AS INTEGER, HAttr AS INTEGER, Prompt AS STRING, Inbuff$, EditWdth AS INTEGER)

	REM $INCLUDE: 'EDITSTR.BI'

END FUNCTION

DEFINT A-Z
FUNCTION GetKey$

	strKeyPress$ = ""
	DO WHILE strKeyPress$ = "": strKeyPress$ = INKEY$: LOOP
	GetKey = strKeyPress$

END FUNCTION

SUB HelpBox (Line1$, Line2$)

	REM $INCLUDE: 'HELPBOX.BI'

END SUB

SUB MAIN

' [------------------------------------------------]
' [MAIN PROCEDURES HERE----------------------------]
' [------------------------------------------------]

ptrMainMenu = 1

DO WHILE ptrMainMenu > 0

	ptrMainMenu = MainMenu(ptrMainMenu)

' [------------------------------------------------]
' [PROCESS SELECTED ITEM FROM MAIN MENU------------]
' [------------------------------------------------]

	SELECT CASE ptrMainMenu
	CASE 1   ' [CONFIGURE PROGRAM DEFAULTS (CODES/PRINT ETC)----]
		CODESLIST

	CASE 2   ' [STAFF LIST MODULE-------------------------------]
		STAFFLIST

	CASE 3   ' [------------------------------------------------]
  
	CASE 4   ' [------------------------------------------------]
  
	CASE 5   ' [EXIT PROGRAM AND CLOSE--------------------------]
		SHUTDOWN
	CASE ELSE
  
	END SELECT

' [------------------------------------------------]
' [------------------------------------------------]
' [------------------------------------------------]

LOOP


END SUB

FUNCTION MainMenu (ptrMainMenu)

	cntMenuItems = 5

	DIM mnuMain$(cntMenuItems)
	mnuMain$(1) = "1.  Configure codes   "
	mnuMain$(2) = "2.  Staff list        "
	mnuMain$(3) = "3.  Create roster     "
	mnuMain$(4) = "4.  Print roster      "
	mnuMain$(5) = "5.  Exit              3"

	CALL HelpBox("Use cursor keys to navigate menu and press ENTER to select", "GSR (C) 1997 David Gilbert")

	MainMenu = DisplayItemList(mnuMain$(), cntMenuItems, "      Select option       ", ptrMainMenu)
 

END FUNCTION

FUNCTION OpenCodesList

' [------------------------------------------------]
' [OPEN CODES LIST FILE FOR RANDOM ACCESS AND -----]
' [RETURN FILEHANDLE-------------------------------]
' [------------------------------------------------]
' [FILE FORMAT-------------------------------------]
' [------------------------------------------------]
' [CODE-DESC---------------------------------------]
' [------------------------------------------------]

DIM intHandle AS INTEGER   ' Free file handle

	intHandle = FREEFILE
	OPEN "codes.lst" FOR RANDOM AS intHandle LEN = LEN(recCodes(1))

OpenCodesList = intHandle

END FUNCTION

FUNCTION OpenStaffList

' [------------------------------------------------]
' [OPEN STAFF LIST FILE FOR RANDOM ACCESS AND -----]
' [RETURN FILEHANDLE-------------------------------]
' [------------------------------------------------]
' [FILE FORMAT-------------------------------------]
' [------------------------------------------------]
' [ID-SURNAME-FIRSTNAME-CODE-DOB-CLASS-RATE-CONTACT]
' [------------------------------------------------]

DIM intHandle AS INTEGER   ' Free file handle

	intHandle = FREEFILE
	OPEN "staff.lst" FOR RANDOM AS intHandle LEN = LEN(recStaff)

OpenStaffList = intHandle

END FUNCTION

FUNCTION ReadCODESfile

	' [------------------------------------------------]
	' [READ IN CODES-----------------------------------]
	' [------------------------------------------------]

	' [READ IN CODES-----------------------------------]
	FOR tmpCOUNTER = 1 TO 10
		GET #intFILEcodes, tmpCOUNTER, recCodes(tmpCOUNTER)
	NEXT tmpCOUNTER
	' [------------------------------------------------]

	ReadCODESfile = 10

END FUNCTION

SUB SetWindowButton (intWINDOW)
	  
		SELECT CASE intWINDOW
		CASE 1      ' [STAFF LIST WINDOW---------]
			FastPrint 3, 3, "[þ]", ROWH%
			FastPrint 3, 34, "[ ]", WOG%
			FastPrint 12, 34, "[ ]", WOBL%
		CASE 2      ' [STAFF DETAILS WINDOW------]
			FastPrint 3, 3, "[ ]", ROWH%
			FastPrint 3, 34, "[þ]", WOG%
			FastPrint 12, 34, "[ ]", WOBL%
		CASE 3      ' [STAFF AVAILABILITY WINDOW-]
			FastPrint 3, 3, "[ ]", ROWH%
			FastPrint 3, 34, "[ ]", WOG%
			FastPrint 12, 34, "[þ]", WOBL%
		CASE ELSE
		END SELECT

END SUB

SUB SHUTDOWN

' [------------------------------------------------]
' [CLOSE ALL OPEN FILES AND CLEAR SCREEN-----------]
' [DISPLAY CLOSING INFORMATION---------------------]
' [------------------------------------------------]

' [------------------------------------------------]
	CLOSE
	CLS
	PRINT "Generic Staff Roster, (C) 1997 David Gilbert"
	END
' [------------------------------------------------]

END SUB

SUB STAFFLIST

' [DEFINE VARIABLES FOR STAFF LIST-----------------]
	DIM strKeyPress$              ' KEY PRESSED BY USER
	DIM intWINDOW     AS INTEGER  ' WINDOW POINTER (1,2,3)
	DIM CONTINUE
	CONTINUE = TRUE
	intWINDOW = 1

' [DISPLAY ROUTINE FOR STAFF DETAILS---------------]
' [CREATE STAFF LIST POPUP AND STAFF DETAILS SCREEN]
	DisplayStaffScreens
	SetWindowButton (intWINDOW)

' [------------------------------------------------]
' [MAIN STAFF LIST ROUTINE WITH KEY HANDLING-------]
' [------------------------------------------------]
  
	DO WHILE CONTINUE = TRUE
	' [GET USER KEY PRESS AND PROCESS------------------]
		strKeyPress$ = GetKey

		SELECT CASE strKeyPress$
	 
		CASE CHR$(15)                 ' [SHIFT-TAB------]
			' [CHANGE WINDOW BACKWARDS---]
			intWINDOW = intWINDOW - 1: IF intWINDOW < 1 THEN intWINDOW = 3

		CASE CHR$(9)                  ' [TAB------------]
			' [CHANGE WINDOW-------------]
			intWINDOW = intWINDOW + 1: IF intWINDOW > 3 THEN intWINDOW = 1
	  
		CASE CHR$(27)                 ' [ESCAPE---------]
			' [EXIT STAFFLIST FORWARDS---]
			CONTINUE = FALSE

		CASE ELSE

		END SELECT

		SetWindowButton (intWINDOW)

	LOOP

' [------------------------------------------------]
' [CLOSE ALL POPUP SCREENS-------------------------]
' [------------------------------------------------]
  
	ShutUp (20)
	ShutUp (20)
	ShutUp (20)

' [CLOSE STAFF LIST FILE---------------------------]
	CLOSE intHandle

END SUB

SUB STARTUP

' [SET SCREEN COLOR AND CLEAR SCREEN---------------]
	COLOR 14, 3: CLS
  
' [DISPLAY PROGRAM TITLES AND LINES----------------]
	Title$ = "Generic Staff Roster : " + RegUser
	FastPrint 1, 1, SPACE$(80), WOBL%
	FastPrint 1, 40 - (LEN(Title$) / 2), Title$, WOBL%
	FastPrint 2, 1, STRING$(80, "Ä"), BLOC%
	FastPrint 21, 1, STRING$(80, "Ä"), BLOC%

' [OPEN STAFF LIST FILE AND COUNT RECORDS----------]
	intFILEstaff = OpenStaffList
	cntRECstaff = LOF(intFILEstaff) / LEN(recStaff)

' [OPEN CODES LIST FILE AND COUNT RECORDS----------]
	intFILEcodes = OpenCodesList
	cntRECcodes = LOF(intFILEcodes) / LEN(recCodes(1))

' [READ IN CODES-----------------------------------]
	retCODE = ReadCODESfile

END SUB

