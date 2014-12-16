Option Explicit

' --------
' Globals
' --------
global VIBREOFFICE_STARTED as boolean ' Defaults to False
global VIBREOFFICE_ENABLED as boolean ' Defaults to False

global oXKeyHandler as object

' Global State
global MODE as string
global VIEW_CURSOR as object
global MULTIPLIER as integer

' -----------
' Singletons
' -----------
Function getCursor
	getCursor = VIEW_CURSOR
End Function

Function getTextCursor
    dim oTextCursor
    oTextCursor = getCursor().getText.createTextCursorByRange(getCursor())
    ' oTextCursor.gotoRange(oTextCursor.getStart(), False)

    getTextCursor = oTextCursor
End Function

' -----------------
' Helper Functions
' -----------------
Sub restoreStatus 'restore original statusbar
	dim oLayout
	oLayout = thisComponent.getCurrentController.getFrame.LayoutManager
	oLayout.destroyElement("private:resource/statusbar/statusbar")
	oLayout.createElement("private:resource/statusbar/statusbar")
End Sub

Sub setRawStatus(rawText)
	thisComponent.Currentcontroller.StatusIndicator.Start(rawText, 0)
End Sub

Sub setStatus(statusText)
	setRawStatus(MODE & " | " & statusText)
End Sub

Sub setMode(modeName)
	MODE = modeName
	setRawStatus(modeName)
End Sub

Function gotoMode(sMode)
    Select Case sMode
        Case "NORMAL":
            setMode("NORMAL")
        Case "INSERT":
            setMode("INSERT")
        Case "VISUAL":
            setMode("VISUAL")

            dim oTextCursor
            oTextCursor = getTextCursor()
            ' Deselect TextCursor
            oTextCursor.gotoRange(oTextCursor.getStart(), False)
            ' Show TextCursor selection
		    thisComponent.getCurrentController.Select(oTextCursor)
    End Select
End Function


' --------------------
' Multiplier functions
' --------------------
Sub _setMultiplier(n as integer)
	MULTIPLIER = n
End Sub

Sub resetMultiplier()
	_setMultiplier(0)
End Sub

Sub addToMultiplier(n as integer)
	dim sMultiplierStr as String
    dim iMultiplierInt as integer

    ' Max multiplier: 10000 (stop accepting additions after 1000)
    If MULTIPLIER <= 1000 then
        sMultiplierStr = CStr(MULTIPLIER) & CStr(n)
        _setMultiplier(CInt(sMultiplierStr))
    End If
End Sub

' Should only be used if you need the raw value
Function getRawMultiplier()
    getRawMultiplier = MULTIPLIER
End Function

' Same as getRawMultiplier, but defaults to 1 if it is unset (0)
Function getMultiplier()
	If MULTIPLIER = 0 Then
		getMultiplier = 1
	Else
		getMultiplier = MULTIPLIER
	End If
End Function


' -------------
' Key Handling
' -------------
Sub sStartXKeyHandler
	sStopXKeyHandler()

	oXKeyHandler = CreateUnoListener("KeyHandler_", "com.sun.star.awt.XKeyHandler")
	thisComponent.CurrentController.AddKeyHandler(oXKeyHandler)
End Sub

Sub sStopXKeyHandler
	thisComponent.CurrentController.removeKeyHandler(oXKeyHandler)
End Sub

Sub XKeyHandler_Disposing(oEvent)
End Sub


' --------------------
' Main Key Processing
' --------------------
function KeyHandler_KeyPressed(oEvent) as boolean
    ' Exit if plugin is not enabled
    If IsMissing(VIBREOFFICE_ENABLED) Or Not VIBREOFFICE_ENABLED Then
        KeyHandler_KeyPressed = False
        Exit Function
    End If

	dim bConsumeInput, bIsMultiplier, bIsModified, oTextCursor
	bConsumeInput = True ' Block all inputs by default
	bIsMultiplier = False ' reset multiplier by default
    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)

    ' --------------------------
	' Process global shortcuts, exit if matched (like ESC)
	If ProcessGlobalKey(oEvent) Then
		' Pass

	ElseIf MODE = "INSERT" Then
		bConsumeInput = False ' Allow all inputs

	' If Change Mode
	ElseIf ProcessModeKey(oEvent) Then
        ' Pass

    ElseIf ProcessNumberKey(oEvent) Then
        bIsMultiplier = True

    ' Normal Key
    ElseIf Not ProcessNormalKey(oEvent) and bIsModified Then
        ' If is modified but doesn't match a normal command, allow input
        '   (Useful for built-in shortcuts like Ctrl+s, Ctrl+w)
        bConsumeInput = False
    End If
    ' --------------------------


	' Reset multiplier
	If not bIsMultiplier Then resetMultiplier()
	setStatus(getMultiplier())

    ' Show terminal-like cursor
    oTextCursor = getTextCursor()
    If MODE = "NORMAL" Then
        oTextCursor.gotoRange(oTextCursor.getStart(), False)
        oTextCursor.goRight(1, False)
        oTextCursor.goLeft(1, True)
		thisComponent.getCurrentController.Select(oTextCursor)

    ElseIf MODE = "INSERT" Then
        oTextCursor.gotoRange(oTextCursor.getStart(), False)
		thisComponent.getCurrentController.Select(oTextCursor)
    End If

    KeyHandler_KeyPressed = bConsumeInput
End Function

Function KeyHandler_KeyReleased(oEvent) As boolean
    KeyHandler_KeyReleased = (MODE = "NORMAL") 'cancel KeyReleased
End Function


' ----------------
' Processing Keys
' ----------------
Function ProcessGlobalKey(oEvent)
	dim bMatched
	bMatched = True
	Select Case oEvent.KeyCode
		' PRESSED ESCAPE
		Case 1281:
            ' Move cursor back if was in INSERT (but stay on same line)
            If MODE <> "NORMAL" And Not getCursor().isAtStartOfLine() Then
                getCursor().goLeft(1, False)
            End If

			gotoMode("NORMAL")
		Case Else:
			bMatched = False
	End Select
	ProcessGlobalKey = bMatched
End Function


Function ProcessNumberKey(oEvent)
	dim c
	c = CStr(oEvent.KeyChar)

	If c >= "0" and c <= "9" Then
		addToMultiplier(CInt(c))
		ProcessNumberKey = True
	Else
		ProcessNumberKey = False
	End If
End Function


Function ProcessModeKey(oEvent)
	dim bMatched
	bMatched = True
	Select Case oEvent.KeyChar
        ' Insert modes
		Case "i", "a", "I", "A":
            If oEvent.KeyChar = "a" Then getCursor().goRight(1, False)
            If oEvent.KeyChar = "I" Then ProcessMovementKey("^")
            If oEvent.KeyChar = "A" Then ProcessMovementKey("$")

			gotoMode("INSERT")
        Case "v":
            gotoMode("VISUAL")
		Case Else:
			bMatched = False
	End Select
	ProcessModeKey = bMatched
End Function


Function ProcessNormalKey(oEvent)
    dim i, bMatched, bIsVisual
    bMatched = False
    bIsVisual = (MODE = "VISUAL") ' is this hardcoding bad? what about visual block?
    For i = 1 To getMultiplier()
        dim bMatchedMovement, bMatchedDelete

        bMatchedMovement = ProcessMovementKey(oEvent.KeyChar, bIsVisual, oEvent.Modifiers)
        bMatchedDelete = ProcessDeleteKey(oEvent)
        bMatched = bMatched or bMatchedMovement or bMatchedDelete

        ' Special case: Break from For loop if in visual mode and has deleted,
        ' since multiplier should not be applied
        If bIsVisual and bMatchedDelete Then Exit For
    Next i

    ProcessNormalKey = bMatched
End Function


Function ProcessDeleteKey(oEvent)
    dim oTextCursor, bMatched
    oTextCursor = getTextCursor()
    bMatched = True
    Select Case oEvent.KeyChar
        ' Case "d":
            ' setSpecial("d")

		Case "x":
			thisComponent.getCurrentController.Select(oTextCursor)
			oTextCursor.setString("")
        Case Else:
            bMatched = False

    End Select

    ProcessDeleteKey = bMatched
End Function


' -----------------------
' Main Movement Function
' -----------------------
'   Default: bExpand = False, keyModifiers = 0
Function ProcessMovementKey(keyChar, Optional bExpand, Optional keyModifiers)
	dim oTextCursor, bSetCursor, bMatched
    oTextCursor = getTextCursor()
    bMatched = True
    If IsMissing(bExpand) Then bExpand = False
    If IsMissing(keyModifiers) Then keyModifiers = 0


    ' Check for modified keys (Ctrl, Alt, not Shift)
    If keyModifiers > 1 Then
        dim isControl
        isControl = (keyModifiers = 2) or (keyModifiers = 8)

        ' Ctrl+d and Ctrl+u
        If isControl and keyChar = "d" Then
            getCursor().ScreenDown(bExpand)
        ElseIf isControl and keyChar = "u" Then
            getCursor().ScreenUp(bExpand)
        Else
            bMatched = False
        End If

        ProcessMovementKey = bMatched
        Exit Function
    End If

    ' Set global cursor to oTextCursor's new position if moved
    bSetCursor = True

	Select Case keyChar
		Case "l":
			oTextCursor.goRight(1, bExpand)
		Case "h":
			oTextCursor.goLeft(1, bExpand)

		' oTextCursor.goUp and oTextCursor.goDown SHOULD work, but doesn't (I dont know why).
		' So this is a weird hack
		Case "k":
			'oTextCursor.goUp(1, False)
			getCursor().goUp(1, bExpand)
            bSetCursor = False
		Case "j":
			'oTextCursor.goDown(1, False)
			getCursor().goDown(1, bExpand)
            bSetCursor = False
		' ----------

		Case "^":
            getCursor().gotoStartOfLine(bExpand)
            bSetCursor = False
		Case "$":
            dim oldPos, newPos
            oldPos = getCursor().getPosition()
            getCursor().gotoEndOfLine(bExpand)
            newPos = getCursor().getPosition()

            ' If the result is at the start of the line, then it must have
            ' jumped down a line; goLeft to return to the previous line.
            '   Except for: Empty lines (check for oldPos = newPos)
            If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
                getCursor().goLeft(1, bExpand)
            End If

            ' maybe eventually cursorGoto... should return True/False for bsetCursor
            bSetCursor = False

		Case "w", "W":
			oTextCursor.gotoNextWord(bExpand)
		Case "b", "B":
			oTextCursor.gotoPreviousWord(bExpand)
		Case "e":
			oTextCursor.gotoEndOfWord(bExpand)

		Case ")":
			oTextCursor.gotoNextSentence(bExpand)
		Case "(":
			oTextCursor.gotoPreviousSentence(bExpand)
		Case "}":
			oTextCursor.gotoNextParagraph(bExpand)
		Case "{":
			oTextCursor.gotoPreviousParagraph(bExpand)
		Case Else:
            bSetCursor = False
            bMatched = False
	End Select

    ' If oTextCursor was moved, set global cursor to its position
    If bSetCursor Then
        getCursor().gotoRange(oTextCursor.getStart(), False)
    End If

    ' If oTextCursor was moved and is in VISUAL mode, update selection
    if bSetCursor and bExpand then
        thisComponent.getCurrentController.Select(oTextCursor)
    end if

    ProcessMovementKey = bMatched
End Function


Sub initVibreoffice
    dim oTextCursor
    ' Initializing
    VIBREOFFICE_STARTED = True
	VIEW_CURSOR = thisComponent.getCurrentController.getViewCursor


	resetMultiplier()
	gotoMode("NORMAL")

    ' Show terminal cursor
    oTextCursor = getTextCursor()
    oTextCursor.goRight(1, False)
    oTextCursor.goLeft(1, True)
	thisComponent.getCurrentController.Select(oTextCursor)

	sStartXKeyHandler()
End Sub


Sub Main
    If Not VIBREOFFICE_STARTED Then
        initVibreoffice()
    End If

    ' Toggle enable/disable
    VIBREOFFICE_ENABLED = Not VIBREOFFICE_ENABLED

    ' Restore statusbar
    If Not VIBREOFFICE_ENABLED Then restoreStatus()
End Sub