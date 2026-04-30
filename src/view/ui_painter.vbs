'=============================================================================
'  This is the View component, responsible for painting UI elements.
'
'=============================================================================



'----------------------------------------------------
' register listener to dumpsys_worker_thread for
' receiving parsed dump info.
'----------------------------------------------------
Const iBaseSlotWidth	= 120
Const iBaseSlotHight	= 220
Const iBaseSlotSpacing	= 100
Const DIRECTION_NONE	= 0
Const DIRECTION_LEFT	= 1
Const DIRECTION_top		= 2
Const DIRECTION_RIGHT	= 4
Const DIRECTION_BOTTOM	= 8
Const DIRECTION_INNER	= 16

Class DumpsysView
	Private mWso
	Private mIdText
	Private mFlagClsInited ' = false
	Private mSlotAPP, mSlotAF, mSlotHAL, mSlotALSA
	Private mAppBlocks, mThreadBlocks, mTrackBlocks, mHalBufferBlocks, mDeviceBlocks
	Private mLastAppLink, mLastTrackLink, mLastBufferLink, mRoutes
	Private mForm

	Public Property Get getIdText()
		getIdText = mIdText
	End Property

	Public Function init(ByRef wso)
		Set mWso = wso
		Set mAppBlocks = CreateObject("System.Collections.ArrayList")
		Set mThreadBlocks = CreateObject("System.Collections.ArrayList")
		Set mTrackBlocks = CreateObject("System.Collections.ArrayList")
		Set mHalBufferBlocks = CreateObject("System.Collections.ArrayList")
		Set mDeviceBlocks = CreateObject("System.Collections.ArrayList")
		Set mRoutes = CreateObject("System.Collections.ArrayList")
		Set mForm = mWso.CreateForm(200,100,850,560)
		mForm.Text = "Dumpsys Analyzer"

		iSlotAPPLeft = 10
		iSlotAPPTop = 20
		Set mSlotAPP = new SlotPanel
		mSlotAPP.initSlotPanel mForm, iSlotAPPLeft, iSlotAPPTop, iBaseSlotWidth, iBaseSlotHight
		mSlotAPP.setText "APPs"

		iSlotAFLeft = iSlotAPPLeft + iBaseSlotWidth + iBaseSlotSpacing
		iSlotAFTop = iSlotAPPTop
		Set mSlotAF = new SlotPanel
		mSlotAF.initSlotPanel mForm, iSlotAFLeft, iSlotAFTop, iBaseSlotWidth, iBaseSlotHight
		mSlotAF.setText "AudioFlinger"

		iSlotAHALLeft = iSlotAFLeft + iBaseSlotWidth + iBaseSlotSpacing
		iSlotAHALTop = iSlotAPPTop
		Set mSlotHAL = new SlotPanel
		mSlotHAL.initSlotPanel mForm, iSlotAHALLeft, iSlotAHALTop, iBaseSlotWidth, iBaseSlotHight
		mSlotHAL.setText "AHAL"

		iSlotALSALeft = iSlotAHALLeft + iBaseSlotWidth + iBaseSlotSpacing
		iSlotALSATop = iSlotAPPTop
		Set mSlotALSA = new SlotPanel
		mSlotALSA.initSlotPanel mForm, iSlotALSALeft, iSlotALSATop, iBaseSlotWidth, iBaseSlotHight
		mSlotALSA.setText "ALSA Nodes"

		mFlagClsInited = true
		'----------------------------------------------------
		' Show UI and start running script
		If mFlagClsInited Then
			mForm.Show()
		End If

		init = 0
	End Function

	Public Function paintUI()
		'----------------------------------------------------
		' Add more UI elements below


		'----------------------------
		' paint lines for linking
		'----------------------------
'		Set appConnection1 = new Linking
'		appConnection1.connectSlotPanel mForm, appBlock1, trackAFBlock2, true
'		Set trackConnection1 = new Linking
'		trackConnection1.connectSlotPanel mForm, trackAFBlock1, trackBlock1, true
'		trackConnection1.setLinkingText "AudioTrack"
'		Set trackConnection2 = new Linking
'		trackConnection2.connectSlotPanel mForm, trackBlock4, trackAFBlock2, true
'		trackConnection2.setLinkingText "AudioRecord"
'		Set deviceConnection1 = new Linking
'		deviceConnection1.connectSlotPanel mForm, trackBlock4, deviceBlock1, true

		'trackConnection2.finalize
		'trackBlock3.finalize
'		trackConnection1.finalize



		paintUI = 0
	End Function

	Public Function setIdText(ByVal idText)
		mIdText = idText
		setIdText = mIdText
	End Function

	Public Function clearRoute(ByRef trackBlockArray)
		For Each trackBlockElement in trackBlockArray
			' trackBlockElement is a SlotPanel class instance
			For Each routeElement in mRoutes
				' each routeElement consists of Array(appLink, trackLink, bufferLink)
				If routeElement(0).LinkingContains(trackBlockElement) Or routeElement(1).LinkingContains(trackBlockElement) Then
					WScript.Echo "clear linking 1..." & TypeName(routeElement(0))
					routeElement(0).finalize
					Set routeElement(0) = Nothing
					WScript.Echo "clear linking 2..."& TypeName(routeElement(1))
					routeElement(1).finalize
					Set routeElement(1) = Nothing
					WScript.Echo "clear linking 3..."& TypeName(routeElement(2))
					routeElement(2).finalize
					Set routeElement(2) = Nothing
					WScript.Echo "linkings cleared"
					'mRoutes.Remove(routeElement)
				End If
			Next
		Next
	End Function

	Public Function findSlotPanelForId(ByRef slotPanelArray, ByVal strId, ByRef findingPanel)
		For Each slotPanelElement in slotPanelArray
			elementId = slotPanelElement.getIdentity()
			If elementId = strId Then
				'MsgBox "legacy block found! text is: " & slotPanelElement.getText()
				Set findingPanel = slotPanelElement
				Exit Function
			End If
		Next
		Set findingPanel = Nothing
	End Function

	' strInfo syntax is "TAG_APP|app client id, parent track id=TAG_THREAD|thread id=TAG_TRACK|track id, parent thread id, frames written=TAG_BUFFER|buffer size=TAG_DEVICE|device name, device status"
	Public Function updateInfo(ByVal strInfo)
		'MsgBox "View: " & strInfo
		bFlagNewApp = false : bFlagNewThread = false : bFlagNewTrack = false : bFlagNewBuffer = false : bFlagNewDevice = false
		bFlagActiveApp = false : bFlagActiveThread = false : bFlagActiveTrack = false
		sFlagStateThread = ""
		bFlagNewRoute = false

		infoArray = Split(strInfo, "=")
		For Each infoElement in infoArray
			'MsgBox infoElement
			elementContentArray = Split(infoElement, "|")
			Select Case elementContentArray(0)
				Case "TAG_APP"
					' "TAG_APP|" & trackDetailedInfoArray(3) & "," & processedPsInfoString & "," & trackDetailedInfoArray(2)
					appBlockTextArray = Split(elementContentArray(1), ",")
					If appBlockTextArray(2) = "yes" Then
						bFlagActiveApp = true
					End If
					If mAppBlocks.Count = 0 Then
						Set appBlock = new SlotPanel
						appBlock.attachToSlotPanel mForm, mSlotAPP, DIRECTION_INNER, false
						mAppBlocks.Add appBlock
						appBlock.setText appBlockTextArray(0)&","&appBlockTextArray(1)
						appBlock.setIdentity appBlockTextArray(0)&","&appBlockTextArray(1)
						appBlock.setColor vbGreen
						bFlagNewApp = true
					Else
						Set appBlockDuplicated = new SlotPanel
						findSlotPanelForId mAppBlocks, appBlockTextArray(0)&","&appBlockTextArray(1), appBlockDuplicated
						'MsgBox "Found app block: " & appBlockDuplicated.getText()
						If TypeName(appBlockDuplicated) = "Nothing" Then
							'MsgBox "duplicated app block not found"
							Set appBlockRecent = mAppBlocks.Item(mAppBlocks.Count-1)
							Set appBlock = new SlotPanel
							appBlock.attachToSlotPanel mForm, appBlockRecent, DIRECTION_BOTTOM, true
							mAppBlocks.Add appBlock
							appBlock.setText appBlockTextArray(0)&","&appBlockTextArray(1)
							appBlock.setIdentity appBlockTextArray(0)&","&appBlockTextArray(1)
							appBlock.setColor vbGreen
							bFlagNewApp = true
						Else
							'MsgBox "using duplicated app block"
							Set appBlock = appBlockDuplicated
						End If
					End If


				Case "TAG_THREAD"
					' "=TAG_THREAD|" & threadInfoArray(i)(0) & "," & threadStateString
					threadBlockTextArray = Split(elementContentArray(1), ",")
					'If threadBlockTextArray(1) = "Running" Then
					'	bFlagActiveThread = true
					'End If
					sFlagStateThread = threadBlockTextArray(1)
					If mThreadBlocks.Count = 0 Then
						'If InStr(elementContentArray(1), "Running") > 0 Then
							Set threadBlock = new SlotPanel
							threadBlock.attachToSlotPanel mForm, mSlotAF, DIRECTION_INNER, false
							mThreadBlocks.Add threadBlock
							threadBlock.setHight threadBlock.getHight()*2
							threadBlock.setText threadBlockTextArray(0)
							threadBlock.setIdentity threadBlockTextArray(0)
							threadBlock.setColor &HFF00FF ' Magenta
							bFlagNewThread = true
						'End If
					Else
						'MsgBox elementContentArray(1)
						'threadElementContentArray = Split(elementContentArray(1), ",")
						'elementContentArray(1) = Split(elementContentArray(1), ",", UBound(threadElementContentArray)-1)
						'MsgBox elementContentArray(1)
						Set threadBlockDuplicated = new SlotPanel
						findSlotPanelForId mThreadBlocks, threadBlockTextArray(0), threadBlockDuplicated
						'MsgBox "Found thread block: " & threadBlockDuplicated.getText()
						If TypeName(threadBlockDuplicated) = "Nothing" Then
							Set threadBlockRecent = mThreadBlocks.Item(mThreadBlocks.Count-1)
							Set threadBlock = new SlotPanel
							threadBlock.attachToSlotPanel mForm, threadBlockRecent, DIRECTION_BOTTOM, true
							mThreadBlocks.Add threadBlock
							threadBlock.setText threadBlockTextArray(0)
							threadBlock.setIdentity threadBlockTextArray(0)
							threadBlock.setColor &HFF00FF ' Magenta
							bFlagNewThread = true
						Else
							Set threadBlock = threadBlockDuplicated
						End If
					End If


				Case "TAG_TRACK"
					' "=TAG_TRACK|" & parentThreadTypeId & "," & trackId & "," & trackType
					trackBlockTextArray = Split(elementContentArray(1), ",")
					trackBlockText = trackBlockTextArray(0) & "," & trackBlockTextArray(1)
					If mTrackBlocks.Count = 0 Then
						Set trackBlock = new SlotPanel
						' TODO: currently parent thread block is fixed, optimize it to choose thread block dynamically
						'Set parentThreadBlock = mThreadBlocks.Item(2)
						Set parentThreadBlock = new SlotPanel
						findSlotPanelForId mThreadBlocks, trackBlockTextArray(0), parentThreadBlock
						If TypeName(parentThreadBlock) = "Nothing" Then
							'WScript.Echo "you are not expected to see this message (1/2)."
						Else
							'WScript.Echo "(1/2) found thread " & parentThreadBlock.getIdentity() & ", for track (" & trackBlockTextArray(1) & ", " & trackBlockTextArray(2) & ")"
						End If
						trackBlock.attachToSlotPanel mForm, parentThreadBlock, DIRECTION_INNER, false
						parentThreadBlock.addToContainer trackBlock
						mTrackBlocks.Add trackBlock
						trackBlock.setText trackBlockTextArray(1) & "," & trackBlockTextArray(2)
						trackBlock.setIdentity trackBlockText
						trackBlock.setColor vbBlue
						bFlagNewTrack = true
					Else
						Set trackBlockDuplicated = new SlotPanel
						findSlotPanelForId mTrackBlocks, trackBlockText, trackBlockDuplicated
						If TypeName(trackBlockDuplicated) <> "SlotPanel" Then
							' TODO: new track is drawn based on the latest track block. Optimize it to be drawn in parent thread block
							'Set trackBlockRecent = mTrackBlocks.Item(mTrackBlocks.Count-1)
							Set parentThreadBlock = new SlotPanel
							findSlotPanelForId mThreadBlocks, trackBlockTextArray(0), parentThreadBlock
							If TypeName(parentThreadBlock) = "Nothing" Then
								'WScript.Echo "you are not expected to see this message (2/2)."
							Else
								'WScript.Echo "(2/2) found thread " & parentThreadBlock.getIdentity() & ", for track (" & trackBlockTextArray(1) & ", " & trackBlockTextArray(2) & ")"
							End If
							Set trackBlock = new SlotPanel
							' TODO: new track block is fixed to be drawn inner parent thread block. Optimize it to drawn based on if it is the first track block of a thread
							Set containedTrackBlock = new SlotPanel
							'WScript.Echo "before searching..."
							parentThreadBlock.getLatestContainedElement containedTrackBlock
							WScript.Echo "got containedTrackBlock type: " & TypeName(containedTrackBlock)
							If TypeName(containedTrackBlock) <> "SlotPanel" Then
								' brand new track block in parent thread
								WScript.Echo "find nothing for (" & trackBlockTextArray(1) & ", " & trackBlockTextArray(2) & ")"
								trackBlock.attachToSlotPanel mForm, parentThreadBlock, DIRECTION_INNER, false
							Else
								' one more track block in parent thread
								WScript.Echo "contained track block found: " & containedTrackBlock.getIdentity() & ", for (" & trackBlockTextArray(1) & ", " & trackBlockTextArray(2) & ")"
								trackBlock.attachToSlotPanel mForm, containedTrackBlock, DIRECTION_BOTTOM, true
							End If
							parentThreadBlock.addToContainer trackBlock
							mTrackBlocks.Add trackBlock
							trackBlock.setText trackBlockTextArray(1) & "," & trackBlockTextArray(2)
							trackBlock.setIdentity trackBlockText
							trackBlock.setColor vbBlue
							bFlagNewTrack = true
						Else
							Set trackBlock = trackBlockDuplicated
						End If
					End If


				Case "TAG_BUFFER"
					' "=TAG_BUFFER|" & ahalAddressString & "," & framesSizeString & "*" & framesCountString & "bytes"
					If mHalBufferBlocks.Count = 0 Then
						Set bufferBlock = new SlotPanel
						bufferBlock.attachToSlotPanel mForm, mSlotHAL, DIRECTION_INNER, false
						mHalBufferBlocks.Add bufferBlock
						bufferBlock.setText elementContentArray(1)
						bufferBlock.setIdentity elementContentArray(1)
						bufferBlock.setColor vbYellow
						bFlagNewBuffer = true
					Else
						Set bufferBlockDuplicated = new SlotPanel
						findSlotPanelForId mHalBufferBlocks, elementContentArray(1), bufferBlockDuplicated
						'MsgBox "Found buffer block: " & bufferBlockDuplicated.getText()
						If TypeName(bufferBlockDuplicated) = "Nothing" Then
							Set bufferBlockRecent = mHalBufferBlocks.Item(mHalBufferBlocks.Count-1)
							Set bufferBlock = new SlotPanel
							bufferBlock.attachToSlotPanel mForm, bufferBlockRecent, DIRECTION_BOTTOM, true
							mHalBufferBlocks.Add bufferBlock
							bufferBlock.setText elementContentArray(1)
							bufferBlock.setIdentity elementContentArray(1)
							bufferBlock.setColor vbYellow
							bFlagNewBuffer = true
						Else
							Set bufferBlock = bufferBlockDuplicated
						End If
					End If


				Case "TAG_DEVICE"
					' "=TAG_DEVICE|pcmC0D0p, running"
					If mDeviceBlocks.Count = 0 Then
						Set deviceBlock = new SlotPanel
						deviceBlock.attachToSlotPanel mForm, mSlotALSA, DIRECTION_INNER, false
						mDeviceBlocks.Add deviceBlock
						deviceBlock.setText elementContentArray(1)
						deviceBlock.setIdentity elementContentArray(1)
						deviceBlock.setColor vbRed
						bFlagNewDevice = true
					Else
						Set deviceBlockDuplicated = new SlotPanel
						findSlotPanelForId mDeviceBlocks, elementContentArray(1), deviceBlockDuplicated
						'MsgBox "Found device block: " & deviceBlockDuplicated.getText()
						If TypeName(deviceBlockDuplicated) = "Nothing" Then
							Set deviceBlockRecent = mDeviceBlocks.Item(mDeviceBlocks.Count-1)
							Set deviceBlock = new SlotPanel
							deviceBlock.attachToSlotPanel mForm, deviceBlockRecent, DIRECTION_BOTTOM, true
							mDeviceBlocks.Add deviceBlock
							deviceBlock.setText elementContentArray(1)
							deviceBlock.setIdentity elementContentArray(1)
							deviceBlock.setColor vbRed
							bFlagNewDevice = true
						Else
							Set deviceBlock = deviceBlockDuplicated
						End If
					End If

				Case Else
					MsgBox "unknown tag: " & chr(34) & elementContentArray(0) & chr(34)
			End Select
		Next
		' a complete piece of updated info is processed till now

		Select Case sFlagStateThread
			Case "Running"
				' do nothing
			Case "Idle"
				' do nothing
			Case "IdleCompletely"
				' remove all track blocks belong to the thread
				'WScript.Echo "clean tracks"
				Set containedTrackArray = CreateObject("System.Collections.ArrayList")
				threadBlock.getContainedElements containedTrackArray
				If containedTrackArray.Count > 0 Then
					WScript.Echo "start clearing routes"
					clearRoute containedTrackArray
					WScript.Echo "start clearing container"
					threadBlock.clearContainer
				End If
				Set containedTrackArray = Nothing
				'WScript.Echo "clean work done"

			Case "RunningCompletely"
				Set containedTrackArray = CreateObject("System.Collections.ArrayList")
				threadBlock.getContainedElements containedTrackArray
				If containedTrackArray.Count > 0 Then
					clearRoute containedTrackArray
					threadBlock.clearContainer
				End If
				Set containedTrackArray = Nothing
				
			Case Else
				MsgBox "unknown state: " & sFlagStateThread & ", for thread: "
		End Select

		'MsgBox "mThreadBlocks count: " & mThreadBlocks.Count
		prevColor = 0
		prevLineWidth = 3
		'If bFlagNewApp Or bFlagNewThread Or bFlagNewTrack Then
		If bFlagNewApp Or bFlagNewTrack Then
			Set appLink = new Linking
			appLink.connectSlotPanel mForm, appBlock, trackBlock, true, prevColor, prevLineWidth
			prevColor = appLink.getColor()
			bFlagNewRoute = true
			Set mLastAppLink = appLink
		'Else
		'	Set appLink = mLastAppLink
		End If
		If bFlagNewTrack Or bFlagNewBuffer Then
			Set trackLink = new Linking
			trackLink.connectSlotPanel mForm, trackBlock, bufferBlock, true, prevColor, prevLineWidth
			prevColor = trackLink.getColor()
			bFlagNewRoute = true
			Set mLastTrackLink = trackLink
		'Else
		'	Set trackLink = mLastTrackLink
		End If
		If bFlagNewBuffer Or bFlagNewDevice Then
			Set bufferLink = new Linking
			bufferLink.connectSlotPanel mForm, bufferBlock, deviceBlock, true, prevColor, prevLineWidth
			prevColor = bufferLink.getColor()
			bFlagNewRoute = true
			Set mLastBufferLink = bufferLink
		'Else
		'	bufferLink = mLastBufferLink
		End If
		If bFlagNewRoute Then
			'MsgBox "new route established: APP " & appLink.getNameSlotFrom()
			mRoutes.Add Array(appLink, trackLink, bufferLink)
		End If

		updateInfo = 0
	End Function
End Class


Class SlotPanel
	Private mParentForm, mParentElement, mThis, mContainedElements
	Private mLeft, mTop, mWidth, mHight
	Private mSlotPadding, mInnerElementPadding
	Private mColor, mText, mIdentity
	Private mInited

	Public Property Get getLeft()
		getLeft = mLeft
	End Property

	Public Property Get getTop()
		getTop = mTop
	End Property

	Public Property Get getWidth()
		getWidth = mWidth
	End Property

	Public Function setWidth(ByVal valW)
		mWidth = valW
		'mThis.Width = mWidth
		mThis.SetBounds mLeft, mTop, mWidth, mHight
		setWidth = mWidth
	End Function

	Public Property Get getHight()
		getHight = mHight
	End Property

	Public Function setHight(ByVal valH)
		mHight = valH
		'mThis.Hight = mHight
		mThis.SetBounds mLeft, mTop, mWidth, mHight
		setHight = mHight
	End Function

	Public Function getForm()
		Set getForm = mParentForm
	End Function

	Public Function getParent(ByRef saveTo)
		Set saveTo = mParentElement
		getParent = 0
	End Function

	Public Function setText(ByVal strText)
		mText.Text = strText
	End Function

	Public Function getText()
		getText = mText.Text
	End Function

	Public Function setIdentity(ByVal strText)
		mIdentity = strText
	End Function

	Public Function getIdentity()
		getIdentity = mIdentity
	End Function

	Public Function setColor(ByVal colorVal)
		mColor = colorVal
		mThis.Color = colorVal
	End Function

	Public Function getContainedElements(ByRef saveTo)
		Set saveTo = mContainedElements
		getContainedElements = 0
	End Function

	Public Function addToContainer(ByRef elementToAdd)
		mContainedElements.Add elementToAdd
	End Function
	
	Public Function clearContainer()
		If mContainedElements.Count > 0 Then
			'mContainedElements.Clear
			iMaxElement = mContainedElements.Count-1
			For iContainedElement=0 to iMaxElement
				mContainedElements.RemoveAt(iContainedElement)
			Next
		End If
		WScript.Echo "clean done. mContainedElements.Count is: " & mContainedElements.Count & ", >>>>>> " & getIdentity()
		clearContainer = 0
	End Function

	Public Function getLatestContainedElement(ByRef saveTo)
		If TypeName(mContainedElements) = "Nothing" Then
			WScript.Echo getIdentity() & "has no child block."
			Set saveTo = Nothing
		Else
			WScript.Echo "start searching >>>"
			If mContainedElements.Count > 0 Then
				WScript.Echo mContainedElements.Count & " Element(s) detected..."
				Set saveTo = mContainedElements.Item(mContainedElements.Count - 1)
				WScript.Echo "Load element done: " & mContainedElements.Count - 1
				getLatestContainedElement = 0
				Exit Function
			End If
		End If
		Set saveTo = Nothing
		getLatestContainedElement = 0
	End Function

	Public Function initSlotPanel(ByRef formHandle, ByVal panelLeft, ByVal panelTop, ByVal panelWidth, ByVal panelHight)
		Set mParentForm = formHandle
		Set mContainedElements = CreateObject("System.Collections.ArrayList")
		mLeft = panelLeft : mTop = panelTop : mWidth = panelWidth : mHight = panelHight
		mSlotPadding = 100 : mInnerElementPadding = panelWidth * 0.2
		mColor = &HFFFFFF ' White
		show
		initSlotPanel = 0
	End Function

	Public Function attachToSlotPanel(ByRef formHandle, ByRef parentElementCls, ByVal direction, ByVal isCopy)
		mLeft = 0 : mTop = 0 : mWidth = 0 : mHight = 0
		mSlotPadding = 100

		Set mParentElement = parentElementCls
		''WScript.Echo "init parent element to: " & mParentElement.getText()
		If TypeName(mContainedElements) = "Nothing" Or TypeName(mContainedElements) = "Empty" Or TypeName(mContainedElements) = "Null" Then
			Set mContainedElements = CreateObject("System.Collections.ArrayList")
		End If

'		If isCopy Then
'			Set parentParentElement = parentElementCls.getParent()
'			mInnerElementPadding = parentParentElement.getWidth() * 0.05
'		Else
			mInnerElementPadding = parentElementCls.getWidth() * 0.09
'		End If
		mColor = &HFFFFFF ' White

		Select Case direction
			Case DIRECTION_INNER
				mLeft = parentElementCls.getLeft()+mInnerElementPadding
				mTop = parentElementCls.getTop()+mInnerElementPadding
				mWidth = parentElementCls.getWidth()-mInnerElementPadding*2
				mHight = parentElementCls.getHight()/5
			Case DIRECTION_BOTTOM
				If Not isCopy Then
					mLeft = parentElementCls.getLeft()+mInnerElementPadding
					mTop = parentElementCls.getTop()+parentElementCls.getHight()+mInnerElementPadding
					mWidth = parentElementCls.getWidth()-mInnerElementPadding*2
					mHight = parentElementCls.getHight()/5
				Else
					mLeft = parentElementCls.getLeft()
					mTop = parentElementCls.getTop()+parentElementCls.getHight()+mInnerElementPadding
					mWidth = parentElementCls.getWidth()
					mHight = parentElementCls.getHight()
				End If
			Case DIRECTION_LEFT
				If Not isCopy Then
					mLeft = parentElementCls.getLeft()-mInnerElementPadding-(parentElementCls.getWidth()-mInnerElementPadding*2)
					mTop = parentElementCls.getTop()+mInnerElementPadding
					mWidth = parentElementCls.getWidth()-mInnerElementPadding*2
					mHight = parentElementCls.getHight()/5
				Else
					mLeft = parentElementCls.getLeft()-mInnerElementPadding-parentElementCls.getWidth()
					mTop = parentElementCls.getTop()
					mWidth = parentElementCls.getWidth()
					mHight = parentElementCls.getHight()
				End If
			Case DIRECTION_TOP
				If Not isCopy Then
					mLeft = parentElementCls.getLeft()+mInnerElementPadding
					mTop = parentElementCls.getTop()-mInnerElementPadding-(parentElementCls.getHight()/5)
					mWidth = parentElementCls.getWidth()-mInnerElementPadding*2
					mHight = parentElementCls.getHight()/5
				Else
					mLeft = parentElementCls.getLeft()
					mTop = parentElementCls.getTop()-mInnerElementPadding-parentElementCls.getHight()
					mWidth = parentElementCls.getWidth()
					mHight = parentElementCls.getHight()
				End If
			Case DIRECTION_RIGHT
				If Not isCopy Then
					mLeft = parentElementCls.getLeft()+parentElementCls.getWidth()+mInnerElementPadding
					mTop = parentElementCls.getTop()+mInnerElementPadding
					mWidth = parentElementCls.getWidth()-mInnerElementPadding*2
					mHight = parentElementCls.getHight()/5
				Else
					mLeft = parentElementCls.getLeft()+parentElementCls.getWidth()+mInnerElementPadding
					mTop = parentElementCls.getTop()
					mWidth = parentElementCls.getWidth()
					mHight = parentElementCls.getHight()
				End If
		End Select

		Set mParentForm = formHandle
		show
		'setColor &HFF00FF ' Magenta
		attachToSlotPanel = 0
	End Function

	Public Function show()
		Set mThis = mParentForm.Bevel(mLeft, mTop, mWidth, mHight)
		Set mText = mParentForm.TextOut(mLeft+2, mTop, "")
		mThis.Color = mColor
		show = 0
	End Function

	Public Function finalize()
		setText ""
		mThis.Hide
		mThis.Destroy
		Set mThis = Nothing
		Set mContainedElements = Nothing

		'Set mParentForm = Nothing
		'Set mParentElement = Nothing
		'Set mLeft = Nothing
		'Set mTop = Nothing
		'Set mWidth = Nothing
		'Set mHight = Nothing
		'Set mSlotPadding = Nothing
		'Set mInnerElementPadding = Nothing
		'Set mColor = Nothing
		'Set mText = Nothing
		'Set mIdentity = Nothing
		'Set mInited = Nothing
	End Function
End Class


Class Linking
	Private mParentForm
	Private mSlotFrom, mSlotTo
	Private mSegmentLines
	Private mHorizontalMid, mVerticalMid, mLinkingDirectionHorizontal, mLinkingDirectionVertical
	Private mColor, mLineWidth

	Public Function setLinkingText(ByVal strText)
		Set segmentLine = mSegmentLines.Item(0)
		segmentLine.setText strText
	End Function

	Public Function getNameSlotFrom()
		getNameSlotFrom = mSlotFrom.getText()
	End Function

	Public Function getNameSlotTo()
		getNameSlotTo = mSlotTo.getText()
	End Function

	Public Function setColor(ByVal linkingColor)
		mColor = linkingColor
	End Function

	Public Function getColor()
		getColor = mColor
	End Function

	Public Function getLineWidth()
		getLineWidth = mLineWidth
	End Function

	Public Function LinkingContains(ByRef slotPanelNode)
		If mSlotFrom.getIdentity() = slotPanelNode.getIdentity() Or mSlotTo.getIdentity() = slotPanelNode.getIdentity() Then
			LinkingContains = true
		Else
			LinkingContains = false
		End If
	End Function

	Public Function connectSlotPanel(ByRef formHandle, ByRef slotFrom, ByRef slotTo, ByVal isHorizontalLink, ByVal dupColor, ByVal lineWidth)
		Set mParentForm = formHandle
		Set mSlotFrom = slotFrom
		Set mSlotTo = slotTo
		If dupColor = 0 Then
			setColor RGB(RND()*255, RND()*255, RND()*255)
		Else
			setColor dupColor
		End If
		mLineWidth = lineWidth
		mLinkingDirection = DIRECTION_NONE
		Set mSegmentLines = CreateObject("System.Collections.ArrayList")
		deltaHorizontal = (slotTo.getLeft()-slotFrom.getLeft())/2
		If deltaHorizontal = 0 Then
			' two slots are in the same column
			If isHorizontalLink Then
				MsgBox "two slots are in the same column"
				Exit Function
			End If
		Else
			If deltaHorizontal > 0 Then
				' slotTo is at the right side of slotFrom
				mLinkingDirectionHorizontal = DIRECTION_RIGHT
			Else
				' slotTo is at the left side of slotFrom
				mLinkingDirectionHorizontal = DIRECTION_LEFT
			End If
		End If

		deltaVertical = (slotTo.getTop()-slotFrom.getTop())/2
		If deltaVertical = 0 Then
			' two slots are in the same line
			If Not isHorizontalLink Then
				MsgBox "two slots are in the same line"
				Exit Function
			End If
		Else
			If deltaVertical > 0 Then
				' slotTo is at down side of slotFrom
				mLinkingDirectionVertical = DIRECTION_BOTTOM
			Else
				' slotTo is at up side of slotFrom
				mLinkingDirectionVertical = DIRECTION_TOP
			End If
		End If
		'MsgBox "painting line... mLinkingDirectionHorizontal=" & mLinkingDirectionHorizontal
		If isHorizontalLink Then
			If mLinkingDirectionHorizontal = DIRECTION_LEFT Then
				mHorizontalMid = (slotFrom.getLeft() - (slotTo.getLeft()+slotTo.getWidth()))/2
				Set segmentFirst = new LinkingSegmentLine
				segmentFirst.setSegmentLine mParentForm, slotFrom.getLeft(), slotFrom.getTop()+slotFrom.getHight()/2, slotFrom.getLeft()-mHorizontalMid, slotFrom.getTop()+slotFrom.getHight()/2
				mSegmentLines.Add segmentFirst
				segmentFirst.show getColor(),getLineWidth()
				Set segmentSecond = new LinkingSegmentLine
				segmentSecond.setSegmentLine mParentForm, slotFrom.getLeft()-mHorizontalMid, slotFrom.getTop()+slotFrom.getHight()/2, slotTo.getLeft()+slotTo.getWidth()+mHorizontalMid, slotTo.getTop()+slotTo.getHight()/2
				mSegmentLines.Add segmentSecond
				segmentSecond.show getColor(),getLineWidth()
				Set segmentThird = new LinkingSegmentLine
				segmentThird.setSegmentLine mParentForm, slotTo.getLeft()+slotTo.getWidth()+mHorizontalMid, slotTo.getTop()+slotTo.getHight()/2, slotTo.getLeft()+slotTo.getWidth(), slotTo.getTop()+slotTo.getHight()/2
				mSegmentLines.Add segmentThird
				segmentThird.show getColor(),getLineWidth()
			ElseIf mLinkingDirectionHorizontal = DIRECTION_RIGHT Then
				mHorizontalMid = (slotTo.getLeft() - (slotFrom.getLeft()+slotFrom.getWidth()))/2
				Set segmentFirst = new LinkingSegmentLine
				segmentFirst.setSegmentLine mParentForm, slotFrom.getLeft()+slotFrom.getWidth(), slotFrom.getTop()+slotFrom.getHight()/2, slotFrom.getLeft()+slotFrom.getWidth()+mHorizontalMid, slotFrom.getTop()+slotFrom.getHight()/2
				mSegmentLines.Add segmentFirst
				segmentFirst.show getColor(),getLineWidth()
				Set segmentSecond = new LinkingSegmentLine
				segmentSecond.setSegmentLine mParentForm, slotFrom.getLeft()+slotFrom.getWidth()+mHorizontalMid, slotFrom.getTop()+slotFrom.getHight()/2, slotFrom.getLeft()+slotFrom.getWidth()+mHorizontalMid, slotTo.getTop()+slotTo.getHight()/2
				mSegmentLines.Add segmentSecond
				segmentSecond.show getColor(),getLineWidth()
				Set segmentThird = new LinkingSegmentLine
				segmentThird.setSegmentLine mParentForm, slotFrom.getLeft()+slotFrom.getWidth()+mHorizontalMid, slotTo.getTop()+slotTo.getHight()/2, slotTo.getLeft(), slotTo.getTop()+slotTo.getHight()/2
				mSegmentLines.Add segmentThird
				segmentThird.show getColor(),getLineWidth()
			End If
		Else
			If mLinkingDirectionVertical = DIRECTION_TOP Then
				mVerticalMid = (slotFrom.getTop() - (slotTo.getTop()+slot.getHight()))/2
				' TODO: vertical linking logic - up front
			ElseIf mLinkingDirectionVertical = DIRECTION_BOTTOM Then
				mVerticalMid = (slotTo.getTop() - (slotFrom.getTop()+slotFrom.getHight()))/2
				' TODO: vertical linking logic - down below
			End If
		End If
	End Function

	Public Function finalize()
		For Each segLine in mSegmentLines
			segLine.finalize
		Next
		Set mSegmentLines = Nothing
		mSlotFrom.finalize
		mSlotTo.finalize

		'Set mParentForm = Nothing
		'Set mSlotFrom = Nothing
		'Set mSlotTo = Nothing
		'Set mHorizontalMid = Nothing
		'Set mVerticalMid = Nothing
		'Set mLinkingDirectionHorizontal = Nothing
		'Set mLinkingDirectionVertical = Nothing
		'Set mColor = Nothing
		'Set mLineWidth = Nothing
	End Function
End Class

Class LinkingSegmentLine
	Private mThis
	Private mParentForm
	Private mAX, mAY, mBX, mBY
	Private mText, mTextX, mTextY
	Private mColor

	Public Function setSegmentLine(ByRef formHandle, ByVal axpos, ByVal aypos, ByVal bxpos, ByVal bypos)
		Set mParentForm = formHandle
		mAX = axpos : mAY = aypos : mBX = bxpos : mBY = bypos
		If mAX > mBX Then
			mTextX = mBX+2 : mTextY = mBY
		Else
			mTextX = mAX+2 : mTextY = mAY
		End If
		Set mText = mParentForm.TextOut(mTextX, mTextY, "")
	End Function

	Public Function setText(ByVal strText)
		mText.Text = strText
	End Function

	Public Function setColor(ByVal valColor)
		mThis.Color = valColor
	End Function

	Public Function setLineWidth(ByVal valWidth)
		mThis.Pen.Width = valWidth
	End Function

	Public Function show(ByVal lineColor, ByVal lineWidth)
		Set mThis = mParentForm.Line(mAX, mAY, mBX, mBY)
		setColor lineColor
		setLineWidth lineWidth
	End Function

	Public Function finalize()
		setText ""
		mThis.Hide
		mThis.Destroy
		Set mThis = Nothing
	End Function
End Class

