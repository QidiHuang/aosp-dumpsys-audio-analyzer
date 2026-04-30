'----------------------------------------------------
iBaseSlotWidth = 120
iBaseSlotHight = 220
iBaseSlotSpacing = 100

'----------------------------------------------------

Set Wso = WScript.CreateObject("Scripting.WindowSystemObject")

Set Form = Wso.CreateForm(0,0,650,400)
'Form.CenterControl()

' button Demo
Set Button = Form.CreateButton(10,10,75,25,"Demo")
Button.AddEventHandler "OnClick",GetRef("ButtonClick")

' button Close
Set CancelButton = Form.CreateButton(100,10,75,25,"Close")
CancelButton.OnClick = GetRef("CloseFormHandler")

Form.CreateStatusBar().Name = "StatusBar"
Form.StatusBar.Add(100).AutoSize = true

' mouse monitor
Form.OnMouseMove = GetRef("MouseMove")
Form.OnMouseLeave = GetRef("MouseExit")
Form.OnMouseUp = GetRef("MouseUp")
Form.OnMouseDown = GetRef("MouseDown")

' keyboard monitor
Form.OnKeyDown = GetRef("KeyDown")

' create a textbox
Set Text1 = Form.TextOut(10,60,"0")

' set a timer
Set t = Wso.CreateTimer()
t.Interval = 1000
t.OnExecute = GetRef("OnTimer")
t.Active = true


' Use Frame as container of all boxes
'Set AppSlot = Form.createFrame(10, 80, 60, 200)
'AppSlot.BorderWidth = 2

iSlotOneLeft = 10
iSlotOneTop = 80
Form.Bevel iSlotOneLeft, iSlotOneTop, iBaseSlotWidth, iBaseSlotHight

iSlotTwoLeft = iSlotOneLeft+iBaseSlotWidth+iBaseSlotSpacing
iSlotTwoTop = iSlotOneTop
Form.Bevel iSlotTwoLeft, iSlotTwoTop, iBaseSlotWidth, iBaseSlotHight


Set mother = New ClsA
mother.init

Set listener_one = New ClsB
listener_one.setIdNum 111
retCode = mother.registerListener(listener_one)
Set listener_two = New ClsB
listener_two.setIdNum 222
retCode = mother.registerListener(listener_two)

mother.showStorage
mother.finalize


'Form.Pen.Width = 4
'Form.Pen.Style = PS_DOT

'Form.Rectangle 10, 80, 30, 140
'Form.Rectangle 60, 80, 30, 140

'Form.Pen.Width = 2
'Form.Pen.Color = &H000000FF
'Form.Line (10+30), 140/2+80, 60, 140/2+80

'----------------------------------------------------
Form.Show()

Wso.Run()


'----------------------------------------------------
Sub OnTimer(Sender)
    Text1.Text = Text1.Text + 1
End Sub

Sub MouseMove(this,x,y,flags)
	this.Form.StatusBar.item(0).Text = CStr(x)+" x "+CStr(y)
End Sub

Sub MouseExit(this)
	this.Form.StatusBar.item(0).Text = "No Mouse In Control"
End Sub

Sub MouseUp(this,x,y,Button,Flags)
	'this.Form.MessageBox "MouseUp "+CStr(x)+" x "+CStr(y)+", Button = "+CStr(Button)
End Sub

Sub MouseDown(this,x,y,Button,Flags)
	Form.Line 0, 0, x, y
	'Form.Circle x, y, 50
End Sub

Sub KeyDown(this,Key,Flags)
	If ((Key <> 27) And (Key <> 112)) Then this.Form.MessageBox "KeyDown "+CStr(Key)
End Sub

Sub ButtonClick(this)
	this.Form.MessageBox "Button "+this.Text+": OnClick"
End Sub

int CanCloseVar

Sub OKResult(this)
	CanCloseVar = 1
	this.Form.Close()
End Sub

'Sub HelpAbout(Sender)
'	Wso.About()
'End Sub

Function StartupDir()
	Dim s
	s = WScript.ScriptFullName
	s = Left(s,InStrRev(s,"\"))
	StartupDir = s
End Function

Sub WSOHelp(Sender)
	Set shell = WScript.CreateObject("WScript.Shell")
	shell.Run """"+StartupDir() + "..\..\wso.chm"+""""
End Sub

Sub AboutWSO_OnHitTest(Sender,x,y,ResultPtr)	
	ResultPtr.put(Wso.Translate("HTCAPTION"))
End Sub

Sub CloseFormHandler(Sender)
	Sender.Form.Close()
End Sub




'----------------------------------------------------
Class ClsA
	Private mListeners
	Private mFlagClsInited ' = false

	Public Function init()
		Set mListeners = CreateObject("System.Collections.ArrayList")
		mFlagClsInited = true
		init = 0
	End Function
	
	Public Function finalize()
		Set mListeners = Nothing
		mFlagClsInited = false
		finalize = 0
	End Function

	Public Function registerListener(ByRef clsB)
		If mFlagClsInited Then
			mListeners.Add clsB
		Else
			MsgBox "register listener clsB failed!", vbOKOnly+vbExclamation, "Error"
			registerListener = 1
		End If
		registerListener = 0
	End Function
	
	Public Function showStorage()
		cnt = mListeners.Count
		i = 1
		For Each listenerElement in mListeners
			MsgBox "(" & i & "/" & cnt & ") My ID: " & listenerElement.id
			listenerElement.update i
			i = i + 1
		Next
	End Function
End Class

Class ClsB
	Private mIdNum
	
	Public Property Get id()
		id = mIdNum
	End Property

	Public Function setIdNum(ByVal idNum)
		mIdNum = idNum
		setIdNum = mIdNum
	End Function
	
	Public Function update(ByVal strInfo)
		MsgBox "I'm " & mIdNum & ", received feedback for " & strInfo
	End Function
End Class