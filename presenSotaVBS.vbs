'===========================================================
'TCP/IP�ŒʐM����T���v��(VBS/VBA)
'�ʐM�I�u�W�F�N�g����
'===========================================================
'�y���ӎ����z
'�@[regsvr32.exe NONCOMSCK.OCX]���K�v
'�@����VBS�T���v����64bit(x64)��VBS�ł�CreateObject�G���[�ɂȂ�܂��B
'�@32bit(x86)�ł�WSH(C:\Windows\SysWow64\cscript.exe)���g�p���Ă��������B
'===========================================================

'===========================================================
'���C������
'===========================================================
'----------
' ���������^�ݒ�
'----------
Dim ipAddess, portNo
Dim i
Dim commandStr(10), k
Dim commandFileName
Dim pptFileName

ipAddess        = "172.16.168.36"
portNo          = 5001
'ipAddess        = "127.0.0.1"
'portNo          = 4000
commandFileName = "genko.txt"
pptFileName     = "LINCREA.pptx"


'----------
' ����
'----------
Set Winsock1 = CreateObject("NonComSck.Winsock")
i = 0

'�p���|�I�u�W�F�N�g�����E�����ݒ�
Set oApp = CreateObject("PowerPoint.Application")
oApp.Visible = True '���ɂ���
oApp.Presentations.Open(getCurPath & "\" & pptFileName)
oApp.ActivePresentation.SlideShowSettings.Run 'PPT�N��

'�R�}���h�p�������z��ɃZ�b�g
Call readCommandFile(getCurPath & "\" & commandFileName)

'�R�}���h�p������𑗐M���郋�[�v
Do While True
	Call startConnection
	If transData = false Then
		Exit Do
	End If
Loop


'----------
' �I������
'----------
Set oApp = Nothing
Set Winsock1 = Nothing

WSCript.Quit



'===========================================================
'���C������
'===========================================================
Function transData()
	
	Dim wText
	Dim encodeSendStr
	Dim splitAryStr
	Dim pageNo
	Dim waitTime
	
	WScript.Echo "---transData-----"
	'----------
	' �f�[�^���M(�������Byte�z��ɕϊ����đ��M)�^End�̏ꍇ�͋����I��
	'----------
	
	'�R�}���h�p��������P�s���������o
	wText = speechText(i)
	WScript.Echo i & ":" & wText

	'�R�}���h�p������i���s�R�[�h<LF>���j��UTF-8�ɕϊ�����
	encodeSendStr = encodeStr(wText & vbLf, "UTF-8")
	
	'�R�}���h�p��������J���}�ŕ�������
	splitAryStr = Split(wText, ",", -1)
	
	If splitAryStr(0) <> "End" Then
		pageNo = splitAryStr(0)
		WScript.Echo "pageNo:" & pageNo
		waitTime = splitAryStr(3)
		WScript.Echo "waitTime:" & waitTime
	Else
		'�����I��(End)�̏ꍇ�APPT�̍ŏI�y�[�W��\��
		pageNo = "999"
	End If

	'----------
    ' PPT�̃y�[�W����
    '----------
    Call goToPptSlideNo(Int(pageNo), oApp)

	'�T�[�o���փR�}���h�p������𑗐M
	Winsock1.SEndData encodeSendStr

    i = i + 1

	'----------
	' �f�[�^��M�i�T�[�o����̎�M�������m�F�j
	'----------
	Winsock1.Start_EventForScript()
	Do
		WScript.Sleep(500)
		Evt = Winsock1.GetEventParameters()
		If Ubound(Evt) >= 0 Then
		
			' Evt(0) : �C�x���g��
			If Evt(0) = "DataArrival" Then
				' Evt(9) : ��M�f�[�^��Byte�z��
				' Byte�z��𕶎���ɕϊ�
				WScript.Echo Winsock1.ByteArrayToStr(Evt(9))
				Exit Do
				
			End If
			
		End If
	Loop
	Winsock1.End_EventForScript()
	
	'�P�`���̑���M���m�F������ؒf�iTCP/IP�ʐM�̐���j
	Call disConnection()
	
	If splitAryStr(0) = "End" Then
		'�I���R�}���h���ݒ肳��Ă�����A�v���O�����I��
		WSCript.Quit
		transData = false
		Exit Function
	End If
	
	'�������I������܂�Wait����
	WScript.Sleep(waitTime)

	transData = true
	
End Function

'===========================================================
' TCP�ʐM�J�n
'===========================================================
Sub startConnection()
	WScript.Echo "---startConnection-----"
	'----------
	' TCP/IP�ڑ�
	'----------
	Winsock1.Connect ipAddess, portNo

	'----------
	' TCP/IP�ڑ��҂�
	'----------
	Do While Winsock1.State = 6
	    WScript.Sleep(500)
	Loop
End Sub

'===========================================================
' TCP�ʐM�ؒf
'===========================================================
Sub disConnection()
	WScript.Echo "---disconnection-----"
	
	Winsock1.Close2
	
End Sub

'===========================================================
' �R�}���h�p������̒��o�i�P�����j
'===========================================================
Function speechText(Byval pSpeechNo) 
	Dim wRetText
	
	wRetText = commandStr(pSpeechNo)
	speechText = wRetText
	
End Function

'===========================================================
' �R�}���h�p�t�@�C���iUTF-8�̃e�L�X�g�t�@�C���j��Ǎ���
'===========================================================
Sub readCommandFile(Byval pFileName)
	Dim objStream

	'----------
	' �t�@�C����Ǎ���
	'----------
	Set objStream = CreateObject("ADODB.Stream")
	
	objStream.Type = 2							' 1�F�o�C�i��, 2�F�e�L�X�g
	objStream.Charset = "UTF-8"					' �����R�[�h�w��
	objStream.Open
	
	objStream.LoadFromFile pFileName
	
	'----------
	' �Ǎ��݃t�@�C������1�s���R�}���h�p������i�z��j�ɏ�����
	'----------
	k = 0
	Do Until objStream.EOS
		commandStr(k) = objStream.ReadText(-2)	' -1�F�S�s�Ǎ���, -2�F��s�Ǎ���
		'WScript.Echo commandStr(k)
		
		k = k + 1
	Loop
	
	'----------
	'�I������
	'----------
	objStream.Close
	Set objStream = Nothing
	
End Sub

'===========================================================
' �����R�[�h�ϊ�
'===========================================================
Function encodeStr(Byval pStrUni, Byval pCharSet) 

	Set objStream = CreateObject("ADODB.Stream")
	
	'----------
	'�w�肳�ꂽ�������Stream�ɏ�����
	'----------
	objStream.Open
	objStream.Type = 2					' 1�F�o�C�i��, 2�F�e�L�X�g
	objStream.Charset = pCharSet
	objStream.WriteText pStrUni 
	objStream.Position = 0

	'----------
	'�����R�[�h�ϊ�����Strem����ǂݏo��
	'----------
	'BOM�����镶���R�[�h�̏ꍇ�́A�ŏ���BOM�����X�L�b�v
	objStream.Type = 1					' 1�F�o�C�i��, 2�F�e�L�X�g
	Select Case UCase(pCharSet)
		Case "UNICODE", "UTF-16"
			objStream.Position = 2
			
		Case "UTF-8"
			objStream.Position = 3
			
	End Select
	
	encodeStr = objStream.Read()
	
	'----------
	'�I������
	'----------
	objStream.Close
	Set objStream = Nothing
	
End Function

'===========================================================
' �J�����g�p�X���擾
'===========================================================
Function getCurPath()
	Dim wFileObj
	
	Set wFileObj = createObject("Scripting.FileSystemObject")
	getCurPath = wFileObj.GetParentFolderName(WScript.ScriptFullName)
	
End Function

'===========================================================
' PPT�t�@�C���̎w��X���C�h�ԍ��ɑJ�ڂ���
'===========================================================
Sub goToPptSlideNo(ByVal pNo, ByRef pPptObj)
	Dim wMaxPageNo
	
	wMaxPageNo = pPptObj.ActivePresentation.Slides.Count
	
	'�L���ȃy�[�W���łȂ��ꍇ�́A�ŏI�y�[�W��\��
	If Not(IsNumeric(pNo)) Then
		pNo = wMaxPageNo
	End If

	'�ő�y�[�W���𒴂������̏ꍇ�́A�ŏI�y�[�W��\��
	If pNo > wMaxPageNo Then
		pNo = wMaxPageNo
	End If
	
	'�P�y�[�W��菬�����ꍇ�́A�P�y�[�W�ڂ�\��
	If pNo < 1 Then
		pNo = 1
	End If
	
	'�w��X���C�h�ֈړ�
	Call oApp.SlideShowWindows(1).View.GotoSlide(Int(pNo))
	
End Sub
