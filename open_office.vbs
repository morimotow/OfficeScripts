Option Explicit

Dim m_filePath
Dim m_readOnly
Dim m_newProcess

' �����擾
If Not GetArgs() Then
	WScript.Quit -1
End If

' �又���Ăяo��
If Not Main(m_filePath, m_readOnly, m_newProcess) Then
	WScript.Quit -1
End If

' -- �����擾 ---
Function GetArgs()

	Dim oArgs
	Set oArgs = WScript.Arguments

	' ����2�ȏ�łȂ��ꍇ�G���[
	If oArgs.Count < 2 Then
		Call DisplayUsage()
		GetArgs = False
		Exit Function
	End If

	' �t�@�C���p�X(�p�X���Ȃ��ꍇ�G���[)
	m_filePath = oArgs(0)
	Dim oFs
	Set oFs = CreateObject("Scripting.FileSystemObject")
	If Not oFs.FileExists(m_filePath) Then
		Call DisplayUsage()
		GetArgs = False
		Exit Function
	End If
	' �t�@�C���p�X���t���p�X�ɕϊ�
	m_filePath = oFs.GetAbsolutePathName(m_filePath)

	' �ǂݎ���p�w��(0:�ʏ�A1:�ǂݎ���p)
	m_readOnly = oArgs(1)
	If m_readOnly = "0" Then
		m_readOnly = False
	ElseIf m_readOnly = "1" Then
		m_readOnly = True
	Else
		Call DisplayUsage()
		GetArgs = False
		Exit Function
	End If

	' �ʃv���Z�X�w��(0:�����̃v���Z�X����Ύg�p�A1:�V�K�쐬)
	If oArgs.Count = 2 Then
		m_newProcess = "0"
	Else
		m_newProcess = "0"
		m_newProcess = oArgs(2)
	End If
	If m_newProcess = "0" Then
		m_newProcess = CBool(False)
	ElseIf m_newProcess = "1" Then
		m_newProcess = True
	Else
		Call DisplayUsage()
		GetArgs = False
		Exit Function
	End If
	GetArgs = True

End Function

' -- �g�p���@�\�� --
Sub DisplayUsage()
	Dim msg
	msg = "Office�t�@�C������⏕�X�N���v�g" & vbCrLf
	msg = msg & "�g�����F" & vbCrLf
	msg = msg & "WScript.exe open_office.vbs <�t�@�C���p�X> <0:�ʏ�|1:�ǎ��p> [<0:�����v���Z�X�g�p|1:�V�v���Z�X�쐬>]" & vbCrLf
	MsgBox msg, , "Office�t�@�C������⏕�X�N���v�g"
End Sub

' �ǂݎ���p
' �ǂݎ���p(�ʃE�B���h�E)
' �ҏW
' �ҏW(�ʃE�C���h�E)

' �又��
Function Main(filePath, readOnly, newProcess)

	Dim objApp

	' �g���q����
	Dim ext
	Dim objFs
	Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
	ext = LCase(objFs.GetExtensionName(filePath))

	If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
		Main = ExcelOpen(filePath, readOnly, newProcess)
	ElseIf ext = "doc" Or ext = "docx" Then
		Main = WordOpen(filePath, readOnly)
	ElseIf ext = "ppt" Or ext = "pptx" Then
		Main = PowerPointOpen(filePath, readOnly)
	Else
		Call MsgBox("�g���q�u" & ext & "�v�͏����ΏۊO�ł�",, "�G���[")
		Main = False
	End If

End Function

Function ExcelOpen(filePath, readOnly, newProcess)

	' �V�����v���Z�X���w�肳��Ă���ꍇ��CreateObject
	Dim objApp
	Set objApp = GetOfficeApp("Excel.Application", newProcess)

	' �����t�@�C�����J���Ă��邩�`�F�b�N
	Dim answer
	answer = vbNo
	If IsSameBookOpen(filePath, objApp) Then
		Dim fileName
		Dim objFs
		Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
		fileName = objFs.GetFileName(filePath)
		answer = MsgBox("�t�@�C���u" & fileName & "�v�͊��ɊJ���Ă��܂��B�ʂ̃E�C���h�E�ŊJ���܂����H", vbOkCancel, "�t�@�C���̏d��")
	End If

	' �����t�@�C���`�F�b�N�ŃL�����Z�������Ƃ��͏I��
	If answer = vbCancel Then
		ExcelOpen = False
		Exit Function
	End If
	' �����t�@�C���`�F�b�N�ŕʃv���Z�X�w�肵���Ƃ��͕�Excel�ŊJ������
	If answer = vbOk Then
		Set objApp = GetOfficeApp("Excel.Application", True)
	End If

	' �A�v���P�[�V������O�ʂɕ\��
	Call SetAppFocus(objApp)

	' �t�@�C�����J��
	On Error Resume Next
	If readOnly Then
		Call objApp.WorkBooks.Open(filePath, , True)
	Else
		Call objApp.WorkBooks.Open(filePath)
	End If

	If Err.Number Then
		MsgBox "�G���[���������܂���" & vbCrLf & vbCrLf & Err.Description, , "�G���["
		ExcelOpen = False
	Else
		ExcelOpen = True
	End If

	On Error GoTo 0

End Function

' �����̃t�@�C�����J���Ă��邩�`�F�b�N
Function IsSameBookOpen(filePath, objApp)

	Dim fileName
	Dim objFs
	Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
	fileName = objFs.GetFileName(filePath)

	On Error Resume Next
	Dim objBook
	Set objBook = objApp.WorkBooks(fileName)
	On Error GoTo 0
	If Not IsEmpty(objBook) Then
		IsSameBookOpen = True
	Else
		IsSameBookOpen = False
	End If

End Function

Function WordOpen(filePath, readOnly)

	' Word�͕K�������̃A�v���P�[�V�����ŊJ���悤�ɂ��܂�
	Dim objApp
	Set objApp = GetOfficeApp("Word.Application", False)

	' �A�v���P�[�V������O�ʂɕ\��
	Call SetAppFocus(objApp)

	' �t�@�C�����J��
	If readOnly Then
		Call objApp.Documents.Open(filePath, , True)
	Else
		Call objApp.Documents.Open(filePath)
	End If

	If Err.Number Then
		MsgBox "�G���[���������܂���" & vbCrLf & vbCrLf & Err.Description, , "�G���["
		WordOpen = False
	Else
		WordOpen = True
	End If

	On Error GoTo 0

End Function

Function PowerPointOpen(filePath, readOnly, newProcess)

	' PowerPoint�͕K�������̃A�v���P�[�V�����ŊJ���悤�ɂ��܂�
	Dim objApp
	Set objApp = GetOfficeApp("Powerpoint.Application", False)

	' �A�v���P�[�V������O�ʂɕ\��
	Call SetAppFocus(objApp)

	' �t�@�C�����J��
	If readOnly Then
		Call objApp.Presentations.Open(filePath, True)
	Else
		Call objApp.Presentations.Open(filePath)
	End If

	If Err.Number Then
		MsgBox "�G���[���������܂���" & vbCrLf & vbCrLf & Err.Description, , "�G���["
		PowerPointOpen = False
	Else
		PowerPointOpen = True
	End If

	On Error GoTo 0

End Function

' Office�̃A�v���P�[�V�����I�u�W�F�N�g���擾
Function GetOfficeApp(progId, newProcess)

	On Error Resume Next
	Dim objApp

	' �V�����v���Z�X���w�肳��Ă���ꍇ��CreateObject
	If newProcess Then
		Set objApp = CreateObject(progId)
	Else
		Set objApp = GetObject(, progId)
	End If
	If Err.Number Then
		Set objApp = CreateObject(progId)
	End If

	On Error GoTo 0
	Set GetOfficeApp = objApp

End Function

' �A�v���P�[�V�����̃E�C���h�E��\�����đO�ʂɕ\��
Sub SetAppFocus(objApp)
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	objApp.Visible = True
	Call objShell.AppActivate(objApp.Caption)
	WScript.Sleep 100
	objShell.SendKeys "% r"
End Sub
