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

	' ����3�łȂ��ꍇ�G���[
	If oArgs.Count <> 3 Then
		Call DisplayUsage()
		GetArgs = False
		Exit Function
	End If

	' �t�@�C���p�X(�p�X���Ȃ��ꍇ�G���[)
	m_filePath = oArgs(2)
	Dim oFs
	Set oFs = CreateObject("Scripting.FileSystemObject")
	If Not oFs.FileExists(m_filePath) Then
		Call DisplayUsage()
		GetArgs = False
		Exit Function
	End If
	' �t�@�C���p�X���t���p�X�ɕϊ�
	m_filePath = oFs.GetAbsolutePathName(m_filePath)

	' �V���[�g�J�b�g�t�@�C�������t�@�C���p�X�ɕϊ�
	m_filePath = ResolveShortcut(m_filePath)

	' �ǂݎ���p�w��(0:�ʏ�A1:�ǂݎ���p)
	m_readOnly = oArgs(0)
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
	m_newProcess = oArgs(1)
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
	msg = msg & "WScript.exe open_office.vbs <0:�ʏ�|1:�ǎ��p> [<0:�����v���Z�X�g�p|1:�V�v���Z�X�쐬>] <�t�@�C���p�X>" & vbCrLf
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

' �V���[�g�J�b�g�t�@�C���Ŏw�肳��Ă�����t�@�C���p�X��Ԃ�
Function ResolveShortcut(filePath)

	' �g���q����
	Dim ext
	Dim objFs
	Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
	ext = LCase(objFs.GetExtensionName(filePath))

	' �V���[�g�J�b�g�t�@�C���łȂ��ꍇ�͉������܂���
	If ext <> "lnk" Then
		ResolveShortcut = filePath
		Exit Function
	End If

	' �V���[�g�J�b�g�t�@�C����ǂݍ��݂܂�
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	Dim objLnk
	Set objLnk = objShell.CreateShortcut(filePath)

	' �p�X���擾���܂�
	Dim result
	result = objLnk.TargetPath
	ResolveShortcut = result

End Function

Function ExcelOpen(filePath, readOnly, newProcess)

	' �����v���Z�X���w�肳��Ă��āA�N�����Ă���Excel���Ȃ��ꍇ�̓V�F���@�\�Ńt�@�C�����J���܂��B
	If (Not newProcess) And (Not ExistsApp("Excel.Application")) Then
		Call ExcelShellOpen(filePath, readOnly)
		ExcelOpen = True
		Exit Function
	End If

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
		Call objApp.WorkBooks.Add(filePath)
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

Function PowerPointOpen(filePath, readOnly)

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

	' �E�B���h�E�̔�\�����������A�A�N�e�B�u�ɂ���
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	objApp.Visible = True
	Call objShell.AppActivate(objApp.Caption)

	' �E�B���h�E���ŏ�������Ă��鎞�����A0.2�b��ɕ���
	If objApp.WindowState = -4140 Then
		WScript.Sleep 200
		objShell.SendKeys "% r"
	End If
End Sub

' �w�肳�ꂽOffice�A�v���P�[�V�������N���ς݂��`�F�b�N
Function ExistsApp(progId)
	On Error Resume Next
	Dim objApp
	Set objApp = GetObject(, progId)
	If Err.Number Then
		ExistsApp = False
	Else
		Set objApp = Nothing
		ExistsApp = True
	End If
	On Error GoTo 0
End Function

' �w�肳�ꂽ�t�@�C����Excel�ŊJ��(�V�F�����s)
Sub ExcelShellOpen(filePath, readOnly)

	' Excel.Application����A���s�t�@�C���p�X���擾
	Dim objApp
	Set objApp = CreateObject("Excel.Application")
	Dim path
	path = objApp.Path
	Call objApp.Quit()
	Set objApp = Nothing

	' ���Excel�A�v���P�[�V�������V�F���@�\�𗘗p���ċN��
	Dim cmd
	cmd = """" & path & "\EXCEL.EXE"""
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	Call objShell.Run(cmd)

	' �N���ς݂�Excel�I�u�W�F�N�g���擾���A�t�@�C�����J��
	' �A�h�C���N���܂ł��΂炭���Ԃ������邱�Ƃ�����̂őҋ@�����ǉ�
	Dim objApp2
	Set objApp2 = WaitGetObject("Excel.Application")
	If objApp2 Is Nothing Then
		Exit Sub
	End If

	On Error Resume Next
	Call objApp2.WorkBooks.Add(filePath)
	If Err.Number Then
		MsgBox "�G���[���������܂���" & vbCrLf & vbCrLf & Err.Description, , "�G���["
	End If

	On Error GoTo 0
End Sub

' GetObject�Ŏw�肳�ꂽ�A�v���P�[�V�������擾�ł���܂őҋ@
Function WaitGetObject(progId)

	On Error Resume Next

	Dim objApp
	Dim i
	For i = 0 To 9

		' �Q�b�ҋ@
		WScript.Sleep 2000

		' 2�b�ŋN�������ꍇ�͂��̂܂܃A�v���P�[�V�����Q�Ƃ�Ԃ�
		Set objApp = GetObject(, progId)
		If Err.Number = 0 Then
			Set WaitGetObject = objApp
			Exit Function
		End If

		' �N�����Ă��Ȃ��ꍇ(�G���[429)�ȊO�̓G���[�\��
		If Err.Number <> 429 Then
			MsgBox Err.Number & ":" & Err.Description
			Set WaitGetObject = Nothing
			Exit Function
		End If

		Err.Clear

	Next

	' 20�b�o�ߌ�A����ɑҋ@���邩�₢���킹��
	Dim wait_continue
	wait_continue = MsgBox("�A�v���P�[�V�������N�����Ȃ��悤�ł��B�p�����đҋ@���܂����H", vbYesNo + vbSystemModal, "�A�v���P�[�V�����N���ҋ@")
	If wait_continue = vbYes Then
		Set WaitGetObject = WaitGetObject(progId)
	Else
		Set WaitGetObject = Nothing
	End If

End Function
