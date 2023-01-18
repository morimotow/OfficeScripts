Option Explicit

Dim m_appPath
Dim m_filePath
Dim m_readOnly
Dim m_newProcess

' �����擾
If Not GetArgs() Then
	WScript.Quit -1
End If

' Office�C���X�g�[���p�X�擾
If Not GetAppPath() Then
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

' Office�C���X�g�[���p�X�擾
Function GetAppPath()

	' Excel���N���ς݂̏ꍇ��Excel�̋N���p�X��Ԃ�
	On Error Resume Next
	Dim objApp
	Set objApp = Nothing
	Set objApp = GetObject(, "Excel.Application")
	On Error GoTo 0
	If Not (objApp Is Nothing) Then
		m_appPath = objApp.Path & "\"
		GetAppPath = True
		Exit Function
	End If

	' �N�����Ă��Ȃ��΂�����Excel���N�����ċN���p�X�擾��I��������
	Set objApp = CreateObject("Excel.Application")
	m_appPath = objApp.Path & "\"
	objApp.Quit
	Set objApp = Nothing
	GetAppPath = True

End Function

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

	If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xltx" Or ext = "xltm" Then
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

	' �V�v���Z�X�w��ȊO�ŁA�����t�@�C�����J���Ă��邩�`�F�b�N
	Dim answer
	answer = vbNo
	Dim sameBook
	Set sameBook = Nothing
	If (Not newProcess) Then
		Set sameBook = GetSameBookOpen(filePath)
		If Not (sameBook Is Nothing) Then
			answer = MsgBox("�t�@�C���u" & sameBook.Name & "�v�͊��ɊJ���Ă��܂��B�ʂ̃E�C���h�E�ŊJ���܂����H", vbOkCancel, "�t�@�C���̏d��")
		End If
	End If

	' �����t�@�C���`�F�b�N�ŃL�����Z�������Ƃ��͏I��
	If answer = vbCancel Then
		ExcelOpen = False
		Exit Function
	End If
	' �����t�@�C�����J����Ă���ꍇ�͈�������������
	If Not (sameBook Is Nothing) Then
		newProcess = True
		' �����t�@�C�����ǂݎ���p�łȂ��ꍇ�͈�������������
		If Not sameBook.ReadOnly Then
			readOnly = True
		End If
	End If

	Dim cmd
	cmd = """" & m_appPath & "EXCEL.exe"""

	' �V�v���Z�X�w��̏ꍇ�͈����ǉ�
	If newProcess Then
		cmd = cmd & " /x"
	End If

	' �ǂݎ���p�̏ꍇ�͈����ǉ�
	If readOnly Then
		cmd = cmd & " /r"
	End If

	' �w�肳�ꂽ�t�@�C����n��
	cmd = cmd & " """ & filePath & """"

	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	Call objShell.Run(cmd)

	If Err.Number Then
		MsgBox "�G���[���������܂���" & vbCrLf & vbCrLf & Err.Description, , "�G���["
		ExcelOpen = False
	Else
		ExcelOpen = True
	End If

	On Error GoTo 0

End Function

' �w�肵���t�@�C���Ɠ����̃t�@�C������J�����u�b�N��Ԃ�
Function GetSameBookOpen(filePath)

	' Excel������Ȃ��ꍇ�̓t�@�C�����J���Ă��Ȃ�
	On Error Resume Next
	Dim objApp
	Set objApp = Nothing
	Set objApp = GetObject(, "Excel.Application")
	On Error GoTo 0
	If objApp Is Nothing Then
		Set GetSameBookOpen = Nothing
		Exit Function
	End If

	' �t�@�C�����̂ݎ擾
	Dim fileName
	Dim objFs
	Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
	fileName = objFs.GetFileName(filePath)

	' �t�@�C�����Ɠ����̃��[�N�u�b�N��Ԃ�
	On Error Resume Next
	Dim objBook
	Set objBook = Nothing
	Set objBook = objApp.WorkBooks(fileName)
	On Error GoTo 0
	If (objBook Is Nothing) Then
		Set GetSameBookOpen = Nothing
	Else
		Set GetSameBookOpen = objBook
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
