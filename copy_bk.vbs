Option Explicit

Dim m_objFso
Set m_objFso = CreateObject("Scripting.FileSystemObject")

' �����`�F�b�N(1�Ȃ��ꍇ�̓G���[)
Dim m_oParam
Set m_oParam = WScript.Arguments
If m_oParam.Count <> 1 Then
	Call DisplayUsage()
	Call WScript.Quit()
End If

' �����`�F�b�N(�t�@�C���Ȃ��ꍇ�̓G���[)
Dim m_path
m_path = m_oParam(0)
If Not m_objFso.FileExists(m_path) Then
	Call DisplayUsage()
	Call WScript.Quit(-1)
End If

' �又���Ăяo��
Call Main(m_path)

' �g�p���@�����\��(�����G���[)
Sub DisplayUsage()
	Dim msg
	msg = "copy_bk.vbs:�o�b�N�A�b�v�t�@�C���쐬" & vbCrLf
	msg = msg & "�g�����F" & vbCrLf
	msg = msg & "WScript.exe copy_bk.vbs <�t�@�C���p�X>" & vbCrLf
	Call MsgBox(msg, ,"�G���[")
End Sub

Sub DisplayError(emsg)
	Dim msg
	msg = "�G���[���������܂���" & vbCrLf
	msg = msg & msg & vbCrLf
	Call MsgBox(emsg,, "�G���[")
End Sub

' �又��
Sub Main(filePath)

	' �t�@�C���̃t���p�X���擾
	Dim orgPath
	orgPath = GetFullPath(filePath)

	' �t�@�C���̃o�b�N�A�b�v�����s
	If Not ExecBackup(orgPath) Then
		Call DisplayError("�o�b�N�A�b�v�̍쐬�Ɏ��s���܂���")
		Call WScript.Quit(-1)
	End If

	Call WScript.Quit(0)

End Sub

' �t�@�C���̃t���p�X���擾
Function GetFullPath(path)
	Dim objFile
	Set objFile = m_objFso.GetFile(path)
	Dim result
	result = objFile.Path
	GetFullPath = result
End Function

' �o�b�N�A�b�v�t�@�C���쐬
' True - �쐬����[�쐬���Ȃ��ꍇ���܂�]
' False - �쐬���s(�����I��)
Function ExecBackup(orgPath)

	' �w��t�@�C���p�X���t�H���_�A�x�[�X���A�g���q�ɕ���
	Dim orgDir
	orgDir = m_objFso.GetParentFolderName(orgPath)
	Dim orgBase
	orgBase = m_objFso.GetBaseName(orgPath)
	Dim orgExt
	orgExt = m_objFso.GetExtensionName(orgPath)

	' bk�t�H���_�̑��݂��m�F(�Ȃ���΃o�b�N�A�b�v�����ɏI��)
	Dim strBkDir
	strBkDir = m_objFso.BuildPath(orgDir, "bk")
	If Not m_objFso.FolderExists(strBkDir) Then
		ExecBackup = False
		Exit Function
	End If

	' �o�b�N�A�b�v�t�@�C�������쐬
	Dim bkPath
	bkPath = CreateBackupPath(orgDir, orgBase, orgExt)

	' �o�b�N�A�b�v�t�@�C�������Ȃ��ꍇ�͏I��
	If Len(bkPath) <= 0 Then
		ExecBackup = False
		Exit Function
	End If

	' �t�@�C���R�s�[���s
	m_objFso.CopyFile orgPath, bkPath
	Call MsgBox("�o�b�N�A�b�v�t�@�C�����F" & m_objFso.GetFileName(bkPath), , "�R�s�[����")
	ExecBackup = True

End Function

' �o�b�N�A�b�v�t�@�C���̃p�X�𐶐����܂�
Function CreateBackupPath(orgDir, orgBase, orgExt)

	Dim result

	' �����̓��t�𕶎���ɕϊ����܂�
	Dim nowDate
	nowDate = Replace(Left(Now(),10), "/", "")

	' �o�b�N�A�b�v���܂��Ȃ��ꍇ�͘A�ԂȂ��̃t�@�C���p�X��Ԃ�
	Dim bkNum
	bkNum = ExistsBackupPath(orgDir, orgBase, orgExt, nowDate)
	If bkNum < 0 Then
		result = CreateNewBackupPath(orgDir, orgBase, orgExt, nowDate, 0)
		CreateBackupPath = result
		Exit Function
	End If

	' �o�b�N�A�b�v�����łɂX����܂ō쐬����Ă���ꍇ��
	' ���b�Z�[�W��\�����ďI��
	if bkNum >= 9 Then
		Call MsgBox("����ȏ�o�b�N�A�b�v���쐬�ł��܂���B�I�����܂�", , "�A�Ԃ����")
		result = ""
		CreateBackupPath = result
		Exit Function
	End If

	' ���łɃo�b�N�A�b�v�t�@�C��������ꍇ�͊m�F���b�Z�[�W�\��
	' (���̂܂܏㏑���A�A�ԍ쐬�A�L�����Z��)
	Dim prompt
	prompt = "���łɃo�b�N�A�b�v�t�@�C��������܂��B:" & CreateNewBackupPathBase(orgBase, orgExt, nowDate, bkNum) & vbCrLf
	prompt = prompt & "�A�Ԃ�ǉ����ĐV�����t�@�C�����쐬���܂����H" & vbCrLf
	prompt = prompt & "�����I�����Ă��������B" & vbCrLf
	prompt = prompt & "�@�͂��F�A�Ԃ�ǉ����ăt�@�C����V�K�쐬" & vbCrLf
	prompt = prompt & "�@�������F�ŐV�̃o�b�N�A�b�v�t�@�C�����㏑�����č쐬" & vbCrLf
	prompt = prompt & "�@�L�����Z���F�o�b�N�A�b�v�t�@�C�����쐬���Ȃ�" & vbCrLf
	Dim msgResult
	msgResult = MsgBox(prompt, vbYesNoCancel, "�t�@�C���d��")

	' �L�����Z���̏ꍇ�͋󕶎���Ԃ�
	If msgResult = vbCancel Then
		result = ""
	End If

	' ���̂܂܏㏑���̏ꍇ�͌��������o�b�N�A�b�v�t�@�C���p�X�����̂܂ܕԂ�
	If msgResult = vbNo Then
		result = CreateNewBackupPath(orgDir, orgBase, orgExt, nowDate, bkNum)
	End If

	' �A�ԍ쐬�̏ꍇ�͘A�Ԃ��P���������o�b�N�A�b�v�t�@�C���p�X��Ԃ�
	If msgResult = vbYes Then
		result = CreateNewBackupPath(orgDir, orgBase, orgExt, nowDate, bkNum + 1)
	End If

	CreateBackupPath = result

End Function

' �A�Ԃ��o�b�N�A�b�v�t�@�C�����݊m�F
' -1 - �o�b�N�A�b�v�t�@�C���Ȃ�
' 0 - �o�b�N�A�b�v�t�@�C��(�A�ԂȂ�)����
' 1 �` 9 - �o�b�N�A�b�v�t�@�C��(�A�Ԃ�)����
Function ExistsBackupPath(orgDir, orgBase, orgExt, nowDate)

	' �A�ԂȂ��`�A��9�܂ŁA�o�b�N�A�b�v�t�@�C�������邩�`�F�b�N���܂�
	Dim bkNum
	bkNum = -1
	Dim i
	Dim chkPath
	For i = 0 To 9
		chkPath = CreateNewBackupPath(orgDir, orgBase, orgExt, nowDate, i)
		If m_objFso.FileExists(chkPath) Then
			bkNum = i
		End If
	Next

	' ���������ő�̘A�Ԃ�Ԃ��܂�
	ExistsBackupPath = bkNum

End Function

' �A�Ԃ��o�b�N�A�b�v�t�@�C�����̍쐬(<���̃t�@�C��>_<yyyymmdd>_<�A��>.<���̊g���q>)
Function CreateNewBackupPath(orgDir, orgBase, orgExt, nowDate, bkNum)

	Dim newFileName
	newFileName = CreateNewBackupPathBase(orgBase, orgExt, nowDate, bkNum)

	Dim result
	result = m_objFso.BuildPath(orgDir & "\bk", newFileName)
	CreateNewBackupPath = result

End Function

Function CreateNewBackupPathBase(orgBase, orgExt, nowDate, bkNum)
	' �A�Ԃ�0�ȉ��̏ꍇ�͘A�ԂȂ��̃o�b�N�A�b�v�t�@�C�������쐬
	Dim newFileName
	If bkNum <= 0 Then
		newFileName = orgBase & "_" & nowDate & "." & orgExt
	Else
		newFileName = orgBase & "_" & nowDate & "_" & CStr(bkNum) & "." & orgExt
	End If

	CreateNewBackupPathBase = newFileName
End Function

