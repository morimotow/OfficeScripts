Option Explicit

Dim m_objFso
Set m_objFso = CreateObject("Scripting.FileSystemObject")

' 引数チェック(1つない場合はエラー)
Dim m_oParam
Set m_oParam = WScript.Arguments
If m_oParam.Count <> 1 Then
	Call DisplayUsage()
	Call WScript.Quit()
End If

' 引数チェック(ファイルない場合はエラー)
Dim m_path
m_path = m_oParam(0)
If Not m_objFso.FileExists(m_path) Then
	Call DisplayUsage()
	Call WScript.Quit(-1)
End If

' 主処理呼び出し
Call Main(m_path)

' 使用方法説明表示(引数エラー)
Sub DisplayUsage()
	Dim msg
	msg = "copy_bk.vbs:バックアップファイル作成" & vbCrLf
	msg = msg & "使い方：" & vbCrLf
	msg = msg & "WScript.exe copy_bk.vbs <ファイルパス>" & vbCrLf
	Call MsgBox(msg, ,"エラー")
End Sub

Sub DisplayError(emsg)
	Dim msg
	msg = "エラーが発生しました" & vbCrLf
	msg = msg & msg & vbCrLf
	Call MsgBox(emsg,, "エラー")
End Sub

' 主処理
Sub Main(filePath)

	' ファイルのフルパスを取得
	Dim orgPath
	orgPath = GetFullPath(filePath)

	' ファイルのバックアップを実行
	If Not ExecBackup(orgPath) Then
		Call DisplayError("バックアップの作成に失敗しました")
		Call WScript.Quit(-1)
	End If

	Call WScript.Quit(0)

End Sub

' ファイルのフルパスを取得
Function GetFullPath(path)
	Dim objFile
	Set objFile = m_objFso.GetFile(path)
	Dim result
	result = objFile.Path
	GetFullPath = result
End Function

' バックアップファイル作成
' True - 作成成功[作成しない場合も含む]
' False - 作成失敗(処理終了)
Function ExecBackup(orgPath)

	' 指定ファイルパスをフォルダ、ベース名、拡張子に分解
	Dim orgDir
	orgDir = m_objFso.GetParentFolderName(orgPath)
	Dim orgBase
	orgBase = m_objFso.GetBaseName(orgPath)
	Dim orgExt
	orgExt = m_objFso.GetExtensionName(orgPath)

	' bkフォルダの存在を確認(なければバックアップせずに終了)
	Dim strBkDir
	strBkDir = m_objFso.BuildPath(orgDir, "bk")
	If Not m_objFso.FolderExists(strBkDir) Then
		ExecBackup = False
		Exit Function
	End If

	' バックアップファイル名を作成
	Dim bkPath
	bkPath = CreateBackupPath(orgDir, orgBase, orgExt)

	' バックアップファイル名がない場合は終了
	If Len(bkPath) <= 0 Then
		ExecBackup = False
		Exit Function
	End If

	' ファイルコピー実行
	m_objFso.CopyFile orgPath, bkPath
	Call MsgBox("バックアップファイル名：" & m_objFso.GetFileName(bkPath), , "コピー完了")
	ExecBackup = True

End Function

' バックアップファイルのパスを生成します
Function CreateBackupPath(orgDir, orgBase, orgExt)

	Dim result

	' 今日の日付を文字列に変換します
	Dim nowDate
	nowDate = Replace(Left(Now(),10), "/", "")

	' バックアップがまだない場合は連番なしのファイルパスを返す
	Dim bkNum
	bkNum = ExistsBackupPath(orgDir, orgBase, orgExt, nowDate)
	If bkNum < 0 Then
		result = CreateNewBackupPath(orgDir, orgBase, orgExt, nowDate, 0)
		CreateBackupPath = result
		Exit Function
	End If

	' バックアップがすでに９世代まで作成されている場合は
	' メッセージを表示して終了
	if bkNum >= 9 Then
		Call MsgBox("これ以上バックアップを作成できません。終了します", , "連番が上限")
		result = ""
		CreateBackupPath = result
		Exit Function
	End If

	' すでにバックアップファイルがある場合は確認メッセージ表示
	' (そのまま上書き、連番作成、キャンセル)
	Dim prompt
	prompt = "すでにバックアップファイルがあります。:" & CreateNewBackupPathBase(orgBase, orgExt, nowDate, bkNum) & vbCrLf
	prompt = prompt & "連番を追加して新しいファイルを作成しますか？" & vbCrLf
	prompt = prompt & "操作を選択してください。" & vbCrLf
	prompt = prompt & "　はい：連番を追加してファイルを新規作成" & vbCrLf
	prompt = prompt & "　いいえ：最新のバックアップファイルを上書きして作成" & vbCrLf
	prompt = prompt & "　キャンセル：バックアップファイルを作成しない" & vbCrLf
	Dim msgResult
	msgResult = MsgBox(prompt, vbYesNoCancel, "ファイル重複")

	' キャンセルの場合は空文字を返す
	If msgResult = vbCancel Then
		result = ""
	End If

	' そのまま上書きの場合は見つかったバックアップファイルパスをそのまま返す
	If msgResult = vbNo Then
		result = CreateNewBackupPath(orgDir, orgBase, orgExt, nowDate, bkNum)
	End If

	' 連番作成の場合は連番を１増分したバックアップファイルパスを返す
	If msgResult = vbYes Then
		result = CreateNewBackupPath(orgDir, orgBase, orgExt, nowDate, bkNum + 1)
	End If

	CreateBackupPath = result

End Function

' 連番つきバックアップファイル存在確認
' -1 - バックアップファイルなし
' 0 - バックアップファイル(連番なし)あり
' 1 ～ 9 - バックアップファイル(連番つき)あり
Function ExistsBackupPath(orgDir, orgBase, orgExt, nowDate)

	' 連番なし～連番9まで、バックアップファイルがあるかチェックします
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

	' 見つかった最大の連番を返します
	ExistsBackupPath = bkNum

End Function

' 連番つきバックアップファイル名の作成(<元のファイル>_<yyyymmdd>_<連番>.<元の拡張子>)
Function CreateNewBackupPath(orgDir, orgBase, orgExt, nowDate, bkNum)

	Dim newFileName
	newFileName = CreateNewBackupPathBase(orgBase, orgExt, nowDate, bkNum)

	Dim result
	result = m_objFso.BuildPath(orgDir & "\bk", newFileName)
	CreateNewBackupPath = result

End Function

Function CreateNewBackupPathBase(orgBase, orgExt, nowDate, bkNum)
	' 連番が0以下の場合は連番なしのバックアップファイル名を作成
	Dim newFileName
	If bkNum <= 0 Then
		newFileName = orgBase & "_" & nowDate & "." & orgExt
	Else
		newFileName = orgBase & "_" & nowDate & "_" & CStr(bkNum) & "." & orgExt
	End If

	CreateNewBackupPathBase = newFileName
End Function

