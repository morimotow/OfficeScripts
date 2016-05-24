Option Explicit

Dim m_filePath
Dim m_readOnly
Dim m_newProcess

' 引数取得
If Not GetArgs() Then
	WScript.Quit -1
End If

' 主処理呼び出し
If Not Main(m_filePath, m_readOnly, m_newProcess) Then
	WScript.Quit -1
End If

' -- 引数取得 ---
Function GetArgs()

	Dim oArgs
	Set oArgs = WScript.Arguments

	' 引数2つ以上でない場合エラー
	If oArgs.Count < 2 Then
		Call DisplayUsage()
		GetArgs = False
		Exit Function
	End If

	' ファイルパス(パスがない場合エラー)
	m_filePath = oArgs(0)
	Dim oFs
	Set oFs = CreateObject("Scripting.FileSystemObject")
	If Not oFs.FileExists(m_filePath) Then
		Call DisplayUsage()
		GetArgs = False
		Exit Function
	End If
	' ファイルパスをフルパスに変換
	m_filePath = oFs.GetAbsolutePathName(m_filePath)

	' 読み取り専用指定(0:通常、1:読み取り専用)
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

	' 別プロセス指定(0:既存のプロセスあれば使用、1:新規作成)
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

' -- 使用方法表示 --
Sub DisplayUsage()
	Dim msg
	msg = "Officeファイル操作補助スクリプト" & vbCrLf
	msg = msg & "使い方：" & vbCrLf
	msg = msg & "WScript.exe open_office.vbs <ファイルパス> <0:通常|1:読取専用> [<0:既存プロセス使用|1:新プロセス作成>]" & vbCrLf
	MsgBox msg, , "Officeファイル操作補助スクリプト"
End Sub

' 読み取り専用
' 読み取り専用(別ウィンドウ)
' 編集
' 編集(別ウインドウ)

' 主処理
Function Main(filePath, readOnly, newProcess)

	Dim objApp

	' 拡張子判定
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
		Call MsgBox("拡張子「" & ext & "」は処理対象外です",, "エラー")
		Main = False
	End If

End Function

Function ExcelOpen(filePath, readOnly, newProcess)

	' 新しいプロセスが指定されている場合はCreateObject
	Dim objApp
	Set objApp = GetOfficeApp("Excel.Application", newProcess)

	' 同名ファイルが開いているかチェック
	Dim answer
	answer = vbNo
	If IsSameBookOpen(filePath, objApp) Then
		Dim fileName
		Dim objFs
		Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
		fileName = objFs.GetFileName(filePath)
		answer = MsgBox("ファイル「" & fileName & "」は既に開いています。別のウインドウで開きますか？", vbOkCancel, "ファイルの重複")
	End If

	' 同名ファイルチェックでキャンセルしたときは終了
	If answer = vbCancel Then
		ExcelOpen = False
		Exit Function
	End If
	' 同名ファイルチェックで別プロセス指定したときは別Excelで開き直す
	If answer = vbOk Then
		Set objApp = GetOfficeApp("Excel.Application", True)
	End If

	' アプリケーションを前面に表示
	Call SetAppFocus(objApp)

	' ファイルを開く
	On Error Resume Next
	If readOnly Then
		Call objApp.WorkBooks.Open(filePath, , True)
	Else
		Call objApp.WorkBooks.Open(filePath)
	End If

	If Err.Number Then
		MsgBox "エラーが発生しました" & vbCrLf & vbCrLf & Err.Description, , "エラー"
		ExcelOpen = False
	Else
		ExcelOpen = True
	End If

	On Error GoTo 0

End Function

' 同名のファイルが開いているかチェック
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

	' Wordは必ず既存のアプリケーションで開くようにします
	Dim objApp
	Set objApp = GetOfficeApp("Word.Application", False)

	' アプリケーションを前面に表示
	Call SetAppFocus(objApp)

	' ファイルを開く
	If readOnly Then
		Call objApp.Documents.Open(filePath, , True)
	Else
		Call objApp.Documents.Open(filePath)
	End If

	If Err.Number Then
		MsgBox "エラーが発生しました" & vbCrLf & vbCrLf & Err.Description, , "エラー"
		WordOpen = False
	Else
		WordOpen = True
	End If

	On Error GoTo 0

End Function

Function PowerPointOpen(filePath, readOnly, newProcess)

	' PowerPointは必ず既存のアプリケーションで開くようにします
	Dim objApp
	Set objApp = GetOfficeApp("Powerpoint.Application", False)

	' アプリケーションを前面に表示
	Call SetAppFocus(objApp)

	' ファイルを開く
	If readOnly Then
		Call objApp.Presentations.Open(filePath, True)
	Else
		Call objApp.Presentations.Open(filePath)
	End If

	If Err.Number Then
		MsgBox "エラーが発生しました" & vbCrLf & vbCrLf & Err.Description, , "エラー"
		PowerPointOpen = False
	Else
		PowerPointOpen = True
	End If

	On Error GoTo 0

End Function

' Officeのアプリケーションオブジェクトを取得
Function GetOfficeApp(progId, newProcess)

	On Error Resume Next
	Dim objApp

	' 新しいプロセスが指定されている場合はCreateObject
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

' アプリケーションのウインドウを表示して前面に表示
Sub SetAppFocus(objApp)
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	objApp.Visible = True
	Call objShell.AppActivate(objApp.Caption)
	WScript.Sleep 100
	objShell.SendKeys "% r"
End Sub
