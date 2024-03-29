Option Explicit

Dim m_appPath
Dim m_filePath
Dim m_readOnly
Dim m_newProcess

' 引数取得
If Not GetArgs() Then
	WScript.Quit -1
End If

' Officeインストールパス取得
If Not GetAppPath() Then
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

	' 引数3つでない場合エラー
	If oArgs.Count <> 3 Then
		Call DisplayUsage()
		GetArgs = False
		Exit Function
	End If

	' ファイルパス(パスがない場合エラー)
	m_filePath = oArgs(2)
	Dim oFs
	Set oFs = CreateObject("Scripting.FileSystemObject")
	If Not oFs.FileExists(m_filePath) Then
		Call DisplayUsage()
		GetArgs = False
		Exit Function
	End If
	' ファイルパスをフルパスに変換
	m_filePath = oFs.GetAbsolutePathName(m_filePath)

	' ショートカットファイルを実ファイルパスに変換
	m_filePath = ResolveShortcut(m_filePath)

	' 読み取り専用指定(0:通常、1:読み取り専用)
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

	' 別プロセス指定(0:既存のプロセスあれば使用、1:新規作成)
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

' -- 使用方法表示 --
Sub DisplayUsage()
	Dim msg
	msg = "Officeファイル操作補助スクリプト" & vbCrLf
	msg = msg & "使い方：" & vbCrLf
	msg = msg & "WScript.exe open_office.vbs <0:通常|1:読取専用> [<0:既存プロセス使用|1:新プロセス作成>] <ファイルパス>" & vbCrLf
	MsgBox msg, , "Officeファイル操作補助スクリプト"
End Sub

' Officeインストールパス取得
Function GetAppPath()

	' Excelが起動済みの場合はExcelの起動パスを返す
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

	' 起動していないばあいはExcelを起動して起動パス取得後終了させる
	Set objApp = CreateObject("Excel.Application")
	m_appPath = objApp.Path & "\"
	objApp.Quit
	Set objApp = Nothing
	GetAppPath = True

End Function

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

	If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xltx" Or ext = "xltm" Then
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

' ショートカットファイルで指定されている実ファイルパスを返す
Function ResolveShortcut(filePath)

	' 拡張子判定
	Dim ext
	Dim objFs
	Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
	ext = LCase(objFs.GetExtensionName(filePath))

	' ショートカットファイルでない場合は何もしません
	If ext <> "lnk" Then
		ResolveShortcut = filePath
		Exit Function
	End If

	' ショートカットファイルを読み込みます
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	Dim objLnk
	Set objLnk = objShell.CreateShortcut(filePath)

	' パスを取得します
	Dim result
	result = objLnk.TargetPath
	ResolveShortcut = result

End Function

Function ExcelOpen(filePath, readOnly, newProcess)

	' 新プロセス指定以外で、同名ファイルが開いているかチェック
	Dim answer
	answer = vbNo
	Dim sameBook
	Set sameBook = Nothing
	If (Not newProcess) Then
		Set sameBook = GetSameBookOpen(filePath)
		If Not (sameBook Is Nothing) Then
			answer = MsgBox("ファイル「" & sameBook.Name & "」は既に開いています。別のウインドウで開きますか？", vbOkCancel, "ファイルの重複")
		End If
	End If

	' 同名ファイルチェックでキャンセルしたときは終了
	If answer = vbCancel Then
		ExcelOpen = False
		Exit Function
	End If
	' 同名ファイルが開かれている場合は引数を書き換え
	If Not (sameBook Is Nothing) Then
		newProcess = True
		' 同名ファイルが読み取り専用でない場合は引数を書き換え
		If Not sameBook.ReadOnly Then
			readOnly = True
		End If
	End If

	Dim cmd
	cmd = """" & m_appPath & "EXCEL.exe"""

	' 新プロセス指定の場合は引数追加
	If newProcess Then
		cmd = cmd & " /x"
	End If

	' 読み取り専用の場合は引数追加
	If readOnly Then
		cmd = cmd & " /r"
	End If

	' 指定されたファイルを渡す
	cmd = cmd & " """ & filePath & """"

	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	Call objShell.Run(cmd)

	If Err.Number Then
		MsgBox "エラーが発生しました" & vbCrLf & vbCrLf & Err.Description, , "エラー"
		ExcelOpen = False
	Else
		ExcelOpen = True
	End If

	On Error GoTo 0

End Function

' 指定したファイルと同名のファイルから開いたブックを返す
Function GetSameBookOpen(filePath)

	' Excelが一つもない場合はファイルを開いていない
	On Error Resume Next
	Dim objApp
	Set objApp = Nothing
	Set objApp = GetObject(, "Excel.Application")
	On Error GoTo 0
	If objApp Is Nothing Then
		Set GetSameBookOpen = Nothing
		Exit Function
	End If

	' ファイル名のみ取得
	Dim fileName
	Dim objFs
	Set objFs = WScript.CreateObject("Scripting.FileSystemObject")
	fileName = objFs.GetFileName(filePath)

	' ファイル名と同名のワークブックを返す
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

Function PowerPointOpen(filePath, readOnly)

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

	' ウィンドウの非表示を解除し、アクティブにする
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	objApp.Visible = True
	Call objShell.AppActivate(objApp.Caption)

	' ウィンドウが最小化されている時だけ、0.2秒後に復元
	If objApp.WindowState = -4140 Then
		WScript.Sleep 200
		objShell.SendKeys "% r"
	End If
End Sub
