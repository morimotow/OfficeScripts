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

	' 既存プロセスが指定されていて、起動しているExcelがない場合はシェル機能でファイルを開きます。
	If (Not newProcess) And (Not ExistsApp("Excel.Application")) Then
		Call ExcelShellOpen(filePath, readOnly)
		ExcelOpen = True
		Exit Function
	End If

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
		Call objApp.WorkBooks.Add(filePath)
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

' 指定されたOfficeアプリケーションが起動済みかチェック
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

' 指定されたファイルをExcelで開く(シェル実行)
Sub ExcelShellOpen(filePath, readOnly)

	' Excel.Applicationから、実行ファイルパスを取得
	Dim objApp
	Set objApp = CreateObject("Excel.Application")
	Dim path
	path = objApp.Path
	Call objApp.Quit()
	Set objApp = Nothing

	' 空のExcelアプリケーションをシェル機能を利用して起動
	Dim cmd
	cmd = """" & path & "\EXCEL.EXE"""
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	Call objShell.Run(cmd)

	' 起動済みのExcelオブジェクトを取得し、ファイルを開く
	' アドイン起動までしばらく時間がかかることがあるので待機処理追加
	Dim objApp2
	Set objApp2 = WaitGetObject("Excel.Application")
	If objApp2 Is Nothing Then
		Exit Sub
	End If

	On Error Resume Next
	Call objApp2.WorkBooks.Add(filePath)
	If Err.Number Then
		MsgBox "エラーが発生しました" & vbCrLf & vbCrLf & Err.Description, , "エラー"
	End If

	On Error GoTo 0
End Sub

' GetObjectで指定されたアプリケーションが取得できるまで待機
Function WaitGetObject(progId)

	On Error Resume Next

	Dim objApp
	Dim i
	For i = 0 To 9

		' ２秒待機
		WScript.Sleep 2000

		' 2秒で起動した場合はそのままアプリケーション参照を返す
		Set objApp = GetObject(, progId)
		If Err.Number = 0 Then
			Set WaitGetObject = objApp
			Exit Function
		End If

		' 起動していない場合(エラー429)以外はエラー表示
		If Err.Number <> 429 Then
			MsgBox Err.Number & ":" & Err.Description
			Set WaitGetObject = Nothing
			Exit Function
		End If

		Err.Clear

	Next

	' 20秒経過後、さらに待機するか問い合わせる
	Dim wait_continue
	wait_continue = MsgBox("アプリケーションが起動しないようです。継続して待機しますか？", vbYesNo + vbSystemModal, "アプリケーション起動待機")
	If wait_continue = vbYes Then
		Set WaitGetObject = WaitGetObject(progId)
	Else
		Set WaitGetObject = Nothing
	End If

End Function
