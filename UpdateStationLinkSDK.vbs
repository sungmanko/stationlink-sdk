' DLLインストーラーファイル
'

'変数
Dim strCurrentFolder	'カレントフォルダ
Dim objFSO		'FileSystemObject
Dim objWshShell		'WScript.Shell
Dim fPrepareEnd		'準備完了フラグ
Dim strCmdParam		'本スクリプトのコマンドライン
Dim strProgramFilePath	'64Bitのチェック用
Dim strSystemPath	'System32のパス作成
Dim strDllName		'DLL名
Dim fCopyResult		'Copy結果

'変数の初期化
fPrepareEnd = False
fChangeAdmin = True
fCmmandLine = False
strCmdParam = ""
strProgramFilePath = ""
strSystemPath = ""
strDllName = "StationLinkSDK.dll"
fCopyResult = False

'FileSystemObjectの生成
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

'現在のカレントを取得
Set objWshShell = WScript.CreateObject("WScript.Shell")
strCurrentFolder = objWshShell.CurrentDirectory
Set objWshShell = Nothing

'起動パラメータをチェック
For Each strArg In WScript.Arguments
	if strArg = "/PrepareEnd" Then
		fPrepareEnd = True
	Else
		' フォルダが渡されていれば管理者権限実行時にカレントフォルダが渡ってきているので保持
		If objFSO.FolderExists(strArg) = True Then
			strCurrentFolder = strArg
		End If
	End If
Next

'管理者権限を必要とするOSかチェック
Dim objWMI, objShell, osInfo, os
flag = False
On Error Resume Next
Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")  
Set osInfo = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")  
For Each os in osInfo  
	If left(os.Version, 3) >= 6.0 Then 
		flag = true  
	End If
Next
On Error Goto 0

'FileSystemObjectは一度開放
Set objFSO = Nothing

'アップデート処理
'管理者権限で実行する必要があれば、本スクリプトを管理者権限で再実行
If fPrepareEnd = True Or flag = False Then

	'FileSystemObjectの生成
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	'StationLinkのパスを取得
	Set objWshShell = WScript.CreateObject("WScript.Shell")
	strProgramFilePath = objWshShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
	If strProgramFilePath = "%ProgramFiles(x86)%" Or strProgramFilePath = "" Then
		strSystemPath = "C:\WINDOWS\system32\"
	Else
		strSystemPath = "C:\WINDOWS\SysWOW64\"
	End If

	'コピー元ファイルのチェック
	If objFSO.FileExists(strCurrentFolder & "\" & strDllName) = TRUE Then
		'ファイルをコピー
		objFSO.CopyFile strCurrentFolder & "\" & strDllName, strSystemPath & strDllName, True
		If Err.Number = 0 Then
			MsgBox "ファイルのコピーに成功しました。" & vbCrLf & vbCrLf & _
				 "コピー元:" & strCurrentFolder & "\" & strDllName & vbCrLf & _
				 "コピー先:" & strSystemPath & strDllName, _
				vbOKOnly OR vbQuestion OR vbSystemModal, _
				strDllName & "アップデート"
			fCopyResult = True

		Else
			MsgBox "ファイルのコピーに失敗しました。" & vbCrLf & vbCrLf & _
				 "コピー元:" & strCurrentFolder & "\" & strDllName & vbCrLf & _
				 "コピー先:" & strSystemPath & strDllName, _
				vbOKOnly OR vbCritical OR vbSystemModal, _
				strDllName & "アップデート"
			fCopyResult = False
		End If
	Else
		fCopyResult = True
	End If

	'レジストリの登録
	if fCopyResult Then
		objWshShell.run strSystemPath & "regsvr32.exe " & strSystemPath & strDllName
	End If

	'ShellObjectはもう使用しないので開放
	Set objWshShell = Nothing

	'FileSystemObjectはもう使用しないので開放
	Set objFSO = Nothing

Else

	'===============================================
	'======= 本スクリプトを管理者権限で再実行
	'===============================================
	'準備終了をパラメータで渡す
	strCmdParam = strCmdParam & " /PrepareEnd"
	strCmdParam = strCmdParam & " " & strCurrentFolder

	'管理者権限でこのスクリプトを再実行する
	Set objShell=CreateObject("Shell.Application")    
	If flag Then 
		objShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """" & strCmdParam,"","runas",1  
	End If
	Set objShell = Nothing
End If

