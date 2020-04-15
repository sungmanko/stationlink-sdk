' DLL�C���X�g�[���[�t�@�C��
'

'�ϐ�
Dim strCurrentFolder	'�J�����g�t�H���_
Dim objFSO		'FileSystemObject
Dim objWshShell		'WScript.Shell
Dim fPrepareEnd		'���������t���O
Dim strCmdParam		'�{�X�N���v�g�̃R�}���h���C��
Dim strProgramFilePath	'64Bit�̃`�F�b�N�p
Dim strSystemPath	'System32�̃p�X�쐬
Dim strDllName		'DLL��
Dim fCopyResult		'Copy����

'�ϐ��̏�����
fPrepareEnd = False
fChangeAdmin = True
fCmmandLine = False
strCmdParam = ""
strProgramFilePath = ""
strSystemPath = ""
strDllName = "StationLinkSDK.dll"
fCopyResult = False

'FileSystemObject�̐���
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

'���݂̃J�����g���擾
Set objWshShell = WScript.CreateObject("WScript.Shell")
strCurrentFolder = objWshShell.CurrentDirectory
Set objWshShell = Nothing

'�N���p�����[�^���`�F�b�N
For Each strArg In WScript.Arguments
	if strArg = "/PrepareEnd" Then
		fPrepareEnd = True
	Else
		' �t�H���_���n����Ă���ΊǗ��Ҍ������s���ɃJ�����g�t�H���_���n���Ă��Ă���̂ŕێ�
		If objFSO.FolderExists(strArg) = True Then
			strCurrentFolder = strArg
		End If
	End If
Next

'�Ǘ��Ҍ�����K�v�Ƃ���OS���`�F�b�N
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

'FileSystemObject�͈�x�J��
Set objFSO = Nothing

'�A�b�v�f�[�g����
'�Ǘ��Ҍ����Ŏ��s����K�v������΁A�{�X�N���v�g���Ǘ��Ҍ����ōĎ��s
If fPrepareEnd = True Or flag = False Then

	'FileSystemObject�̐���
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	'StationLink�̃p�X���擾
	Set objWshShell = WScript.CreateObject("WScript.Shell")
	strProgramFilePath = objWshShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
	If strProgramFilePath = "%ProgramFiles(x86)%" Or strProgramFilePath = "" Then
		strSystemPath = "C:\WINDOWS\system32\"
	Else
		strSystemPath = "C:\WINDOWS\SysWOW64\"
	End If

	'�R�s�[���t�@�C���̃`�F�b�N
	If objFSO.FileExists(strCurrentFolder & "\" & strDllName) = TRUE Then
		'�t�@�C�����R�s�[
		objFSO.CopyFile strCurrentFolder & "\" & strDllName, strSystemPath & strDllName, True
		If Err.Number = 0 Then
			MsgBox "�t�@�C���̃R�s�[�ɐ������܂����B" & vbCrLf & vbCrLf & _
				 "�R�s�[��:" & strCurrentFolder & "\" & strDllName & vbCrLf & _
				 "�R�s�[��:" & strSystemPath & strDllName, _
				vbOKOnly OR vbQuestion OR vbSystemModal, _
				strDllName & "�A�b�v�f�[�g"
			fCopyResult = True

		Else
			MsgBox "�t�@�C���̃R�s�[�Ɏ��s���܂����B" & vbCrLf & vbCrLf & _
				 "�R�s�[��:" & strCurrentFolder & "\" & strDllName & vbCrLf & _
				 "�R�s�[��:" & strSystemPath & strDllName, _
				vbOKOnly OR vbCritical OR vbSystemModal, _
				strDllName & "�A�b�v�f�[�g"
			fCopyResult = False
		End If
	Else
		fCopyResult = True
	End If

	'���W�X�g���̓o�^
	if fCopyResult Then
		objWshShell.run strSystemPath & "regsvr32.exe " & strSystemPath & strDllName
	End If

	'ShellObject�͂����g�p���Ȃ��̂ŊJ��
	Set objWshShell = Nothing

	'FileSystemObject�͂����g�p���Ȃ��̂ŊJ��
	Set objFSO = Nothing

Else

	'===============================================
	'======= �{�X�N���v�g���Ǘ��Ҍ����ōĎ��s
	'===============================================
	'�����I�����p�����[�^�œn��
	strCmdParam = strCmdParam & " /PrepareEnd"
	strCmdParam = strCmdParam & " " & strCurrentFolder

	'�Ǘ��Ҍ����ł��̃X�N���v�g���Ď��s����
	Set objShell=CreateObject("Shell.Application")    
	If flag Then 
		objShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """" & strCmdParam,"","runas",1  
	End If
	Set objShell = Nothing
End If

