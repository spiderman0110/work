This is a sample project students can use during Matthew's Git class.

Here is an addition by me

We can have a bit of fun with this repo, knowing that we can always reset it to a known good state.  We can apply labels, and branch, then add new code and merge it in to the master branch.

As a quick reminder, this came from one of three locations in either SSH, Git, or HTTPS format:

* git@github.com:matthewmccullough/hellogitworld.git
* git://github.com/matthewmccullough/hellogitworld.git
* https://matthewmccullough@github.com/matthewmccullough/hellogitworld.git

We can, as an example effort, even modify this README and change it as if it were source code for the purposes of the class.

This demo also includes an image with changes on a branch for examination of image diff on GitHub.


Word Kill Process

Option Explicit

'スクリプト名称（ダイアログに表示）
Const SCRIPT_NAME = "WordのプロセスKill"

If MsgBox("Wordの全てのプロセスをKillします。" & vbLf & "【注意】開いているDocは全て保存せずにクローズします。", vbOKCancel + vbExclamation, SCRIPT_NAME) = vbOK Then
	If killProcess Then
		MsgBox "完了！", vbOkOnly, SCRIPT_NAME
	End If
End If

WScript.Quit

'■主処理
Private Function killProcess()
	Call terminateProcess("WINWORD.EXE")
	killProcess = True
End Function

'■スクリプトの起動を検知
Private Sub terminateProcess(processName)
	Dim process, count, ret
	For Each process In GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_Process where Name='" & processName & "'")
		process.Terminate
	Next
End Sub

Excel Kill Process

Option Explicit

'スクリプト名称（ダイアログに表示）
Const SCRIPT_NAME = "ExcelのプロセスKill"

If MsgBox("Excelの全てのプロセスをKillします。" & vbLf & "【注意】開いているブックは全て保存せずにクローズします。", vbOKCancel + vbExclamation, SCRIPT_NAME) = vbOK Then
	If killProcess Then
		MsgBox "完了！", vbOkOnly, SCRIPT_NAME
	End If
End If

WScript.Quit

'■主処理
Private Function killProcess()
	Call terminateProcess("EXCEL.EXE")
	killProcess = True
End Function

'■スクリプトの起動を検知
Private Sub terminateProcess(processName)
	Dim process, count, ret
	For Each process In GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_Process where Name='" & processName & "'")
		process.Terminate
	Next
End Sub
