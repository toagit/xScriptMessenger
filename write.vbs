dir =""
name = "IDxxxx"
filePath = dir & "MSG_" & name & ".txt"
interval = 10000

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(filePath, 2, True)

inputText = InputBox("���b�Z�[�W����͂��Ă�������")
objFile.WriteLine(inputText)
WScript.Sleep interval

objFile.Close
Set objFile = Nothing
Set objFSO = Nothing