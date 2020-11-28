Set WshShell = CreateObject("WScript.Shell")

Function convertToKey(Key)
'Uses the Windows Registry; convertToKey(Key) taken from: 
'https://answers.microsoft.com/en-us/insider/forum/insider_wintp-insider_repair/how-to-find-all-windows-version-serial-key/a6d7e4eb-2adf-4e57-8ead-0bd85ec2758d?auth=1 
Const KeyOffset = 52
i = 28
Chars = "BCDFGHJKMPQRTVWXY2346789"
Do
Cur = 0
x = 14
Do
Cur = Cur * 256
Cur = Key(x + KeyOffset) + Cur
Key(x + KeyOffset) = (Cur \ 24) And 255
Cur = Cur Mod 24
x = x -1
Loop While x >= 0
i = i -1
KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
If (((29 - i) Mod 6) = 0) And (i <> -1) Then
i = i -1
KeyOutput = "-" & KeyOutput
End If
Loop While i >= 0
ConvertToKey = KeyOutput
End Function


Function runCommand(command)
runCommand= WshShell.Exec(command).StdOut.ReadLine()
End Function


intAnswer = _
    Msgbox("If you press yes you will be able to view your Windows product key.You will be given the option to export it to a text file later. Press Ctrl-C to copy the key to you clipboard", _
        vbYesNo, "Do you want to view your Window's product key?")

If intAnswer = vbYes Then
     ' Checks to see if the command returns a valid product key, if it doesn’t then the Windows Registry should contain it.
     tester = Replace(runCommand("cmd /c wmic path softwarelicensingservice get OA3xOriginalProductKey")," ","",1,6)
     If Len(tester) < 24 Then ' A Windows product key will always be 25 characters long
      key = tester
     Else 
      key= convertToKey(WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))
     End If
answer = _
    Msgbox("Windows product key: "+ key, _
        vbYesNo, "Do you want to export your Window's product key to a text file?")
    If answer = vbYes Then ' Exports the key to a text file called "Windows_product_key.txt"

      runCommand "cmd /c echo " + key + " > Windows_product_key.txt"
    End If

End if
