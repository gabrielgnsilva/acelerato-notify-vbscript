Dim result, args, pscommand, tickets, fs, ignoredTickets, ignore, executor, shell, file, found, data, outputFile, cmd, scriptPath, currentDir, config, configFile, url, login, token

Set fs = CreateObject("Scripting.FileSystemObject")
scriptPath = WScript.ScriptFullName
currentDir = Left(scriptPath, InStrRev(scriptPath, "\"))

Set configFile = fs.OpenTextFile(currentDir & "\.config", 1, False, -1)
config = Split(configFile.ReadAll(), vbCrLf)
configFile.Close

url = config(0)
login = config(1)
token = config(2)

Set file = fs.OpenTextFile(currentDir & "\ignoreList.txt", 1, False, -1)
ignoredTickets = Split(file.ReadAll(), vbCrLf)
file.Close

cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle hidden -File " & currentDir & "\getTickets.ps1 " & url & " " & login & " " & token & " " & currentDir & "\output.txt"
Set shell = CreateObject("WScript.Shell")
shell.Run cmd, 0, True

Set outputFile  = fs.OpenTextFile(currentDir & "\output.txt", 1, False, -1)
tickets = Split(outputFile.ReadAll(), ",")
outputFile.Close
For Each ticket In tickets

    continue = True
    If ticket = "-" + vbCrLf Then continue = False End If
    If continue Then
        data = Split(ticket, "ZGXT")
        
        found = False
        For Each ignoredTicket In ignoredTickets
            If Trim(ignoredTicket) = data(0) Then
                found = True
                Exit For
            End If
        Next
        
        If Not found Then
            result = MsgBox("A new ticket has been issued, would you like to open it?" & vbCrLf & vbCrLf & "Ticket: " & data(0) & vbCrLf & vbCrLf & "Title: " & data(1) & vbCrLf & vbCrLf & "Reporter: " & data(2), vbInformation + vbYesNo, "New support ticket!")
            Select Case result
                Case vbYes
                    shell.Run data(3)
                Case vbNo
                    ignore = MsgBox("Would you like to ignore this ticket?" & vbCrLf & vbCrLf & "Ticket: " & data(0) & vbCrLf & vbCrLf & "Title: " & data(1) & vbCrLf & vbCrLf & "Reporter: " & data(2), vbExclamation + vbYesNo, "Ignore")
                    Select Case ignore
                        Case vbYes
                            Set file = fs.OpenTextFile(currentDir & "\ignoreList.txt", 8, True, -1)
                            file.WriteLine(CStr(data(0)))
                            file.Close
                    End Select
            End Select
        End If
    End If
Next

' Cleanup
Set result = Nothing
Set args = Nothing
Set pscommand = Nothing
Set tickets = Nothing
Set fs = Nothing
Set ignoredTickets = Nothing
Set ignore = Nothing
Set executor = Nothing
Set shell = Nothing
Set file = Nothing
Set found = Nothing
Set data = Nothing
Set outputFile = Nothing
Set cmd = Nothing
Set scriptPath = Nothing
Set currentDir = Nothing
Set configFile = Nothing
Set configFilePath = Nothing
Set url = Nothing
Set login = Nothing
Set token = Nothing
