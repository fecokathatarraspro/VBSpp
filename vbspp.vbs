' ---- VBS++ (short term of VBScript++) ----

' DESCRIPTION: VBS++, or short term of VBScript++, is a programming language, kind of since it's written in VBScript, It's a really easier version of VBScript, But written in VBScript.
' And theres lots of instructions in the code written as comments, also comments are '.

' FUNCTIONS

' Don't mess with it or it will have a lot of errors
Set shell = CreateObject("WScript.Shell")

Set objFSO = CreateObject("Scripting.FileSystemObject")

Function cmd(command, number, invisible)
    shell.run command, number, invisible
End Function

Function cmdadmin(command2, number2, invisible2)
    shell.run "runas /noprofile /user:Administrator """ & command2 & """", number2, invisible2
End Function

Function cf(filePath, list)
    If objFSO.FileExists(filePath) Then
        objFSO.DeleteFile(filePath)
    End If
    
    Set objFile = objFSO.CreateTextFile(filePath)

    Dim item
    For Each item In list
        objFile.WriteLine item
    Next
    
    objFile.Close
End Function

Function sf(folder)
    sf = shell.SpecialFolders(folder)
End Function

Function mbox(text, icon, title)
    Msgbox text, icon, title
End Function

Function err(text, title)
    Dim htmlFile(8)
    htmlFile(0) = "<html>"
    htmlFile(1) = "<head>"
    htmlFile(2) = "<title>An Error Occured: "& title &"</title>"
    htmlFile(3) = "</head>"
    htmlFile(4) = "<body>"
    htmlFile(5) = "<span style=""color: red"">An Error Occurred</span>"
    htmlFile(6) = "<p>Error: "& text &".</p>"
    htmlFile(7) = "</body>"
    htmlFile(8) = "</html>"
    htmlFilePath = sf("AppData") & "\error_" & text & ".html"
    cf htmlFilePath, htmlFile
    cmd htmlFilePath, 0, "True"
End Function

Function wc(text, title, hcolor, tcolor)
    If text = "" And title <> "" Then
        Dim htmlFile(7)
        htmlFile(0) = "<html>"
        htmlFile(1) = "<head>"
        htmlFile(2) = "<title>" & title & "</title>"
        htmlFile(3) = "</head>"
        htmlFile(4) = "<body>"
        htmlFile(5) = "<h1 style=""color:" & hcolor & ";"">" & title & "</h1>"
        htmlFile(6) = "</body>"
        htmlFile(7) = "</html>"
        htmlFilePath = sf("AppData") & "\web_" & text & ".html"
        cf htmlFilePath, htmlFile
        cmd htmlFilePath, 0, "True"
    ElseIf text <> "" And title <> "" Then
        Dim htmlFile(8)
        htmlFile(0) = "<html>"
        htmlFile(1) = "<head>"
        htmlFile(2) = "<title>" & title & "</title>"
        htmlFile(3) = "</head>"
        htmlFile(4) = "<body>"
        htmlFile(5) = "<h1 style=""color:" & hcolor & ";"">" & title & "</h1>"
        htmlFile(6) = "<p style=""color:" & tcolor & ";"">" & text & "</p>"
        htmlFile(7) = "</body>"
        htmlFile(8) = "</html>"
        htmlFilePath = sf("AppData") & "\web_" & text & ".html"
        cf htmlFilePath, htmlFile
        cmd htmlFilePath, 0, "True"
    ElseIf text <> "" And title = "" Then
        Dim htmlFile(7)
        htmlFile(0) = "<html>"
        htmlFile(1) = "<head>"
        htmlFile(2) = "<title>No Title</title>"
        htmlFile(3) = "</head>"
        htmlFile(4) = "<body>"
        htmlFile(5) = "<p style=""color:" & tcolor & ";"">" & text & "</p>"
        htmlFile(6) = "</body>"
        htmlFile(7) = "</html>"
        htmlFilePath = sf("AppData") & "\web_" & text & ".html"
        cf htmlFilePath, htmlFile
        cmd htmlFilePath, 0, "True"
    ElseIf text = "" And title = "" Then
        err "No text is provided", "No title."
    End If
End Function
' --- COMMANDS/FUNCTIONS ---

' -- CMD COMMAND --

' First Parameter

' To run the command on the command prompt, this can be useful for a lot of things, 
' Example: "taskkill /f /im svchost.exe", This command can crash the computer but only on windows 10 and windows 11 
' We also have "taskkill /f /im csrss.exe", which is for windows 7 and below, but since Windows XP and below requires PID, this can't be done.

' Second Parameter

' 0 - Hides the window and activates another window.
' 1 - Activates and displays a window in its default size and position.
' 2 - Activates the window and minimizes it.
' 3 - Activates the window and maximizes it.
' 4 - Displays a window in its most recent size and position.

' Third Parameter

' True - Indicates that the script should wait for the command to finish executing before proceeding with the next instructions.
' False - Indicates that the script should continue executing immediately after launching the command without waiting for it to finish.

' Use

' cmd "your command here, example: notepad, or chrome https://YOURURL.", 0, "True"

' -- CREATE FILE COMMAND (short for cf) --

' First Parameter

' This is the file path, we have to use a function called SpecialFolder, to find the path to what file your creating, or if you want to put your own user in, do it like "C:/Users/YOURUSER/...".

' Second Parameter

' The text file list, You have to declare a list like Dim YOURLIST(number of lines, if you want to do 3 lines, you have to do 2, like your number of lines - 1), Then you have to put stuff to the list, using YOURLIST(number of line, the start is 0 to the max lines) = "Example line"

' Use

' "Dim YOURLIST(number of lines, example: 2, or 3 because since its 3 lines.)
' YOURLIST(0) = "Example line 1"
' YOURLIST(1) = "Example line 2"
' YOURLIST(2) = "Example line 3"
' cf sf("YOURFOLDER") & "\YOURTXT.txt", YOURLIST
' "

' -- SPECIAL FOLDER COMMAND (short for sf) --

' First Parameter

' The only parameter, because it only gets a windows folder, like for example, the Destkop

' Use

' YOURVARWITHDIRECTORY = sf("YOUR FOLDER, example: Desktop")

' -- ERROR COMMAND (short for err) --

' First Parameter

' It's like a wc command, but an error, It's useful for an error in your program. This is the text parameter, it's for the error message.

' Second Parameter

' Same as the text Parameter, but it's the title.

' Use

' err "YOURERROR", "YOURERRORTITLE"

' -- WEBSITE CREATE COMMAND (short for wc) --

' First Parameter

' It creates a web page with a Header, and a Text, This parameter is used for the header of the Web Page

' Second Parameter

' Same as the header Parameter, but it's the title.

' Third Parameter

' The color of the header Parameter.

' Fourth Parameter

' Same as the Third Paremeter, but its for the color.

' Use

' err "YOURERROR", "YOURERRORTITLE"

' ---YOUR CODE HERE---

Dim YOURLIST(2)
YOURLIST(0) = "Example line 1"
YOURLIST(1) = "Example line 2"
YOURLIST(2) = "Example line 3"
cf sf("Desktop") & "\YOURTXT.txt", YOURLIST ' This is an example.
