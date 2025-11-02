' Create an ASP file with a redirect
Dim fso, aspFile
Set fso = CreateObject("Scripting.FileSystemObject")
url = InputBox("Enter URL")
' Define the path for the new ASP file
aspFile = InputBox("Enter ShortURL ID")&".asp"

' Create the ASP file and write the redirect code to it
Dim file
Set file = fso.CreateTextFile(aspFile, True)

file.WriteLine("<%")
file.WriteLine("' Redirect to another page")
file.WriteLine("Response.Redirect("&chr(34)&url&chr(34)&")")
file.WriteLine("%>")
file.Close


WScript.Echo "ASP file created successfully at: " & aspFile
