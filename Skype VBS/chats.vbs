Set oSkype = Wscript.CreateObject("Skype4COM.Skype", "Skype_")
oSkype.Attach 5
WScript.Echo "All chats:"
For Each oChat In oSkype.Chats
WScript.Echo oChat.Timestamp & " " & oChat.Name & " " & oChat.FriendlyName
Next 
'WScript.Echo "Caca"