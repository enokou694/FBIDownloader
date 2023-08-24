Option Explicit

Const url = "https://github.com/Steveice10/FBI/releases/download/2.6.1/FBI.zip"

Dim http, stream, file, tmp
Set http = CreateObject("Msxml2.ServerXMLHTTP")
Set stream = CreateObject("ADODB.Stream")

tmp = Split(url, "/")
file = tmp(UBound(tmp))

On Error Resume Next
Call http.Open("GET", url, False )
http.Send
If(Err.Number <> 0)Then
	MsgBox("Error!")
	WScript.Quit
end if

stream.Open
stream.Type = 1
stream.Write http.responseBody
stream.SaveToFile file, 2
stream.Close
Set stream = Nothing
Set http = Nothing

MsgBox("Succsses!")