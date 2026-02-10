Imports System
Imports System.Net

Module Program
    Sub Main()
        Dim args() As String = Environment.GetCommandLineArgs()
        
        Dim url As String
        If args.Length > 1 Then
            url = args(1)
        Else
            url = "https://httpbin.org/get"
        End If
        
        Console.WriteLine("Fetching: " & url)
        Console.WriteLine()
        
        Dim client As New WebClient()
        Dim result As String = client.DownloadString(url)
        
        Console.WriteLine(result)
    End Sub
End Module
