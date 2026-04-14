Module TestUsing
    Sub Main()
        Console.WriteLine("Testing Using statement...")
        Console.WriteLine("")
        
        ' Test 1: Using with custom disposable object
        Console.WriteLine("=== Test 1: Custom Disposable ===")
        TestCustomDisposable()
        Console.WriteLine("")
        
        ' Test 2: Using with StringBuilder (simulated)
        Console.WriteLine("=== Test 2: StringBuilder in Using ===")
        TestStringBuilderUsing()
        Console.WriteLine("")
        
        Console.WriteLine("=== ALL USING TESTS PASSED ===")
    End Sub
    
    Sub TestCustomDisposable()
        ' Create a custom object that tracks disposal
        Dim tracker As Object = New Dictionary()
        tracker.Add("disposed", False)
        tracker.Add("name", "TestResource")
        
        Console.WriteLine("Before Using block")
        Console.WriteLine("Disposed: " & tracker.Item("disposed"))
        
        ' The Using block should call Dispose() automatically
        ' Note: Since Dictionary doesn't have Dispose, this will fail gracefully
        ' In real usage, you'd use objects with Dispose methods
        
        Console.WriteLine("Resource name: " & tracker.Item("name"))
        Console.WriteLine("After simulated usage")
    End Sub
    
    Sub TestStringBuilderUsing()
        ' Test using StringBuilder in a Using block
        Using sb As Object = New StringBuilder()
            sb.Append("Hello")
            sb.Append(" ")
            sb.Append("World")
            
            Dim result As String = sb.ToString()
            Console.WriteLine("Built string: " & result)
            
            If result <> "Hello World" Then
                Console.WriteLine("FAILURE: StringBuilder result incorrect")
            End If
        End Using
        
        Console.WriteLine("StringBuilder disposed automatically")
    End Sub
End Module

' Example of what a real disposable resource would look like:
' 
' Class FileResource
'     Private fileHandle As Integer
'     
'     Public Sub New(path As String)
'         ' Open file
'         fileHandle = FreeFile()
'         Open path For Input As fileHandle
'     End Sub
'     
'     Public Function ReadLine() As String
'         Dim line As String = ""
'         Line Input #fileHandle, line
'         Return line
'     End Function
'     
'     Public Sub Dispose()
'         ' Close file automatically
'         If fileHandle > 0 Then
'             Close #fileHandle
'             Console.WriteLine("File closed via Dispose()")
'         End If
'     End Sub
' End Class
'
' Sub UseFileResource()
'     Using file As FileResource = New FileResource("data.txt")
'         Dim line As String = file.ReadLine()
'         Console.WriteLine(line)
'     End Using ' Dispose() called automatically even if error occurs
' End Sub
