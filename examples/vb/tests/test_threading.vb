Imports System.Threading

Module ThreadingTest
    Sub Main()
        Console.WriteLine("--- Threading Verification ---")
        
        Dim result = 0
        Dim t = New Thread(Sub()
            Console.WriteLine("Background thread: Doing work...")
            Thread.Sleep(200)
            result = 100
            Console.WriteLine("Background thread: Work complete.")
        End Sub)
        
        t.Start()
        Console.WriteLine("Main thread: Waiting for background thread...")
        
        ' Busy wait or sleep
        Thread.Sleep(500)
        
        Console.WriteLine("Final Result: " & result)
        
        If result = 100 Then
            Console.WriteLine("Success: Background thread updated shared state (snapshot).")
        Else
            Console.WriteLine("Failure: Background thread did not update shared state.")
        End If
    End Sub
End Module
