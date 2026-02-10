Module TestAsync
    Async Function Calculate() As Integer
        Return 42
    End Function

    Async Sub Printer(msg As String)
        Console.WriteLine("Msg: " & msg)
    End Sub

    Sub Main()
        Console.WriteLine("Start")
        ' In simulation, Await simply evaluates. 
        ' Since Calculate returns 42 (not a Task object in this phase), Await 42 -> 42.
        Dim res = Await Calculate()
        Console.WriteLine("Result: " & res)
        
        ' Printer is async but void (Sub). Await Printer("Hello") -> Nothing?
        ' In VB, Await requires Awaitable (Task). Sub returns nothing (void).
        ' Await Printer(...) is not valid in VB if Printer returns Void.
        ' Expected: Parser allows it, Runtime might evaluate to Nothing.
        ' But for test simplicity let's stick to Function.
        
        Printer("Async Sub Call")
        Console.WriteLine("End")
    End Sub
End Module
