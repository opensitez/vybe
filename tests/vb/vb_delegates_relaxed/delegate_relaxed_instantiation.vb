' vybe-test: vb/vb_delegates_relaxed/delegate_relaxed_instantiation
' origin: languages/vb/tests/vb/test_vb_delegates_relaxed.rs

Module M
    Sub PrintMessage(msg As String)
        Console.WriteLine(msg)
    End Sub

    ' Delegate requires an Object and EventArgs
    Delegate Sub EventHandler(sender As Object, e As EventArgs)

    Sub Main()
        ' Relaxed delegate instantiation allows dropping parameters if the target doesn't need them
        ' OR passing arguments that can be widened/narrowed automatically
        ' A common VB feature is assigning a Sub with no parameters to an EventHandler
        Dim handler As EventHandler = AddressOf LogEvent
        handler(Nothing, Nothing)
    End Sub

    Sub LogEvent()
        Console.WriteLine("Event Logged without parameters")
    End Sub
End Module
