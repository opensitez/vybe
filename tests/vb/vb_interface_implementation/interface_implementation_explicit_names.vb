' vybe-test: vb/vb_interface_implementation/interface_implementation_explicit_names
' origin: languages/vb/tests/vb/test_vb_interface_implementation.rs

Interface IPrinter
    Sub Print()
End Interface

Class ConsolePrinter
    Implements IPrinter
    
    ' In VB.NET, the implementing method name doesn't have to match the interface method name
    ' The Implements clause defines what it implements
    Public Sub Output() Implements IPrinter.Print
        Console.WriteLine("Printed explicitly")
    End Sub
End Class

Module M
    Sub Main()
        Dim p As IPrinter = New ConsolePrinter()
        p.Print()
        
        Dim cp As New ConsolePrinter()
        cp.Output()
    End Sub
End Module
