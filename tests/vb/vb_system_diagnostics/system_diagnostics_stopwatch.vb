' vybe-test: vb/vb_system_diagnostics/system_diagnostics_stopwatch
' origin: languages/vb/tests/vb/test_vb_system_diagnostics.rs

Imports System.Diagnostics

Module M
    Sub Main()
        Dim sw As New Stopwatch()
        sw.Start()
        ' Simulate some work
        Dim sum = 0
        For i = 1 To 1000
            sum += i
        Next
        sw.Stop()
        
        Console.WriteLine(sw.IsRunning)
        Console.WriteLine(sw.ElapsedMilliseconds >= 0)
    End Sub
End Module
