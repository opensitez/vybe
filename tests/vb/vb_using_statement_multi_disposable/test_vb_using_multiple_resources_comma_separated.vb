' vybe-test: vb/vb_using_statement_multi_disposable/test_vb_using_multiple_resources_comma_separated
' origin: languages/vb/tests/vb/test_vb_using_statement_multi_disposable.rs

Imports System

Class Tracker
    Implements IDisposable
    Public Name As String
    Public Sub New(n As String)
        Me.Name = n
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed:" & Name)
    End Sub
End Class

Module Program
    Sub Main()
        Using r1 As New Tracker("R1"), r2 As New Tracker("R2")
            Console.WriteLine("Inside block")
        End Using
    End Sub
End Module
