' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_using_statement_multiple_resources_same_type
' origin: languages/vb/tests/vb/test_vb_idisposable_double_dispose_safe.rs

Imports System

Class Tracker
    Implements IDisposable
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed " & Name)
    End Sub
End Class

Module Program
    Sub Main()
        Using r1 As New Tracker("R1"), r2 As New Tracker("R2")
            Console.WriteLine("Inside Multi Using")
        End Using
    End Sub
End Module
