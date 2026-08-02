' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_idisposable_nested_using_order
' origin: languages/vb/tests/vb/test_vb_idisposable_double_dispose_safe.rs

Imports System

Class LevelTracker
    Implements IDisposable
    Private level As String
    Public Sub New(l As String)
        level = l
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Exit Level " & level)
    End Sub
End Class

Module Program
    Sub Main()
        Using Outer As New LevelTracker("Outer")
            Using Inner As New LevelTracker("Inner")
                Console.WriteLine("Innermost Action")
            End Using
        End Using
    End Sub
End Module
