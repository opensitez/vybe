' vybe-test: vb/vb_directcast_value_types/test_vb_directcast_interface_cast_succeeds
' origin: languages/vb/tests/vb/test_vb_directcast_value_types.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System

Interface IRunner
    Sub Run()
End Interface

Class Worker
    Implements IRunner
    Public Sub Run() Implements IRunner.Run
        __Check(CStr("Worker Running"), "Worker Running")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As Object = New Worker()
        Dim runner As IRunner = DirectCast(obj, IRunner)
        runner.Run()
    End Sub
End Module
