' vybe-test: vb/vb_system_queue_stack_matrix/stack_to_array_preserves_order_reversed
' origin: languages/vb/tests/vb/test_vb_system_queue_stack_matrix.rs

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

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim stack As New Stack(Of Integer)()
        stack.Push(1)
        stack.Push(2)
        stack.Push(3)
        Dim items() As Integer = stack.ToArray()
        __Check(CStr(items.Length), "3")
        __Check(CStr(items(0)), "3")
        __Check(CStr(items(2)), "1")
    End Sub
End Module
