' vybe-test: vb/vb_generic_delegate_type_args/test_vb_generic_delegate_tuple_arguments
' origin: languages/vb/tests/vb/test_vb_generic_delegate_type_args.rs

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

Delegate Function TupleProcessor(Of T1, T2)(pair As (Key As T1, Value As T2)) As String

Module Program
    Sub Main()
        Dim tp As TupleProcessor(Of String, Integer) = Function(pair) pair.Key & "=" & pair.Value
        __Check(CStr(tp(("Age", 30))), "Age=30")
    End Sub
End Module
