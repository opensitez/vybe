' vybe-test: vb/vb_generic_delegate_type_args/test_vb_generic_delegate_two_type_args
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

Delegate Function Evaluator(Of TIn, TOut)(input As TIn) As TOut

Module Program
    Sub Main()
        Dim ev As Evaluator(Of String, Integer) = Function(s) s.Length
        __Check(CStr(ev("VisualBasic")), "11")
    End Sub
End Module
