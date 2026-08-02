' vybe-test: vb/vb_stack_generic_operations/test_vb_stack_peek_non_destructive
' origin: languages/vb/tests/vb/test_vb_stack_generic_operations.rs

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

Module Program
    Sub Main()
        Dim st As New Stack(Of String)()
        st.Push("Top")
        __Check(CStr(st.Peek()), "Top")
        __Check(CStr(st.Count), "1")
    End Sub
End Module
