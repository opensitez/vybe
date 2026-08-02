' vybe-test: vb/vb_string_tokenizer_split_adv/test_vb_string_concat_enumerable
' origin: languages/vb/tests/vb/test_vb_string_tokenizer_split_adv.rs

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
        Dim nums As New List(Of Integer) From {1, 2, 3, 4}
        Dim res As String = String.Concat(nums)
        __Check(CStr(res), "1234")
    End Sub
End Module
