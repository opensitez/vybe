' vybe-test: vb/vb_string_tokenizer_split_adv/test_vb_string_split_string_separator_array
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

Module Program
    Sub Main()
        Dim text As String = "Foo<BR>Bar<BR>Baz"
        Dim parts As String() = text.Split(New String() {"<BR>"}, StringSplitOptions.None)
        __Check(CStr(parts.Length), "3")
        __Check(CStr(parts(1)), "Bar")
    End Sub
End Module
