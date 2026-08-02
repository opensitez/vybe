' vybe-test: vb/vb_interop/h65_string_concatenation_mixed_types
' origin: languages/vb/tests/vb/vb_interop_test.rs

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

__Check(CStr("Count: " & 42), "Count: 42")
__Check(CStr("Pi: " & 3.14), "Pi: 3.14")
__Check(CStr("Active: " & True), "Active: true")
__Check(CStr("Hello" & " " & "World"), "Hello World")
