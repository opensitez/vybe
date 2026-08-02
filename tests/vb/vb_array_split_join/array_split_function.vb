' vybe-test: vb/vb_array_split_join/array_split_function
' origin: languages/vb/tests/vb/test_vb_array_split_join.rs

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

Module M
    Sub Main()
        Dim text As String = "apple,banana,cherry"
        Dim parts As String() = Split(text, ",")
        
        __Check(CStr(parts(0)), "apple")
        __Check(CStr(parts(2)), "cherry")
    End Sub
End Module
