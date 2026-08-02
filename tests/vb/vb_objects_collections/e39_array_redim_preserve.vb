' vybe-test: vb/vb_objects_collections/e39_array_redim_preserve
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

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

Dim arr(3) As Integer
arr(0) = 10
arr(1) = 20
arr(2) = 30
ReDim Preserve arr(5)
__Check(CStr(arr(0)), "10")
__Check(CStr(arr(1)), "20")
__Check(CStr(arr(2)), "30")
__Check(CStr(UBound(arr)), "5")
