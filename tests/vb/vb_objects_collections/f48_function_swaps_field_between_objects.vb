' vybe-test: vb/vb_objects_collections/f48_function_swaps_field_between_objects
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

Class Pair
    Public Val As String
End Class
Sub SwapVals(a As Pair, b As Pair)
    Dim tmp As String = a.Val
    a.Val = b.Val
    b.Val = tmp
End Sub
Dim p1 As New Pair()
p1.Val = "hello"
Dim p2 As New Pair()
p2.Val = "world"
SwapVals(p1, p2)
__Check(CStr(p1.Val), "world")
__Check(CStr(p2.Val), "hello")
