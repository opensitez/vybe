' vybe-test: vb/vb_objects_collections/e44_array_of_arrays
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

Dim outer(2) As Object
Dim inner1(2) As Integer
inner1(0) = 10
inner1(1) = 20
Dim inner2(2) As Integer
inner2(0) = 30
inner2(1) = 40
outer(0) = inner1
outer(1) = inner2
__Check(CStr(outer(0)(0)), "10")
__Check(CStr(outer(1)(1)), "40")
