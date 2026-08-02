' vybe-test: vb/vb_array_copy_clear_clone/test_vb_array_clone_shallow_copy
' origin: languages/vb/tests/vb/test_vb_array_copy_clear_clone.rs

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
        Dim orig As String() = {"X", "Y", "Z"}
        Dim cloneArr As String() = CType(orig.Clone(), String())
        cloneArr(0) = "W"
        __Check(CStr(orig(0)), "X")
        __Check(CStr(cloneArr(0)), "W")
    End Sub
End Module
