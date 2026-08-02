' vybe-test: vb/vb_oop_edges/generic_method_type_inference_arrays
' origin: languages/vb/tests/vb/test_vb_oop_edges.rs

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
    Sub PrintFirst(Of T)(arr() As T)
        __Check(CStr(arr(0)), "10")
    End Sub

    Sub Main()
        Dim nums = {10, 20}
        PrintFirst(nums) ' Type inference T=Integer
    End Sub
End Module
