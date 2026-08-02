' vybe-test: vb/vb_late_binding_arrays/late_binding_arrays
' origin: languages/vb/tests/vb/test_vb_late_binding_arrays.rs

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

Option Strict Off

Module M
    Sub Main()
        Dim obj As Object = New Integer() {1, 2, 3}
        
        ' Late bound array indexing
        __Check(CStr(obj(1)), "2")
        
        obj(2) = 10
        __Check(CStr(obj(2)), "10")
    End Sub
End Module
