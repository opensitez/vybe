' vybe-test: vb/vb_array_bounds/array_bounds_upper
' origin: languages/vb/tests/vb/test_vb_array_bounds.rs

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
        ' In VB.NET, you specify the upper bound, not the length
        Dim arr(2) As Integer
        
        arr(0) = 10
        arr(1) = 20
        arr(2) = 30
        
        __Check(CStr(arr.Length), "3") ' Length is 3
        __Check(CStr(arr.GetUpperBound(0)), "2") ' Upper bound is 2
    End Sub
End Module
