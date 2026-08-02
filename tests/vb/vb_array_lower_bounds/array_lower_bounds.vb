' vybe-test: vb/vb_array_lower_bounds/array_lower_bounds
' origin: languages/vb/tests/vb/test_vb_array_lower_bounds.rs

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
        ' VB supports specifying the lower bound, though it must be 0 in .NET
        Dim arr(0 To 2) As Integer
        arr(0) = 1
        arr(1) = 2
        arr(2) = 3
        
        __Check(CStr(arr.Length), "3")
        __Check(CStr(arr(1)), "2")
    End Sub
End Module
