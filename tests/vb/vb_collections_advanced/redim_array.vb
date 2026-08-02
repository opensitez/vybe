' vybe-test: vb/vb_collections_advanced/redim_array
' origin: languages/vb/tests/vb/test_vb_collections_advanced.rs

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
        Dim arr(2) As Integer
        arr(0) = 1
        arr(1) = 2
        arr(2) = 3
        ReDim Preserve arr(4)
        arr(3) = 4
        arr(4) = 5
        __Check(CStr(UBound(arr)), "4")
        __Check(CStr(arr(0)), "1")
        __Check(CStr(arr(4)), "5")
    End Sub
End Module
