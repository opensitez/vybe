' vybe-test: vb/vb_redim_preserve_multidim/redim_preserve_multidim
' origin: languages/vb/tests/vb/test_vb_redim_preserve_multidim.rs

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
        ' Only the last dimension can be resized when using Preserve
        Dim arr(1, 1) As Integer
        arr(0, 0) = 1
        arr(0, 1) = 2
        arr(1, 0) = 3
        arr(1, 1) = 4
        
        ReDim Preserve arr(1, 2)
        arr(0, 2) = 5
        arr(1, 2) = 6
        
        __Check(CStr(arr(0, 0)), "1")
        __Check(CStr(arr(1, 1)), "4")
        __Check(CStr(arr(1, 2)), "6")
    End Sub
End Module
