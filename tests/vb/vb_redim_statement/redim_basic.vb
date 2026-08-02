' vybe-test: vb/vb_redim_statement/redim_basic
' origin: languages/vb/tests/vb/test_vb_redim_statement.rs

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
        Dim arr() As Integer
        ReDim arr(2)
        arr(0) = 10
        arr(1) = 20
        arr(2) = 30
        __Check(CStr(arr(1)), "20")
        
        ReDim arr(4)
        __Check(CStr(arr(1)), "0") ' Should be 0 because elements are reset
    End Sub
End Module
