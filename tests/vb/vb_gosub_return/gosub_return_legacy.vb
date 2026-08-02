' vybe-test: vb/vb_gosub_return/gosub_return_legacy
' origin: languages/vb/tests/vb/test_vb_gosub_return.rs

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
        Dim x As Integer = 1
        GoSub DoubleIt
        GoSub DoubleIt
        __Check(CStr(x), "4")
        Exit Sub
        
DoubleIt:
        x *= 2
        Return ' In a Sub with GoSub, Return jumps back to the line after GoSub
    End Sub
End Module
