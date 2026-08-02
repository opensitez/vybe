' vybe-test: vb/vb_empty_catch_block/empty_catch_block
' origin: languages/vb/tests/vb/test_vb_empty_catch_block.rs

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
        Try
            Dim i = 1 \ 0
        Catch
            ' Empty catch block to swallow exceptions
        End Try
        __Check(CStr("Survived"), "Survived")
    End Sub
End Module
