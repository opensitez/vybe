' vybe-test: vb/vb_comments/comments_rem_and_quote
' origin: languages/vb/tests/vb/test_vb_comments.rs

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
        REM This is a comment using REM keyword
        __Check(CStr("Start"), "Start")
        ' This is a standard single quote comment
        Dim x As Integer = 5 REM Inline REM comment
        Dim y As Integer = 10 ' Inline quote comment
        __Check(CStr(x + y), "15")
    End Sub
End Module
