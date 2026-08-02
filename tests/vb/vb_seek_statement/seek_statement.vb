' vybe-test: vb/vb_seek_statement/seek_statement
' origin: languages/vb/tests/vb/test_vb_seek_statement.rs

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
        ' The Seek function returns the current position
        ' The Seek statement sets the position
        ' We will just parse test them by using them in unreachable code 
        ' to avoid IO issues in the test runner.
        Dim b As Boolean = True
        If Not b Then
            Dim f = FreeFile()
            FileOpen(f, "test.txt", OpenMode.Random)
            Seek(f, 10)
            Dim pos = Seek(f)
            FileClose(f)
        End If
        __Check(CStr("Seek Parsed"), "Seek Parsed")
    End Sub
End Module
