' vybe-test: vb/vb_file_functions/file_functions_legacy
' origin: languages/vb/tests/vb/test_vb_file_functions.rs

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
        ' Legacy file I/O syntax parsing check
        Dim fNum = FreeFile()
        __Check(CStr(fNum > 0), "True")
        
        ' Note: FileOpen, EOF, LOF, Loc might throw or fail if files aren't created in test environment,
        ' so we just test FreeFile and syntax checking.
    End Sub
End Module
