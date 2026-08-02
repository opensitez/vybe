' vybe-test: vb/vb_file_io_legacy/fileio_filelen
' origin: languages/vb/tests/vb/test_vb_file_io_legacy.rs

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

Module M: Sub Main(): Dim f = FreeFile(): FileOpen(f, "test_filelen.txt", OpenMode.Output): PrintLine(f, "Hello"): FileClose(f): __Check(CStr(FileLen("test_filelen.txt") > 0), "True"): End Sub: End Module
