' vybe-test: vb/vb_using_statement/using_statement_basic
' origin: languages/vb/tests/vb/test_vb_using_statement.rs

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

Class FakeFile
    Implements IDisposable
    
    Public Sub Write(text As String)
        __Check(CStr("Writing: " & text), "Writing: Hello")
    End Sub
    
    Public Sub Dispose() Implements IDisposable.Dispose
        __Check(CStr("FakeFile Disposed"), "FakeFile Disposed")
    End Sub
End Class

Module M
    Sub Main()
        Using f As New FakeFile()
            f.Write("Hello")
        End Using
        __Check(CStr("Done"), "Done")
    End Sub
End Module
