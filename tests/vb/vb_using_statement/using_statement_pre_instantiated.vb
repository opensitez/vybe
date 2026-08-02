' vybe-test: vb/vb_using_statement/using_statement_pre_instantiated
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

Class Resource
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        __Check(CStr("Disposed"), "Inside")
    End Sub
End Class

Module M
    Sub Main()
        Dim r As New Resource()
        Using r
            __Check(CStr("Inside"), "Disposed")
        End Using
    End Sub
End Module
