' vybe-test: vb/vb_legacy_error_handling/err_object_properties
' origin: languages/vb/tests/vb/test_vb_legacy_error_handling.rs

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
        On Error GoTo Handler
        Err.Raise(1234, "MySource", "MyDescription")
        Exit Sub
        
Handler:
        __Check(CStr(Err.Number), "1234")
        __Check(CStr(Err.Source), "MySource")
        __Check(CStr(Err.Description), "MyDescription")
        Err.Clear()
        __Check(CStr(Err.Number), "0")
    End Sub
End Module
