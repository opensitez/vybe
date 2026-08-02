' vybe-test: vb/vb_optional_arguments_objects/optional_arguments_objects
' origin: languages/vb/tests/vb/test_vb_optional_arguments_objects.rs

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

Class Configuration
End Class

Module M
    ' Optional parameters of object type must be Nothing
    Sub Initialize(Optional config As Configuration = Nothing)
        If config Is Nothing Then
            __Check(CStr("Default Config"), "Default Config")
        Else
            __Check(CStr("Custom Config"), "Custom Config")
        End If
    End Sub

    Sub Main()
        Initialize()
        Initialize(New Configuration())
    End Sub
End Module
