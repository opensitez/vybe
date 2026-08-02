' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_auto_property_initializer_with_private_set
' origin: languages/vb/tests/vb/test_vb_property_protected_internal_mix.rs

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

Class SystemDefaults
    Public Property MaxRetries As Integer { Get; Private Set; } = 3
End Class

Module Program
    Sub Main()
        Dim sd As New SystemDefaults()
        __Check(CStr(sd.MaxRetries), "3")
    End Sub
End Module
