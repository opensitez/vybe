' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_reflection_get_set_access_rights
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

Class Sample
    Public Property Text As String { Get; Private Set; } = "Init"
End Class

Module Program
    Sub Main()
        Dim prop = GetType(Sample).GetProperty("Text")
        __Check(CStr((prop.GetMethod IsNot Nothing) & "|" & (prop.SetMethod IsNot Nothing)), "True|True")
    End Sub
End Module
