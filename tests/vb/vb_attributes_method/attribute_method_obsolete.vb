' vybe-test: vb/vb_attributes_method/attribute_method_obsolete
' origin: languages/vb/tests/vb/test_vb_attributes_method.rs

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

Class LegacyCode
    <Obsolete("Use NewMethod instead")>
    Public Sub OldMethod()
        __Check(CStr("Old"), "Old")
    End Sub
End Class

Module M
    Sub Main()
        Dim l As New LegacyCode()
        ' Should run even if obsolete
        l.OldMethod()
        __Check(CStr("Done"), "Done")
    End Sub
End Module
