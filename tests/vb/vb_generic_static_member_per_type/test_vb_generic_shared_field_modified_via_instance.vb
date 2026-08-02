' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_field_modified_via_instance
' origin: languages/vb/tests/vb/test_vb_generic_static_member_per_type.rs

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

Class SharedAccess(Of T)
    Public Shared Tag As String = "Default"
End Class

Module Program
    Sub Main()
        Dim o1 As New SharedAccess(Of Integer)()
        Dim o2 As New SharedAccess(Of Integer)()
        SharedAccess(Of Integer).Tag = "Modified"

        __Check(CStr(SharedAccess(Of Integer).Tag), "Modified")
    End Sub
End Module
