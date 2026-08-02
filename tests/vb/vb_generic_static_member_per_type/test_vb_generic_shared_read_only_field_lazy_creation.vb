' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_read_only_field_lazy_creation
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

Imports System

Class LazyContainer(Of T As New)
    Public Shared ReadOnly Instance As New T()
End Class

Class User
    Public Name As String = "DefaultUser"
End Class

Module Program
    Sub Main()
        __Check(CStr(LazyContainer(Of User).Instance.Name), "DefaultUser")
    End Sub
End Module
