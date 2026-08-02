' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_field_nested_class_access
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

Class Outer(Of T)
    Public Shared OuterData As String = "Outer"

    Public Class Inner
        Public Function GetOuterData() As String
            Return OuterData
        End Function
    End Class
End Class

Module Program
    Sub Main()
        Outer(Of Integer).OuterData = "IntOuter"
        Outer(Of String).OuterData = "StringOuter"

        Dim inInt As New Outer(Of Integer).Inner()
        Dim inStr As New Outer(Of String).Inner()

        __Check(CStr(inInt.GetOuterData() & "|" & inStr.GetOuterData()), "IntOuter|StringOuter")
    End Sub
End Module
