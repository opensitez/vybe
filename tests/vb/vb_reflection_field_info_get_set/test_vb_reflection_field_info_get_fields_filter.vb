' vybe-test: vb/vb_reflection_field_info_get_set/test_vb_reflection_field_info_get_fields_filter
' origin: languages/vb/tests/vb/test_vb_reflection_field_info_get_set.rs

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

Imports System.Reflection

Class MultiField
    Public F1 As Integer
    Public F2 As String
    Private F3 As Double
End Class

Module Program
    Sub Main()
        Dim pubFields = GetType(MultiField).GetFields(BindingFlags.Instance Or BindingFlags.Public)
        Dim allFields = GetType(MultiField).GetFields(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        __Check(CStr(pubFields.Length & "|" & allFields.Length), "2|3")
    End Sub
End Module
