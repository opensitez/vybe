' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_get_properties_binding_flags
' origin: languages/vb/tests/vb/test_vb_reflection_property_info_indexers.rs

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

Class FilterTest
    Public Property P1 As Integer
    Private Property P2 As String
    Public Shared Property P3 As Double
End Class

Module Program
    Sub Main()
        Dim props = GetType(FilterTest).GetProperties(BindingFlags.Instance Or BindingFlags.Public)
        __Check(CStr(props.Length & ":" & props(0).Name), "1:P1")
    End Sub
End Module
