' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_private_setter
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

Class Config
    Public Property Mode As String
        Get
            Return "Production"
        End Get
        Private Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim prop = GetType(Config).GetProperty("Mode")
        Dim pubSet = prop.GetSetMethod(False)
        Dim nonPubSet = prop.GetSetMethod(True)
        __Check(CStr((pubSet Is Nothing) & "|" & (nonPubSet IsNot Nothing)), "True|True")
    End Sub
End Module
