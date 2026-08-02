' vybe-test: vb/vb_anonymous_type_generic_inference/test_vb_anonymous_type_generic_extension_method_invocation
' origin: languages/vb/tests/vb/test_vb_anonymous_type_generic_inference.rs

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

Imports System.Runtime.CompilerServices

Module GenericExtensions
    <Extension()>
    Public Function ToJsonLikeString(Of T)(obj As T) As String
        Return obj.ToString()
    End Function
End Module

Module Program
    Sub Main()
        Dim item = New With {.ID = 10, .Name = "Item10"}
        __Check(CStr(item.ToJsonLikeString().Contains("ID = 10")), "True")
    End Sub
End Module
