' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_method_type_argument_inference
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

Module GenericHelper
    Public Function Identity(Of T)(item As T) As T
        Return item
    End Function
End Module

Module Program
    Sub Main()
        __Check(CStr(GenericHelper.Identity(10) & "|" & GenericHelper.Identity("ABC")), "10|ABC")
    End Sub
End Module
