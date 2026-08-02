' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_method_info_anonymous_type_return
' origin: languages/vb/tests/vb/test_vb_reflection_method_info_generic_invoke.rs

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

Class AnonService
    Public Function GetAnon() As Object
        Return New With {.Status = "AnonSuccess"}
    End Function
End Class

Module Program
    Sub Main()
        Dim svc As New AnonService()
        Dim m = GetType(AnonService).GetMethod("GetAnon")
        Dim res As Dynamic = m.Invoke(svc, Nothing)
        __Check(CStr(res.Status), "AnonSuccess")
    End Sub
End Module
