' vybe-test: vb/vb_trycast_reference_types/test_vb_trycast_generic_method_type_parameter
' origin: languages/vb/tests/vb/test_vb_trycast_reference_types.rs

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

Module Program
    Private Function SafeCast(Of T As Class)(input As Object) As T
        Return TryCast(input, T)
    End Function

    Sub Main()
        Dim strObj As Object = "GenericCast"
        Dim intObj As Object = 100
        Dim resStr = SafeCast(Of String)(strObj)
        Dim resIntStr = SafeCast(Of String)(intObj)
        __Check(CStr(resStr & "|" & (resIntStr Is Nothing)), "GenericCast|True")
    End Sub
End Module
