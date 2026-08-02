' vybe-test: vb/vb_anonymous_type_generic_inference/test_vb_anonymous_type_multiple_generic_type_parameters_inferred
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

Module Program
    Private Function PairUp(Of T1, T2)(a As T1, b As T2) As Object
        Return New With {.First = a, .Second = b}
    End Function

    Sub Main()
        Dim pair As Object = PairUp(100, "Hundred")
        Dim fProp = pair.GetType().GetProperty("First")
        Dim sProp = pair.GetType().GetProperty("Second")
        __Check(CStr(fProp.GetValue(pair) & "=" & sProp.GetValue(pair)), "100=Hundred")
    End Sub
End Module
