' vybe-test: vb/vb_generics_methods/generic_method_type_inference
' origin: languages/vb/tests/vb/test_vb_generics_methods.rs

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

Module M
    Sub Swap(Of T)(ByRef a As T, ByRef b As T)
        Dim temp As T = a
        a = b
        b = temp
    End Sub

    Sub Main()
        Dim x As Integer = 1
        Dim y As Integer = 2
        ' Type parameter omitted, compiler infers (Of Integer)
        Swap(x, y)
        __Check(CStr(x), "2")
        __Check(CStr(y), "1")
    End Sub
End Module
