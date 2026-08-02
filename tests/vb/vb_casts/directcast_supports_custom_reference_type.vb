' vybe-test: vb/vb_casts/directcast_supports_custom_reference_type
' origin: languages/vb/tests/vb/test_vb_casts.rs

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
    Class Greeter
        Public Message As String = "hi"
    End Class

    Sub Main()
        Dim boxed As Object = New Greeter()
        Dim value As Greeter = DirectCast(boxed, Greeter)
        __Check(CStr(value.Message), "hi")
    End Sub
End Module
