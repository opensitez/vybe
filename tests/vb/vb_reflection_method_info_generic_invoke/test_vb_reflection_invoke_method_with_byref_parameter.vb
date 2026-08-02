' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_invoke_method_with_byref_parameter
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

Class Processor
    Public Sub DoubleValue(ByRef num As Integer)
        num *= 2
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Processor()
        Dim m = GetType(Processor).GetMethod("DoubleValue")
        Dim args As Object() = {25}
        m.Invoke(p, args)
        __Check(CStr(args(0)), "50")
    End Sub
End Module
