' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_private_method_invocation_binding_flags
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

Imports System.Reflection

Class InternalProcessor
    Private Function SecretFormula(x As Integer) As Integer
        Return x * 7
    End Function
End Class

Module Program
    Sub Main()
        Dim proc As New InternalProcessor()
        Dim m = GetType(InternalProcessor).GetMethod("SecretFormula", BindingFlags.Instance Or BindingFlags.NonPublic)
        Dim res = m.Invoke(proc, {5})
        __Check(CStr(res), "35")
    End Sub
End Module
