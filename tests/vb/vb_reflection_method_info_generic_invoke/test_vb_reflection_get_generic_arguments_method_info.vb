' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_get_generic_arguments_method_info
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

Class Service
    Public Sub Execute(Of T1, T2)() : End Sub
End Class

Module Program
    Sub Main()
        Dim m = GetType(Service).GetMethod("Execute")
        Dim typeArgs = m.GetGenericArguments()
        __Check(CStr(typeArgs.Length & ":" & typeArgs(0).Name & "," & typeArgs(1).Name), "2:T1,T2")
    End Sub
End Module
