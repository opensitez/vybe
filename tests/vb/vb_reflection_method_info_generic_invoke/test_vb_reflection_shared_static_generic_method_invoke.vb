' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_shared_static_generic_method_invoke
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

Class Helper
    Public Shared Function Wrap(Of T)(item As T) As String
        Return "[" & item.ToString() & "]"
    End Function
End Class

Module Program
    Sub Main()
        Dim m = GetType(Helper).GetMethod("Wrap").MakeGenericMethod(GetType(Double))
        Dim res = m.Invoke(Nothing, {3.14})
        __Check(CStr(res), "[3.14]")
    End Sub
End Module
