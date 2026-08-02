' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_overloaded_methods_resolution
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

Class OverloadSample
    Public Function Compute(x As Integer) As String : Return "Int_" & x : End Function
    Public Function Compute(x As String) As String : Return "Str_" & x : End Function
End Class

Module Program
    Sub Main()
        Dim s As New OverloadSample()
        Dim mInt = GetType(OverloadSample).GetMethod("Compute", {GetType(Integer)})
        Dim mStr = GetType(OverloadSample).GetMethod("Compute", {GetType(String)})

        __Check(CStr(mInt.Invoke(s, {10}) & "|" & mStr.Invoke(s, {"abc"})), "Int_10|Str_abc")
    End Sub
End Module
