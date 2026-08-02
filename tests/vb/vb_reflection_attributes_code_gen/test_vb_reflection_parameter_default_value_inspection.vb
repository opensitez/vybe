' vybe-test: vb/vb_reflection_attributes_code_gen/test_vb_reflection_parameter_default_value_inspection
' origin: languages/vb/tests/vb/test_vb_reflection_attributes_code_gen.rs

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

Imports System

Class Options
    Public Sub Process(Optional timeout As Integer = 30)
    End Sub
End Class

Module Program
    Sub Main()
        Dim param = GetType(Options).GetMethod("Process").GetParameters()(0)
        __Check(CStr(param.HasDefaultValue & "|" & param.DefaultValue), "True|30")
    End Sub
End Module
