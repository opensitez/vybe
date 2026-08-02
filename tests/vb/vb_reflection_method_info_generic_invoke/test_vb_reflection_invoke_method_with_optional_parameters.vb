' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_invoke_method_with_optional_parameters
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

Class Printer
    Public Function PrintMsg(msg As String, Optional prefix As String = "LOG:") As String
        Return prefix & " " & msg
    End Function
End Class

Module Program
    Sub Main()
        Dim p As New Printer()
        Dim m = GetType(Printer).GetMethod("PrintMsg")
        Dim res = m.Invoke(p, BindingFlags.OptionalParamBinding, Nothing, {"Hello", Missing.Value}, Nothing)
        __Check(CStr(res), "LOG: Hello")
    End Sub
End Module
