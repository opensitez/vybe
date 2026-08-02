' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_method_info_invoke_throws_target_invocation_exception
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

Imports System
Imports System.Reflection

Class FaultyMethod
    Public Sub Fail()
        Throw New ArgumentException("Invalid Argument")
    End Sub
End Class

Module Program
    Sub Main()
        Dim fm As New FaultyMethod()
        Dim m = GetType(FaultyMethod).GetMethod("Fail")
        Try
            m.Invoke(fm, Nothing)
        Catch ex As TargetInvocationException
            __Check(CStr(ex.InnerException.GetType().Name & ": " & ex.InnerException.Message), "ArgumentException: Invalid Argument")
        End Try
    End Sub
End Module
