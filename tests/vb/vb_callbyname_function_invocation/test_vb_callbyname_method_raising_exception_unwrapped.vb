' vybe-test: vb/vb_callbyname_function_invocation/test_vb_callbyname_method_raising_exception_unwrapped
' origin: languages/vb/tests/vb/test_vb_callbyname_function_invocation.rs

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
Imports Microsoft.VisualBasic

Class FaultyComponent
    Public Sub Crash()
        Throw New InvalidOperationException("Component Crashed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim comp As New FaultyComponent()
        Try
            CallByName(comp, "Crash", CallType.Method)
        Catch ex As InvalidOperationException
            __Check(CStr(ex.Message), "Component Crashed")
        End Try
    End Sub
End Module
