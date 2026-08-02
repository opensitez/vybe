' vybe-test: vb/vb_callbyname_function_invocation/test_vb_callbyname_null_target_throws_exception
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

Module Program
    Sub Main()
        Try
            CallByName(Nothing, "AnyMethod", CallType.Method)
        Catch ex As ArgumentNullException
            __Check(CStr("ArgumentNullException Caught on Null CallByName Instance"), "ArgumentNullException Caught on Null CallByName Instance")
        End Try
    End Sub
End Module
