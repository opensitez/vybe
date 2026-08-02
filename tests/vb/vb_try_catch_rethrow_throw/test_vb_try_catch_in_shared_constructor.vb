' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_try_catch_in_shared_constructor
' origin: languages/vb/tests/vb/test_vb_try_catch_rethrow_throw.rs

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

Class SafeSharedInit
    Public Shared Loaded As Boolean = False
    Shared Sub New()
        Try
            Throw New Exception("Init Fail")
        Catch ex As Exception
            __Check(CStr("Caught in Shared Sub New"), "Caught in Shared Sub New")
            Loaded = True
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        __Check(CStr(SafeSharedInit.Loaded), "True")
    End Sub
End Module
