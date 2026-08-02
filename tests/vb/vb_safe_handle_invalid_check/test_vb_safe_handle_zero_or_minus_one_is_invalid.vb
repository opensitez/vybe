' vybe-test: vb/vb_safe_handle_invalid_check/test_vb_safe_handle_zero_or_minus_one_is_invalid
' origin: languages/vb/tests/vb/test_vb_safe_handle_invalid_check.rs

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
Imports System.Runtime.InteropServices

Class StandardSafeHandle
    Inherits SafeHandleZeroOrMinusOneIsInvalid

    Public Sub New(ownsHandle As Boolean)
        MyBase.New(ownsHandle)
        SetHandle(New IntPtr(-1))
    End Sub

    Protected Overrides Function ReleaseHandle() As Boolean
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New StandardSafeHandle(True)
        ' SafeHandleZeroOrMinusOneIsInvalid considers -1 as invalid!
        __Check(CStr(h.IsInvalid), "True")
    End Sub
End Module
