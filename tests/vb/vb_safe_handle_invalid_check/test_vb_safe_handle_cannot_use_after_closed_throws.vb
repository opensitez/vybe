' vybe-test: vb/vb_safe_handle_invalid_check/test_vb_safe_handle_cannot_use_after_closed_throws
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

Class StrictSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(1))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New StrictSafeHandle()
        h.Dispose()

        Try
            Dim success = False
            h.DangerousAddRef(success)
        Catch ex As ObjectDisposedException
            __Check(CStr("ObjectDisposedException Caught on Closed SafeHandle"), "ObjectDisposedException Caught on Closed SafeHandle")
        End Try
    End Sub
End Module
