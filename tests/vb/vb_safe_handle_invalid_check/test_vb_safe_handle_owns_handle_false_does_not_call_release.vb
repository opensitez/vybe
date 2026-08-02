' vybe-test: vb/vb_safe_handle_invalid_check/test_vb_safe_handle_owns_handle_false_does_not_call_release
' origin: languages/vb/tests/vb/test_vb_safe_handle_invalid_check.rs

Imports System
Imports System.Runtime.InteropServices

Class NonOwnerSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=False) ' Does not own handle!
        SetHandle(New IntPtr(444))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Console.WriteLine("Should Not Be Called")
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New NonOwnerSafeHandle()
        h.Dispose()
        Console.WriteLine("NonOwner Disposed Safely")
    End Sub
End Module
