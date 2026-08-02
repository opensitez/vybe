' vybe-test: vb/vb_safe_handle_invalid_check/test_vb_custom_safe_handle_is_invalid_check
' origin: languages/vb/tests/vb/test_vb_safe_handle_invalid_check.rs

Imports System
Imports System.Runtime.InteropServices

Class DummySafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Console.WriteLine("Handle Released")
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New DummySafeHandle()
        Console.WriteLine(h.IsInvalid & "|" & h.IsClosed)
    End Sub
End Module
