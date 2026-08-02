' vybe-test: vb/vb_safe_handle_invalid_check/test_vb_safe_handle_set_handle_as_invalid
' origin: languages/vb/tests/vb/test_vb_safe_handle_invalid_check.rs

Imports System
Imports System.Runtime.InteropServices

Class ResettableSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(100))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Public Sub Invalidate()
        SetHandleAsInvalid()
    End Sub

    Protected Overrides Function ReleaseHandle() As Boolean
        Console.WriteLine("Released")
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New ResettableSafeHandle()
        h.Invalidate()
        Console.WriteLine(h.IsClosed & "|" & h.IsInvalid)
    End Sub
End Module
