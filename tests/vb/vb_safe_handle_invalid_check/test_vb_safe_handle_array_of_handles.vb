' vybe-test: vb/vb_safe_handle_invalid_check/test_vb_safe_handle_array_of_handles
' origin: languages/vb/tests/vb/test_vb_safe_handle_invalid_check.rs

Imports System
Imports System.Runtime.InteropServices

Class ArrayItemHandle
    Inherits SafeHandle
    Public Sub New(val As Long)
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(val))
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
        Dim handles As ArrayItemHandle() = {New ArrayItemHandle(1), New ArrayItemHandle(2)}
        For Each h In handles
            h.Dispose()
        Next
        Console.WriteLine("All Array Handles Disposed")
    End Sub
End Module
