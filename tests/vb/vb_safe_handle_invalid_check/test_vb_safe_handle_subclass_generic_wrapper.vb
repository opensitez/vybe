' vybe-test: vb/vb_safe_handle_invalid_check/test_vb_safe_handle_subclass_generic_wrapper
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

Class GenericSafeHandle(Of T)
    Inherits SafeHandle

    Public Sub New(initialPtr As IntPtr)
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(initialPtr)
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        __Check(CStr("Generic Release: " & GetType(T).Name), "Generic Release: String")
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New GenericSafeHandle(Of String)(New IntPtr(100))
        h.Dispose()
    End Sub
End Module
