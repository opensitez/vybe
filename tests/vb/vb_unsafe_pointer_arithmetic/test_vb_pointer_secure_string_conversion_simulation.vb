' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_secure_string_conversion_simulation
' origin: languages/vb/tests/vb/test_vb_unsafe_pointer_arithmetic.rs

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
Imports System.Security

Module Program
    Sub Main()
        Dim ss As New SecureString()
        ss.AppendChar("A"c)
        ss.AppendChar("B"c)
        Dim ptr As IntPtr = Marshal.SecureStringToGlobalAllocUnicode(ss)
        Dim str = Marshal.PtrToStringUni(ptr)
        Marshal.ZeroFreeGlobalAllocUnicode(ptr)
        ss.Dispose()
        __Check(CStr(str), "AB")
    End Sub
End Module
