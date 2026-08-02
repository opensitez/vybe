' vybe-test: vb/vb_system_extension_method_matrix/extension_method_generic_like_list_projection
' origin: languages/vb/tests/vb/test_vb_system_extension_method_matrix.rs

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

Imports System.Runtime.CompilerServices

Module EnumerableExtensions
    <Extension()>
    Public Function FirstOrDefaultIfExists(values As Integer(), fallback As Integer) As Integer
        If values Is Nothing OrElse values.Length = 0 Then
            Return fallback
        End If
        Return values(0)
    End Function
End Module

Module M
    Sub Main()
        Dim values As Integer() = {12, 24, 36}
        Dim empty As Integer() = {}

        __Check(CStr(values.FirstOrDefaultIfExists(99)), "12")
        __Check(CStr(empty.FirstOrDefaultIfExists(77)), "77")
    End Sub
End Module
