' vybe-test: vb/vb_system_extension_method_matrix/extension_method_string_transforms_are_visible
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

Module TextExtensions
    <Extension()>
    Public Function Wrap(value As String, prefix As String, suffix As String) As String
        Return prefix & value & suffix
    End Function

    <Extension()>
    Public Function ReverseText(value As String) As String
        Dim chars As Char() = value.ToCharArray()
        Array.Reverse(chars)
        Return New String(chars)
    End Function
End Module

Module M
    Sub Main()
        Dim original As String = "abc"
        __Check(CStr(original.Wrap("[", "]")), "[abc]")
        __Check(CStr(original.ReverseText()), "cba")
    End Module
End Module
