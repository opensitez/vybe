' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_method_info_extension_method_check
' origin: languages/vb/tests/vb/test_vb_reflection_method_info_generic_invoke.rs

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

Imports System.Reflection
Imports System.Runtime.CompilerServices

Module StringExtensions
    <Extension()>
    Public Function ReverseString(s As String) As String
        Dim chars = s.ToCharArray()
        Array.Reverse(chars)
        Return New String(chars)
    End Function
End Module

Module Program
    Sub Main()
        Dim m = GetType(StringExtensions).GetMethod("ReverseString")
        Dim isExt = m.IsDefined(GetType(ExtensionAttribute), False)
        __Check(CStr(isExt), "True")
    End Sub
End Module
