' vybe-test: vb/vb_modules_namespaces/mod_extension_methods_require_imports
' origin: languages/vb/tests/vb/test_vb_modules_namespaces.rs

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
Namespace N1
Public Module Mod1
<Extension()>
Public Function DoubleIt(v As Integer) As Integer
Return v * 2
End Function
End Module
End Namespace
Imports N1 ' Required to use extension methods in Mod1
Module M
Sub Main()
Dim x = 5
__Check(CStr(x.DoubleIt()), "10")
End Sub
End Module
