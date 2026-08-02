' vybe-test: vb/vb_modules_extension_methods/module_extension_methods
' origin: languages/vb/tests/vb/test_vb_modules_extension_methods.rs

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

Module StringExtensions
    <Extension()>
    Public Function WordCount(str As String) As Integer
        Return str.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries).Length
    End Function
End Module

Module M
    Sub Main()
        Dim text As String = "Hello world from VB.NET"
        ' Calling extension method like an instance method
        __Check(CStr(text.WordCount()), "4")
    End Sub
End Module
