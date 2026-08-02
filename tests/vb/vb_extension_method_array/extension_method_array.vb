' vybe-test: vb/vb_extension_method_array/extension_method_array
' origin: languages/vb/tests/vb/test_vb_extension_method_array.rs

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

Module Extensions
    <Extension()>
    Public Function SumFirstTwo(arr() As Integer) As Integer
        Return arr(0) + arr(1)
    End Function
End Module

Module M
    Sub Main()
        Dim nums() As Integer = {5, 10, 15}
        __Check(CStr(nums.SumFirstTwo()), "15")
    End Sub
End Module
