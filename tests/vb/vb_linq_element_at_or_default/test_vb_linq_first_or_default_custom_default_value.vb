' vybe-test: vb/vb_linq_element_at_or_default/test_vb_linq_first_or_default_custom_default_value
' origin: languages/vb/tests/vb/test_vb_linq_element_at_or_default.rs

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

Imports System.Linq

Module Program
    Sub Main()
        Dim emptyList As New System.Collections.Generic.List(Of String)()
        Dim firstVal As String = emptyList.FirstOrDefault()
        __Check(CStr(firstVal Is Nothing), "True")
    End Sub
End Module
