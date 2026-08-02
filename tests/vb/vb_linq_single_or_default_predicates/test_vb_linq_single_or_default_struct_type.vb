' vybe-test: vb/vb_linq_single_or_default_predicates/test_vb_linq_single_or_default_struct_type
' origin: languages/vb/tests/vb/test_vb_linq_single_or_default_predicates.rs

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

Structure Point
    Public X As Integer
    Public Y As Integer
End Structure

Module Program
    Sub Main()
        Dim empty As Point() = {}
        Dim pt = empty.SingleOrDefault()
        __Check(CStr(pt.X & "," & pt.Y), "0,0")
    End Sub
End Module
