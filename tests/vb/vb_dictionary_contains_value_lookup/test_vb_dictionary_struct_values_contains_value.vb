' vybe-test: vb/vb_dictionary_contains_value_lookup/test_vb_dictionary_struct_values_contains_value
' origin: languages/vb/tests/vb/test_vb_dictionary_contains_value_lookup.rs

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

Imports System.Collections.Generic

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
End Structure

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Point) From {
            {"P1", New Point(1, 2)},
            {"P2", New Point(3, 4)}
        }
        __Check(CStr(dict.ContainsValue(New Point(3, 4))), "True")
        __Check(CStr(dict.ContainsValue(New Point(0, 0))), "False")
    End Sub
End Module
