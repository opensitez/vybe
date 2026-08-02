' vybe-test: vb/vb_key_value_pair_struct_usage/test_vb_key_value_pair_struct_key_and_value
' origin: languages/vb/tests/vb/test_vb_key_value_pair_struct_usage.rs

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
        Dim kv As New KeyValuePair(Of Point, Point)(New Point(0, 0), New Point(10, 20))
        __Check(CStr(kv.Key.X & " to " & kv.Value.X & "," & kv.Value.Y), "0 to 10,20")
    End Sub
End Module
