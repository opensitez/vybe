' vybe-test: vb/vb_array_empty_and_null_bounds/test_vb_array_create_instance_type_length
' origin: languages/vb/tests/vb/test_vb_array_empty_and_null_bounds.rs

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

Imports System

Module Program
    Sub Main()
        Dim arr As Array = Array.CreateInstance(GetType(String), 3)
        arr.SetValue("Item0", 0)
        arr.SetValue("Item1", 1)
        __Check(CStr(arr.Length), "3")
        __Check(CStr(arr.GetValue(0) & "," & arr.GetValue(1)), "Item0,Item1")
    End Sub
End Module
