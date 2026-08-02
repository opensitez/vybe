' vybe-test: vb/vb_system_array_matrix/array_createinstance_sets_and_reads_values
' origin: languages/vb/tests/vb/test_vb_system_array_matrix.rs

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

Module M
    Sub Main()
        Dim boxed As Array = Array.CreateInstance(GetType(Integer), 3)
        boxed.SetValue(7, 0)
        boxed.SetValue(8, 1)
        boxed.SetValue(9, 2)

        __Check(CStr(CInt(boxed.GetValue(0))), "7")
        __Check(CStr(CInt(boxed.GetValue(2))), "9")
        __Check(CStr(Array.IndexOf(CType(boxed, Integer()), 8)), "1")
    End Sub
End Module
