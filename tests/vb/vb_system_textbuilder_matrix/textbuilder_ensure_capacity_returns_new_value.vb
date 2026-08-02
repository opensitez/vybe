' vybe-test: vb/vb_system_textbuilder_matrix/textbuilder_ensure_capacity_returns_new_value
' origin: languages/vb/tests/vb/test_vb_system_textbuilder_matrix.rs

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

Imports System.Text

Module M
    Sub Main()
        Dim sb As New StringBuilder("abc")
        Dim oldCapacity As Integer = sb.Capacity
        Dim newCapacity As Integer = sb.EnsureCapacity(20)
        __Check(CStr(newCapacity >= 20), "True")
        __Check(CStr(sb.Capacity >= 20), "True")
        __Check(CStr(sb.Capacity >= oldCapacity), "True")
    End Sub
End Module
