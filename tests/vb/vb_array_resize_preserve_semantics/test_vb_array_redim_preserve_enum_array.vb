' vybe-test: vb/vb_array_resize_preserve_semantics/test_vb_array_redim_preserve_enum_array
' origin: languages/vb/tests/vb/test_vb_array_resize_preserve_semantics.rs

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

Enum Priority
    Low = 0
    Medium = 1
    High = 2
End Enum

Module Program
    Sub Main()
        Dim priorities(0) As Priority
        priorities(0) = Priority.High
        ReDim Preserve priorities(1)
        __Check(CStr(priorities(0) & ":" & priorities(1)), "High:Low")
    End Sub
End Module
