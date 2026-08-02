' vybe-test: vb/vb_spec_object_model/object_model_spec_structure_method_reads_state
' origin: languages/vb/tests/vb/test_vb_spec_object_model.rs

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

Structure Point
    Public X As Integer
    Public Function DoubleX() As Integer
        Return X * 2
    End Function
End Structure
Module M
    Sub Main()
        Dim p As Point
        p.X = 6
        __Check(CStr(p.DoubleX()), "12")
    End Sub
End Module
