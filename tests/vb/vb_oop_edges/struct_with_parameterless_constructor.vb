' vybe-test: vb/vb_oop_edges/struct_with_parameterless_constructor
' origin: languages/vb/tests/vb/test_vb_oop_edges.rs

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

Structure S
    Public Val As Integer
    ' Parameterless constructors in structs are allowed in VB 14+
    Public Sub New()
        Val = 42
    End Sub
End Structure

Module M
    Sub Main()
        Dim s As New S()
        __Check(CStr(s.Val), "42")
    End Sub
End Module
