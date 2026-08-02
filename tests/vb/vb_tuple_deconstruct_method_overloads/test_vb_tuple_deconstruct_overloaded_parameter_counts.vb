' vybe-test: vb/vb_tuple_deconstruct_method_overloads/test_vb_tuple_deconstruct_overloaded_parameter_counts
' origin: languages/vb/tests/vb/test_vb_tuple_deconstruct_method_overloads.rs

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

Class DateTimeInfo
    Public Sub Deconstruct(ByRef year As Integer, ByRef month As Integer)
        year = 2025 : month = 12
    End Sub
    Public Sub Deconstruct(ByRef year As Integer, ByRef month As Integer, ByRef day As Integer)
        year = 2025 : month = 12 : day = 31
    End Sub
End Class

Module Program
    Sub Main()
        Dim info As New DateTimeInfo()
        Dim y As Integer = 0, m As Integer = 0, d As Integer = 0
        info.Deconstruct(y, m)
        __Check(CStr(y & "-" & m), "2025-12")
        info.Deconstruct(y, m, d)
        __Check(CStr(y & "-" & m & "-" & d), "2025-12-31")
    End Sub
End Module
