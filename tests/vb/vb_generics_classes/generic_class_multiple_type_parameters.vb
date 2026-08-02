' vybe-test: vb/vb_generics_classes/generic_class_multiple_type_parameters
' origin: languages/vb/tests/vb/test_vb_generics_classes.rs

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

Class Pair(Of T1, T2)
    Public First As T1
    Public Second As T2
    
    Public Sub New(f As T1, s As T2)
        First = f
        Second = s
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Pair(Of String, Integer)("Age", 30)
        __Check(CStr(p.First), "Age")
        __Check(CStr(p.Second), "30")
    End Sub
End Module
