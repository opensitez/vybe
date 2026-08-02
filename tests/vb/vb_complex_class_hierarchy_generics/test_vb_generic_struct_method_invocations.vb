' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_generic_struct_method_invocations
' origin: languages/vb/tests/vb/test_vb_complex_class_hierarchy_generics.rs

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

Structure Pair(Of T1, T2)
    Public First As T1
    Public Second As T2
    Public Sub New(f As T1, s As T2)
        First = f
        Second = s
    End Sub
    Public Function Swap() As Pair(Of T2, T1)
        Return New Pair(Of T2, T1)(Second, First)
    End Function
End Structure

Module Program
    Sub Main()
        Dim p As New Pair(Of String, Integer)("Age", 30)
        Dim swapped = p.Swap()
        __Check(CStr(swapped.First & "=" & swapped.Second), "30=Age")
    End Sub
End Module
