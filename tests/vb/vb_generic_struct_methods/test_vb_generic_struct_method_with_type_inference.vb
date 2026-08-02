' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_method_with_type_inference
' origin: languages/vb/tests/vb/test_vb_generic_struct_methods.rs

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
    Public Item1 As T1
    Public Item2 As T2
    Public Sub New(i1 As T1, i2 As T2)
        Item1 = i1 : Item2 = i2
    End Sub
    Public Function Swap() As Pair(Of T2, T1)
        Return New Pair(Of T2, T1)(Item2, Item1)
    End Function
End Structure

Module Program
    Sub Main()
        Dim p As New Pair(Of String, Integer)("Age", 30)
        Dim s = p.Swap()
        __Check(CStr(s.Item1 & ":" & s.Item2), "30:Age")
    End Sub
End Module
