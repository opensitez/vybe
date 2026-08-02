' vybe-test: vb/vb_tuple_deconstruct_method_overloads/test_vb_tuple_generic_deconstruct
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

Class Container(Of T1, T2)
    Public V1 As T1
    Public V2 As T2
    Public Sub New(v1 As T1, v2 As T2) : Me.V1 = v1 : Me.V2 = v2 : End Sub
    Public Sub Deconstruct(ByRef out1 As T1, ByRef out2 As T2)
        out1 = V1 : out2 = V2
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Container(Of String, Double)("PI", 3.14)
        Dim k As String = Nothing
        Dim v As Double = 0.0
        c.Deconstruct(k, v)
        __Check(CStr(k & "=" & v), "PI=3.14")
    End Sub
End Module
