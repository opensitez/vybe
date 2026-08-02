' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_optional_parameters
' origin: languages/vb/tests/vb/test_vb_property_indexed_multi_parameter.rs

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

Class OptionalIndexer
    Public Property Element(row As Integer, Optional col As Integer = 0) As String
        Get
            Return "R=" & row & ",C=" & col
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim o As New OptionalIndexer()
        __Check(CStr(o.Element(5) & "|" & o.Element(5, 3)), "R=5,C=0|R=5,C=3")
    End Sub
End Module
