' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_tuple_key
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

Imports System.Collections.Generic

Class TupleMap
    Private dict As New Dictionary(Of (Integer, Integer), String)()
    Default Public Property Item(r As Integer, c As Integer) As String
        Get
            If dict.ContainsKey((r, c)) Then Return dict((r, c))
            Return Nothing
        End Get
        Set(value As String)
            dict((r, c)) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim tm As New TupleMap()
        tm(3, 4) = "Position34"
        __Check(CStr(tm(3, 4)), "Position34")
    End Sub
End Module
