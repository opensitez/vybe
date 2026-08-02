' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_string_and_int_keys
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

Class MultiMap
    Private dict As New Dictionary(Of String, List(Of String))()
    Default Public Property Value(category As String, index As Integer) As String
        Get
            If dict.ContainsKey(category) AndAlso index < dict(category).Count Then
                Return dict(category)(index)
            End If
            Return Nothing
        End Get
        Set(val As String)
            If Not dict.ContainsKey(category) Then
                dict(category) = New List(Of String)()
            End If
            dict(category).Add(val)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim m As New MultiMap()
        m("Fruits", 0) = "Apple"
        m("Fruits", 1) = "Banana"
        __Check(CStr(m("Fruits", 0) & "|" & m("Fruits", 1)), "Apple|Banana")
    End Sub
End Module
