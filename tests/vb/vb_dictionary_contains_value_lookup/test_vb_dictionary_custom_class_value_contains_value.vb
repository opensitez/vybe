' vybe-test: vb/vb_dictionary_contains_value_lookup/test_vb_dictionary_custom_class_value_contains_value
' origin: languages/vb/tests/vb/test_vb_dictionary_contains_value_lookup.rs

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

Class Element
    Public Property Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Overrides Function Equals(obj As Object) As Boolean
        Dim e = TryCast(obj, Element)
        Return e IsNot Nothing AndAlso e.Name = Me.Name
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return Name.GetHashCode()
    End Function
End Class

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of Integer, Element) From {
            {1, New Element("Gold")},
            {2, New Element("Silver")}
        }
        __Check(CStr(dict.ContainsValue(New Element("Gold"))), "True")
    End Sub
End Module
