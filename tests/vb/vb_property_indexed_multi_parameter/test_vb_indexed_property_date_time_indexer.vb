' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_date_time_indexer
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

Imports System
Imports System.Collections.Generic

Class Schedule
    Private events As New Dictionary(Of DateTime, String)()
    Default Public Property EventName(dt As DateTime) As String
        Get
            If events.ContainsKey(dt) Then Return events(dt)
            Return "Free"
        Get
        End Get
        Set(value As String)
            events(dt) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim s As New Schedule()
        Dim dt = New DateTime(2025, 6, 1)
        s(dt) = "Conference"
        __Check(CStr(s(dt) & "|" & s(dt.AddDays(1))), "Conference|Free")
    End Sub
End Module
