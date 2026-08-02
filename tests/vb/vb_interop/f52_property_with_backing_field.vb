' vybe-test: vb/vb_interop/f52_property_with_backing_field
' origin: languages/vb/tests/vb/vb_interop_test.rs

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

Public Class Temperature
    Dim _celsius As Double
    Public Property Celsius As Double
        Get
            Return _celsius
        End Get
        Set(value As Double)
            _celsius = value
        End Set
    End Property
    Public Function GetFahrenheit() As Double
        Return _celsius * 9 / 5 + 32
    End Function
    Public Sub New(c As Double)
        _celsius = c
    End Sub
End Class
Dim t As New Temperature(100)
__Check(CStr(t.Celsius), "100")
__Check(CStr(t.GetFahrenheit()), "212")
t.Celsius = 0
__Check(CStr(t.GetFahrenheit()), "32")
