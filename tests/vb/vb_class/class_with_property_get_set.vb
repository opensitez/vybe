' vybe-test: vb/vb_class/class_with_property_get_set
' origin: languages/vb/tests/vb/vb_class_test.rs

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

Module Program
    Class Temperature
        Private _celsius As Double

        Sub New(c As Double)
            _celsius = c
        End Sub

        Property Celsius() As Double
            Get
                Return _celsius
            End Get
            Set(value As Double)
                _celsius = value
            End Set
        End Property

        Property Fahrenheit() As Double
            Get
                Return _celsius * 9 / 5 + 32
            End Get
            Set(value As Double)
                _celsius = (value - 32) * 5 / 9
            End Set
        End Property
    End Class

    Sub Main()
        Dim t As New Temperature(100)
        __Check(CStr(t.Celsius), "100")
        __Check(CStr(t.Fahrenheit), "212")
        t.Fahrenheit = 32
        __Check(CStr(t.Celsius), "0")
    End Sub
End Module
