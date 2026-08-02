' vybe-test: vb/vb_comprehensive/property_get_set_basic
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

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

Module M
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
    End Class

    Sub Main()
        Dim t As New Temperature(100)
        __Check(CStr(t.Celsius), "100")
        t.Celsius = 0
        __Check(CStr(t.Celsius), "0")
    End Sub
End Module
