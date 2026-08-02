' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_struct_implements_interface
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IFormattableValue
    Function Format() As String
End Interface

Structure Currency
    Implements IFormattableValue
    Public Amount As Decimal
    Public Sub New(amt As Decimal)
        Amount = amt
    End Sub
    Public Function Format() As String Implements IFormattableValue.Format
        Return "$" & Amount
    End Function
End Structure

Module Program
    Sub Main()
        Dim c As IFormattableValue = New Currency(49.99D)
        __Check(CStr(c.Format()), "$49.99")
    End Sub
End Module
