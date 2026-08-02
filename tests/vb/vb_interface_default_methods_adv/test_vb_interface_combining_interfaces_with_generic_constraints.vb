' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_combining_interfaces_with_generic_constraints
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

Interface IValidatable
    Function IsValid() As Boolean
End Interface

Class Processor(Of T As IValidatable)
    Public Function Process(item As T) As String
        If item.IsValid() Then
            Return "Valid"
        Else
            Return "Invalid"
        End If
    End Function
End Class

Class FormInput
    Implements IValidatable
    Public Input As String
    Public Sub New(i As String)
        Input = i
    End Sub
    Public Function IsValid() As Boolean Implements IValidatable.IsValid
        Return Not String.IsNullOrEmpty(Input)
    End Function
End Class

Module Program
    Sub Main()
        Dim p As New Processor(Of FormInput)()
        __Check(CStr(p.Process(New FormInput("OK"))), "Valid")
        __Check(CStr(p.Process(New FormInput(""))), "Invalid")
    End Sub
End Module
