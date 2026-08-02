' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_readonly_property
' origin: languages/vb/tests/vb/test_vb_generic_struct_methods.rs

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

Structure OptionVal(Of T)
    Private _val As T
    Private _hasVal As Boolean
    Public ReadOnly Property Value As T
        Get
            Return _val
        End Get
    End Property
    Public ReadOnly Property HasValue As Boolean
        Get
            Return _hasVal
        End Get
    End Property
    Public Sub New(val As T)
        _val = val
        _hasVal = True
    End Sub
End Structure

Module Program
    Sub Main()
        Dim opt As New OptionVal(Of String)("Hello")
        __Check(CStr(opt.HasValue & "|" & opt.Value), "True|Hello")
    End Sub
End Module
