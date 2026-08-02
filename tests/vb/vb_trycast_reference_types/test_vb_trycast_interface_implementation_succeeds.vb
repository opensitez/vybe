' vybe-test: vb/vb_trycast_reference_types/test_vb_trycast_interface_implementation_succeeds
' origin: languages/vb/tests/vb/test_vb_trycast_reference_types.rs

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

Interface IPlayable
    Sub Play()
End Interface

Class Widget
    Implements IPlayable
    Public Sub Play() Implements IPlayable.Play
        __Check(CStr("Widget Playing"), "True")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As Object = New Widget()
        Dim p As IPlayable = TryCast(obj, IPlayable)
        __Check(CStr(p IsNot Nothing), "Widget Playing")
        If p IsNot Nothing Then p.Play()
    End Sub
End Module
