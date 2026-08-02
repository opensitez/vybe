' vybe-test: vb/vb_interface_explicit_implementation_adv/test_vb_interface_multiple_same_name_members
' origin: languages/vb/tests/vb/test_vb_interface_explicit_implementation_adv.rs

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

Interface IControl
    Sub Paint()
End Interface

Interface ISurface
    Sub Paint()
End Interface

Class Canvas
    Implements IControl, ISurface

    Private Sub PaintControl() Implements IControl.Paint
        __Check(CStr("Control Paint"), "Control Paint")
    End Sub

    Private Sub PaintSurface() Implements ISurface.Paint
        __Check(CStr("Surface Paint"), "Surface Paint")
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Canvas()
        Dim ctrl As IControl = c
        Dim surf As ISurface = c
        ctrl.Paint()
        surf.Paint()
    End Sub
End Module
