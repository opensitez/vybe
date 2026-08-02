' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_with_enum_type_parameter
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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

Enum Mode
    Standard
    Advanced
End Enum

Interface IConfig(Of TEnum As Structure)
    Property CurrentMode As TEnum
End Interface

Class ModeConfig
    Implements IConfig(Of Mode)
    Public Property CurrentMode As Mode Implements IConfig(Of Mode).CurrentMode = Mode.Advanced
End Class

Module Program
    Sub Main()
        Dim cfg As IConfig(Of Mode) = New ModeConfig()
        __Check(CStr(cfg.CurrentMode.ToString()), "Advanced")
    End Sub
End Module
