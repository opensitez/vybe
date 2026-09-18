' vybe-test: vb/vb_type_conversion_coercion/narrow_int_byte_err
' origin: languages/vb/tests/vb/test_vb_type_conversion_coercion.rs

Option Strict On

Module M
    Sub Main()
        ' ⛔ A `:` does NOT end a VB comment — writing the rest of the method
        ' after one commented away `End Sub` and `End Module` (BC30625).
        ' `Dim b As Byte = 1000` is the narrowing this covers: under
        ' `Option Strict On` it is a COMPILE error, so it stays commented and
        ' the runtime path is what gets checked.
        Console.WriteLine("Parsed")
    End Sub
End Module
