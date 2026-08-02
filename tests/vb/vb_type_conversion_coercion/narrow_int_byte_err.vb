' vybe-test: vb/vb_type_conversion_coercion/narrow_int_byte_err
' origin: languages/vb/tests/vb/test_vb_type_conversion_coercion.rs

Option Strict On: Module M: Sub Main(): ' Dim b As Byte = 1000 ' Compile Error with Option Strict: Console.WriteLine("Parsed"): End Sub: End Module
