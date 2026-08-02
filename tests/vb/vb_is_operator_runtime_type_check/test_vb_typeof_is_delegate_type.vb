' vybe-test: vb/vb_is_operator_runtime_type_check/test_vb_typeof_is_delegate_type
' origin: languages/vb/tests/vb/test_vb_is_operator_runtime_type_check.rs

Imports System

Module Program
    Sub Main()
        Dim act As Object = CType(Sub() Console.WriteLine("Hi"), Action)
        Console.WriteLine(TypeOf act Is Action & "|" & TypeOf act Is Delegate)
    End Sub
End Module
