' vybe-test: vb/vb_system_exception_types_matrix/exception_key_not_found_is_thrown_for_missing_dictionary_key
' origin: languages/vb/tests/vb/test_vb_system_exception_types_matrix.rs

Imports System
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of String, Integer)()
        map("present") = 11

        Try
            Console.WriteLine(map("missing"))
        Catch ex As KeyNotFoundException
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
