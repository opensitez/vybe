' vybe-test: vb/vb_environment_get_environment_variable/test_vb_environment_variables_iterate_keys_values
' origin: languages/vb/tests/vb/test_vb_environment_get_environment_variable.rs

Imports System
Imports System.Collections

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("ITER_VAR", "IterVal")
        Dim dict = Environment.GetEnvironmentVariables()
        Dim found = False
        For Each de As DictionaryEntry In dict
            If de.Key.ToString() = "ITER_VAR" Then
                found = True
                Console.WriteLine(de.Key.ToString() & "=" & de.Value.ToString())
            End If
        Next
    End Sub
End Module
