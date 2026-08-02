' vybe-test: vb/vb_system_process_matrix/process_main_module_can_be_read
' origin: languages/vb/tests/vb/test_vb_system_process_matrix.rs

Imports System
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Dim hasMainModule As Boolean = True

        Try
            Console.WriteLine(p.MainModule.ModuleName.Length > 0)
        Catch ex As Exception
            hasMainModule = False
            Console.WriteLine(hasMainModule)
        End Try
    End Sub
End Module
