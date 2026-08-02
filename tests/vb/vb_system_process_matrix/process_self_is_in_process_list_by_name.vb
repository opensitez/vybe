' vybe-test: vb/vb_system_process_matrix/process_self_is_in_process_list_by_name
' origin: languages/vb/tests/vb/test_vb_system_process_matrix.rs

Imports System.Diagnostics

Module M
    Sub Main()
        Dim current As Process = Process.GetCurrentProcess()
        Dim found As Boolean = False

        For Each p As Process In Process.GetProcessesByName(current.ProcessName)
            If p.Id = current.Id Then
                found = True
            End If
        Next

        Console.WriteLine(found)
        Console.WriteLine(Process.GetProcesses().Length > 0)
    End Sub
End Module
