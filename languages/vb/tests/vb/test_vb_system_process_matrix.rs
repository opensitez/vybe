use super::helpers::run_vb;

#[test]
fn process_current_process_exists_and_has_name() {
    let out = run_vb(
        r#"
Imports System
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Console.WriteLine(p.Id > 0)
        Console.WriteLine(p.ProcessName.Length > 0)
        Console.WriteLine(Not String.IsNullOrWhiteSpace(p.ProcessName))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn process_self_is_in_process_list_by_name() {
    let out = run_vb(
        r#"
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
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn process_current_start_time_is_in_past() {
    let out = run_vb(
        r#"
Imports System
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Dim started As DateTime = p.StartTime

        Console.WriteLine(started <= DateTime.Now)
        Console.WriteLine(p.Threads.Count > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn process_has_modules_and_threads() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Console.WriteLine(p.Modules.Count >= 1)
        Console.WriteLine(p.Threads.Count >= 1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn process_not_exited_after_creation() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Console.WriteLine(Not p.HasExited)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn process_wait_for_exit_with_timeout_for_current_is_false() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Console.WriteLine(Not p.WaitForExit(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn process_handle_is_valid() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Console.WriteLine(p.Handle <> IntPtr.Zero)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn process_worker_id_stability_check() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim first As Process = Process.GetCurrentProcess()
        Dim second As Process = Process.GetProcessById(first.Id)

        Console.WriteLine(first.Id = second.Id)
        Console.WriteLine(first.ProcessName = second.ProcessName)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn process_priority_class_is_defined() {
    let out = run_vb(
        r#"
Imports System
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Dim priorityValue As Object = p.PriorityClass

        Console.WriteLine(System.Enum.IsDefined(GetType(ProcessPriorityClass), priorityValue))
        Console.WriteLine(priorityValue IsNot Nothing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn process_working_set_is_non_negative() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Console.WriteLine(p.WorkingSet64 >= 0)
        Console.WriteLine(p.PeakWorkingSet64 >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn process_total_cpu_time_is_non_negative() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Console.WriteLine(p.TotalProcessorTime.Ticks >= 0)
        Console.WriteLine(p.UserProcessorTime.Ticks >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn process_main_module_file_name_is_present() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Console.WriteLine(p.MainModule.ModuleName.Length > 0)
        Console.WriteLine(p.MainModule.FileName.Length > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn process_memory_info_is_retrievable() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()

        Console.WriteLine(p.VirtualMemorySize64 >= 0)
        Console.WriteLine(p.PrivateMemorySize64 >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn process_main_module_can_be_read() {
    let out = run_vb(
        r#"
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
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn process_lookup_missing_id_throws() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim threw As Boolean = False

        Try
            Process.GetProcessById(-1)
        Catch ex As Exception
            threw = True
        End Try

        Console.WriteLine(threw)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
