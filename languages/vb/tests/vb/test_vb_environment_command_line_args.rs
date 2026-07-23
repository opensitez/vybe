use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Environment Command Line Arguments & Process Properties
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_environment_get_command_line_args() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim args = Environment.GetCommandLineArgs()
        Console.WriteLine(args.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_command_line_string() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim cmd = Environment.CommandLine
        Console.WriteLine(cmd.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_processor_count() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim procCount = Environment.ProcessorCount
        Console.WriteLine(procCount >= 1)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_is_64bit_operating_system() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim is64Os = Environment.Is64BitOperatingSystem
        Dim is64Proc = Environment.Is64BitProcess
        Console.WriteLine(is64Os & "|" & is64Proc)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_environment_machine_name() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim name = Environment.MachineName
        Console.WriteLine(name.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_user_domain_name() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim domain = Environment.UserDomainName
        Console.WriteLine(domain.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_user_name() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim user = Environment.UserName
        Console.WriteLine(user.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_os_version() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim os = Environment.OSVersion
        Console.WriteLine(os.Platform.ToString() & "|" & os.Version.Major > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Unix|True"]);
}

#[test]
fn test_vb_environment_current_managed_thread_id() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim threadId = Environment.CurrentManagedThreadId
        Console.WriteLine(threadId >= 1)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_tick_count_ms() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ticks = Environment.TickCount
        Console.WriteLine(ticks <> 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_tick_count_64bit() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ticks64 = Environment.TickCount64
        Console.WriteLine(ticks64 > 0L)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_system_directory() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim sysDir = Environment.SystemDirectory
        Console.WriteLine(sysDir IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_working_set_memory() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ws = Environment.WorkingSet
        Console.WriteLine(ws > 0L)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_new_line_string() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim nl = Environment.NewLine
        Console.WriteLine(nl.Length = 1 OrElse nl.Length = 2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_exit_code_get_set() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.ExitCode = 0
        Console.WriteLine(Environment.ExitCode)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_environment_get_logical_drives() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim drives = Environment.GetLogicalDrives()
        Console.WriteLine(drives.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_get_folder_path_special_folder() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Console.WriteLine(desktopPath IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_has_shutdown_started() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(Environment.HasShutdownStarted)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_environment_system_page_size() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim pageSize = Environment.SystemPageSize
        Console.WriteLine(pageSize >= 4096)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_version_dotnet_runtime() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ver = Environment.Version
        Console.WriteLine(ver.Major >= 6)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
