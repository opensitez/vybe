use super::helpers::run_vb;

#[test]
fn environment_newline_and_version_are_available() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Environment.NewLine.Length > 0)
        Console.WriteLine(Environment.Version.Major >= 1)
        Console.WriteLine(Environment.Version.Revision >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn environment_processor_count_is_positive() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Environment.ProcessorCount > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn environment_current_directory_exists() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim cwd As String = Environment.CurrentDirectory
        Console.WriteLine(cwd.Length > 0)
        Console.WriteLine(Environment.CurrentDirectory = cwd)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn environment_tickcount_progresses() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim first As Integer = Environment.TickCount
        Thread.Sleep(5)
        Dim second As Integer = Environment.TickCount
        Console.WriteLine(second >= first)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn environment_system_directory_exists() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Environment.SystemDirectory.Length > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn environment_folder_paths_return_values() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim docs As String = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
        Dim tmp As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Console.WriteLine(docs.Length > 0)
        Console.WriteLine(tmp.Length > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn environment_command_line_arg_count() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim args() As String = Environment.GetCommandLineArgs()
        Console.WriteLine(args.Length >= 1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn environment_missing_variable_is_null() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim value As String = Environment.GetEnvironmentVariable("VB_TEST_ENV_MISSING_123456")
        Console.WriteLine(value Is Nothing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn environment_path_variable_can_be_read() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim pathValue As String = Environment.GetEnvironmentVariable("PATH")
        Console.WriteLine(pathValue IsNot Nothing)
        Console.WriteLine(pathValue.Length > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn environment_exit_code_roundtrip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Environment.ExitCode = 123
        Console.WriteLine(Environment.ExitCode = 123)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn environment_user_and_machine_names_are_populated() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Console.WriteLine(Not String.IsNullOrWhiteSpace(Environment.UserName))
        Console.WriteLine(Not String.IsNullOrWhiteSpace(Environment.MachineName))
        Console.WriteLine(Not String.IsNullOrWhiteSpace(Environment.OSVersion.ToString()))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn environment_can_set_and_clear_environment_variable() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim key As String = "VYBE_VB_ENV_SURFACE_2026"
        Environment.SetEnvironmentVariable(key, "active")

        Dim roundTrip As String = Environment.GetEnvironmentVariable(key)
        Console.WriteLine(roundTrip = "active")

        Environment.SetEnvironmentVariable(key, Nothing)
        Console.WriteLine(Environment.GetEnvironmentVariable(key) Is Nothing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn environment_temp_path_is_available() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim tempPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        Console.WriteLine(tempPath.Length > 0)
        Console.WriteLine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData).Length > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}
