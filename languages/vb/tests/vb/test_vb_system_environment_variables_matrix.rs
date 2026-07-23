use super::helpers::run_vb;

#[test]
fn environment_variables_set_get_and_clear_roundtrip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim key As String = "VYBE_VB_ENV_MATRIX"
        Environment.SetEnvironmentVariable(key, "present")
        Console.WriteLine(Environment.GetEnvironmentVariable(key) = "present")

        Environment.SetEnvironmentVariable(key, Nothing)
        Console.WriteLine(Environment.GetEnvironmentVariable(key) Is Nothing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn environment_variables_expand_tokens_and_command_line() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim expanded As String = Environment.ExpandEnvironmentVariables("prefix_%PATH%");
        Console.WriteLine(expanded.Contains("prefix_"))
        Dim args() As String = Environment.GetCommandLineArgs()
        Console.WriteLine(args.Length >= 1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn environment_variable_target_machine_and_process_set_supported() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim key As String = "VYBE_VB_ENV_MACHINE_TEST"
        Environment.SetEnvironmentVariable(key, "1", EnvironmentVariableTarget.Process)

        Dim processValue As String = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.Process)
        Console.WriteLine(processValue = "1")

        Dim machineValue As String = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine)
        Console.WriteLine(Not String.IsNullOrWhiteSpace(machineValue))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn environment_variables_map_contains_common_keys() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim values As Collections.IDictionary = Environment.GetEnvironmentVariables()
        Console.WriteLine(values IsNot Nothing)
        Console.WriteLine(values.Count > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}
