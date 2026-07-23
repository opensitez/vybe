use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Environment Variables, Expansion & Target Scopes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_environment_get_set_environment_variable_process() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("VYBE_TEST_VAR", "VybeValue", EnvironmentVariableTarget.Process)
        Dim val = Environment.GetEnvironmentVariable("VYBE_TEST_VAR", EnvironmentVariableTarget.Process)
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["VybeValue"]);
}

#[test]
fn test_vb_environment_clear_environment_variable() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("VYBE_TEMP_VAR", "Data")
        Environment.SetEnvironmentVariable("VYBE_TEMP_VAR", Nothing)
        Dim val = Environment.GetEnvironmentVariable("VYBE_TEMP_VAR")
        Console.WriteLine(val Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_get_environment_variables_dictionary() {
    let src = r#"
Imports System
Imports System.Collections

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("VYBE_DICT_VAR", "DictVal")
        Dim dict As IDictionary = Environment.GetEnvironmentVariables()
        Console.WriteLine(dict.Contains("VYBE_DICT_VAR") & "|" & dict("VYBE_DICT_VAR"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|DictVal"]);
}

#[test]
fn test_vb_environment_expand_environment_variables() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("MY_HOST", "localhost")
        Dim expanded = Environment.ExpandEnvironmentVariables("http://%MY_HOST%:8080/api")
        Console.WriteLine(expanded)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["http://localhost:8080/api"]);
}

#[test]
fn test_vb_environment_path_variable_lookup() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim path = Environment.GetEnvironmentVariable("PATH")
        Console.WriteLine(path IsNot Nothing AndAlso path.Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_variable_case_insensitive_lookup_windows_unix() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("VYBE_CASE_VAR", "CaseVal")
        Dim val = Environment.GetEnvironmentVariable("VYBE_CASE_VAR")
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CaseVal"]);
}

#[test]
fn test_vb_environment_variable_null_or_empty_key_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Environment.GetEnvironmentVariable("")
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException on Empty Variable Key Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentException on Empty Variable Key Caught"]
    );
}

#[test]
fn test_vb_environment_variables_target_process_isolation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("PROC_VAR", "ProcVal", EnvironmentVariableTarget.Process)
        Dim vars = Environment.GetEnvironmentVariables(EnvironmentVariableTarget.Process)
        Console.WriteLine(vars("PROC_VAR"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ProcVal"]);
}

#[test]
fn test_vb_environment_variable_overwrite_existing() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("OVERWRITE_VAR", "V1")
        Environment.SetEnvironmentVariable("OVERWRITE_VAR", "V2")
        Console.WriteLine(Environment.GetEnvironmentVariable("OVERWRITE_VAR"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["V2"]);
}

#[test]
fn test_vb_environment_expand_unresolved_variable_remains() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim unexpanded = Environment.ExpandEnvironmentVariables("Prefix_%NON_EXISTENT_VAR_12345%_Suffix")
        Console.WriteLine(unexpanded)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Prefix_%NON_EXISTENT_VAR_12345%_Suffix"]);
}

#[test]
fn test_vb_environment_variable_special_characters_in_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim specialVal = "Line1;" & vbTab & "Line2=Val"
        Environment.SetEnvironmentVariable("SPECIAL_VAR", specialVal)
        Dim val = Environment.GetEnvironmentVariable("SPECIAL_VAR")
        Console.WriteLine(val = specialVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_variables_iterate_keys_values() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["ITER_VAR=IterVal"]);
}

#[test]
fn test_vb_environment_variable_with_spaces_in_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("SPACE_VAR", "Hello World From Vybe")
        Console.WriteLine(Environment.GetEnvironmentVariable("SPACE_VAR"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World From Vybe"]);
}

#[test]
fn test_vb_environment_variable_empty_string_value_removes_variable() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("EMPTY_VAL_VAR", "Content")
        Environment.SetEnvironmentVariable("EMPTY_VAL_VAR", "")
        Dim val = Environment.GetEnvironmentVariable("EMPTY_VAL_VAR")
        Console.WriteLine(val Is Nothing OrElse val = "")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_variable_long_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim longVal = New String("A"c, 2048)
        Environment.SetEnvironmentVariable("LONG_VAR", longVal)
        Dim val = Environment.GetEnvironmentVariable("LONG_VAR")
        Console.WriteLine(val.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2048"]);
}

#[test]
fn test_vb_environment_user_home_directory_variable() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim home = Environment.GetEnvironmentVariable("HOME")
        If home Is Nothing Then home = Environment.GetEnvironmentVariable("USERPROFILE")
        Console.WriteLine(home IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_expand_multiple_variables() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("VAR_A", "Alpha")
        Environment.SetEnvironmentVariable("VAR_B", "Beta")
        Dim res = Environment.ExpandEnvironmentVariables("%VAR_A% and %VAR_B%")
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Alpha and Beta"]);
}

#[test]
fn test_vb_environment_variable_numeric_value_parsing() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("PORT_NUM", "9090")
        Dim portStr = Environment.GetEnvironmentVariable("PORT_NUM")
        Dim port As Integer = Integer.Parse(portStr)
        Console.WriteLine(port)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["9090"]);
}

#[test]
fn test_vb_environment_variable_boolean_flag_parsing() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("ENABLE_FEATURE", "True")
        Dim flagStr = Environment.GetEnvironmentVariable("ENABLE_FEATURE")
        Dim enabled As Boolean = Boolean.Parse(flagStr)
        Console.WriteLine(enabled)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_environment_variable_get_non_existent_returns_null() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim val = Environment.GetEnvironmentVariable("DEFINITELY_NON_EXISTENT_VAR_XYZ")
        Console.WriteLine(val Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
