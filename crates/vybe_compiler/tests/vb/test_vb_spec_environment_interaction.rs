use super::helpers::{run_vb, run_vb_gui};

macro_rules! vb_expr_spec {
    ($name:ident, $body:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let output = run_vb($body);
            assert_eq!(output, vec![$expected.to_string()]);
        }
    };
}

vb_expr_spec!(
    environment_spec_current_directory_is_non_empty,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(Len(Environment.CurrentDirectory) > 0)
    End Sub
End Module
"#,
    "True"
);

vb_expr_spec!(
    environment_spec_newline_is_string,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(TypeName(Environment.NewLine))
    End Sub
End Module
"#,
    "String"
);

vb_expr_spec!(
    environment_spec_machine_name_is_string,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(TypeName(Environment.MachineName))
    End Sub
End Module
"#,
    "String"
);

vb_expr_spec!(
    environment_spec_user_name_is_string,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(TypeName(Environment.UserName))
    End Sub
End Module
"#,
    "String"
);

vb_expr_spec!(
    environment_spec_os_version_is_string,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(TypeName(Environment.OSVersion))
    End Sub
End Module
"#,
    "String"
);

vb_expr_spec!(
    environment_spec_processor_count_is_integer,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(TypeName(Environment.ProcessorCount))
    End Sub
End Module
"#,
    "Integer"
);

vb_expr_spec!(
    environment_spec_tick_count_is_integer,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(TypeName(Environment.TickCount))
    End Sub
End Module
"#,
    "Integer"
);

vb_expr_spec!(
    environment_spec_get_environment_variable_reads_host_value,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(Len(Environment.GetEnvironmentVariable("HOME")) > 0)
    End Sub
End Module
"#,
    "True"
);

vb_expr_spec!(
    environment_spec_set_environment_variable_roundtrip,
    r#"
Module Program
    Sub Main()
        Environment.SetEnvironmentVariable("VYBEX_VB_ENV_TEST", "present")
        Console.WriteLine(Environment.GetEnvironmentVariable("VYBEX_VB_ENV_TEST"))
    End Sub
End Module
"#,
    "present"
);

vb_expr_spec!(
    environment_spec_get_folder_path_returns_string,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(Len(Environment.GetFolderPath(0)) > 0)
    End Sub
End Module
"#,
    "True"
);

vb_expr_spec!(
    interaction_spec_beep_returns_to_program_flow,
    r#"
Module Program
    Sub Main()
        Beep()
        Console.WriteLine("after-beep")
    End Sub
End Module
"#,
    "after-beep"
);

vb_expr_spec!(
    interaction_spec_shell_returns_process_identifier,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(Shell("/bin/echo vb-shell") > 0)
    End Sub
End Module
"#,
    "True"
);

vb_expr_spec!(
    interaction_spec_app_path_is_non_empty,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(Len(App.Path) > 0)
    End Sub
End Module
"#,
    "True"
);

vb_expr_spec!(
    interaction_spec_app_title_is_string,
    r#"
Module Program
    Sub Main()
        Console.WriteLine(TypeName(App.Title))
    End Sub
End Module
"#,
    "String"
);

#[test]
fn interaction_spec_screen_dimensions_are_positive() {
    let (_vm, _gui, output) = run_vb_gui(
        r#"
Module Program
    Sub Main()
        Console.WriteLine(Screen.Width > 0)
        Console.WriteLine(Screen.Height > 0)
    End Sub
End Module
"#,
    );

    let output = output.lock().unwrap().clone();
    assert_eq!(output, vec!["True", "True"]);
}
