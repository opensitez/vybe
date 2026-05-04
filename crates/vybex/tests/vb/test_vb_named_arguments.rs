use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_named_arguments(name: &str, prefix: &str, suffix: &str) {
    let src = format!(r#"
Module M
    Function Describe(name As String, prefix As String, suffix As String) As String
        Return prefix & ":" & name & ":" & suffix
    End Function

    Sub Main()
        Console.WriteLine(Describe(suffix:="{suffix}", name:="{name}", prefix:="{prefix}"))
    End Sub
End Module
"#, name = name, prefix = prefix, suffix = suffix);
    assert_vb_output_owned(src, vec![format!("{}:{}:{}", prefix, name, suffix)]);
}

macro_rules! named_argument_cases {
    ($($name:ident => ($person:expr, $prefix:expr, $suffix:expr)),* $(,)?) => {
        $(#[test] fn $name() { assert_named_arguments($person, $prefix, $suffix); })*
    };
}

named_argument_cases! {
    named_arguments_001 => ("Ada", "Hi", "!"),
    named_arguments_002 => ("Bob", "Hello", "?"),
    named_arguments_003 => ("Cora", "Start", "."),
    named_arguments_004 => ("Dax", "Outer", "!"),
    named_arguments_005 => ("Eli", "North", "*"),
    named_arguments_006 => ("Faye", "Red", "!"),
    named_arguments_007 => ("Gus", "Open", "?"),
    named_arguments_008 => ("Hope", "Fast", "."),
    named_arguments_009 => ("Ivy", "Warm", "!"),
    named_arguments_010 => ("Jade", "Early", "?"),
}