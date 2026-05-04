use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_option_directives(directives: &str) {
    let src = format!(r#"
{directives}
Module M
    Sub Main()
        Dim total As Integer = 1
        total = total + 1
        Console.WriteLine(total)
    End Sub
End Module
"#, directives = directives);
    assert_vb_output_owned(src, vec!["2".to_string()]);
}

macro_rules! option_directive_cases {
    ($($name:ident => $directives:expr),* $(,)?) => {
        $(#[test] fn $name() { assert_option_directives($directives); })*
    };
}

option_directive_cases! {
    option_directives_001 => "Option Explicit On",
    option_directives_002 => "Option Explicit Off",
    option_directives_003 => "Option Strict On",
    option_directives_004 => "Option Strict Off",
    option_directives_005 => "Option Infer On",
    option_directives_006 => "Option Infer Off",
    option_directives_007 => "Option Explicit On\nOption Strict On",
    option_directives_008 => "Option Explicit On\nOption Infer Off",
    option_directives_009 => "Option Strict On\nOption Infer On",
    option_directives_010 => "Option Explicit On\nOption Strict On\nOption Infer On",
}