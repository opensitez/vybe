use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_typeof_is(value_expr: &str, type_name: &str, expected: bool) {
    let src = format!(r#"
Module M
    Class Greeter
    End Class

    Sub Main()
        Dim obj As Object = {value_expr}
        Console.WriteLine(TypeOf obj Is {type_name})
    End Sub
End Module
"#, value_expr = value_expr, type_name = type_name);
    assert_vb_output_owned(src, vec![if expected { "true" } else { "false" }.to_string()]);
}

macro_rules! typeof_is_cases {
    ($($name:ident => ($value_expr:expr, $type_name:expr, $expected:expr)),* $(,)?) => {
        $(#[test] fn $name() { assert_typeof_is($value_expr, $type_name, $expected); })*
    };
}

typeof_is_cases! {
    typeof_is_001 => ("\"hello\"", "String", true),
    typeof_is_002 => ("\"hello\"", "Integer", false),
    typeof_is_003 => ("42", "Integer", true),
    typeof_is_004 => ("42", "Double", false),
    typeof_is_005 => ("3.14", "Double", true),
    typeof_is_006 => ("3.14", "String", false),
    typeof_is_007 => ("True", "Boolean", true),
    typeof_is_008 => ("True", "Object", false),
    typeof_is_009 => ("New Greeter()", "Greeter", true),
    typeof_is_010 => ("New Greeter()", "String", false),
}