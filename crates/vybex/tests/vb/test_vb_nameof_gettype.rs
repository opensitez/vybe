use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_nameof(expression: &str, expected: &str) {
    let src = format!(r#"
Module M
    Function ComputeTotal() As Integer
        Return 10
    End Function

    Sub Main()
        Dim total As Integer = 5
        Console.WriteLine(NameOf({expression}))
    End Sub
End Module
"#, expression = expression);
    assert_vb_output_owned(src, vec![expected.to_string()]);
}

fn assert_gettype(type_name: &str) {
    let src = format!(r#"
Module M
    Sub Main()
        Dim t As Object = GetType({type_name})
        Console.WriteLine(IsNothing(t))
    End Sub
End Module
"#, type_name = type_name);
    assert_vb_output_owned(src, vec!["False".to_string()]);
}

macro_rules! nameof_gettype_cases {
    ($($name:ident => nameof($expression:expr, $expected:expr)),* ; $($name2:ident => gettype($type_name:expr)),* $(,)?) => {
        $(#[test] fn $name() { assert_nameof($expression, $expected); })*
        $(#[test] fn $name2() { assert_gettype($type_name); })*
    };
}

nameof_gettype_cases! {
    nameof_gettype_001 => nameof("total", "total"),
    nameof_gettype_002 => nameof("ComputeTotal", "ComputeTotal"),
    nameof_gettype_003 => nameof("Console", "Console"),
    nameof_gettype_004 => nameof("Main", "Main"),
    nameof_gettype_005 => nameof("Integer", "Integer")
    ;
    nameof_gettype_006 => gettype("Integer"),
    nameof_gettype_007 => gettype("String"),
    nameof_gettype_008 => gettype("Double"),
    nameof_gettype_009 => gettype("Boolean"),
    nameof_gettype_010 => gettype("Object"),
}