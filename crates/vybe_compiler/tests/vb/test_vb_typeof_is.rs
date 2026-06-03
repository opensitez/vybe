use super::helpers::run_vb;

macro_rules! vb_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, super::helpers::dotnet_expected_lines(&[$($expected),*]));
        }
    };
}

vb_case!(
    typeof_is_reports_true_for_string_value,
    r#"
Module M
    Sub Main()
        Dim obj As Object = "hello"
        Console.WriteLine(TypeOf obj Is String)
    End Sub
End Module
"#,
    ["true"]
);

vb_case!(
    typeof_is_reports_false_for_wrong_primitive_type,
    r#"
Module M
    Sub Main()
        Dim obj As Object = 42
        Console.WriteLine(TypeOf obj Is Double)
    End Sub
End Module
"#,
    ["false"]
);

vb_case!(
    typeof_is_reports_true_for_integer_value,
    r#"
Module M
    Sub Main()
        Dim obj As Object = 42
        Console.WriteLine(TypeOf obj Is Integer)
    End Sub
End Module
"#,
    ["true"]
);

vb_case!(
    typeof_is_reports_true_for_double_value,
    r#"
Module M
    Sub Main()
        Dim obj As Object = 3.14
        Console.WriteLine(TypeOf obj Is Double)
    End Sub
End Module
"#,
    ["true"]
);

vb_case!(
    typeof_is_reports_true_for_boolean_value,
    r#"
Module M
    Sub Main()
        Dim obj As Object = True
        Console.WriteLine(TypeOf obj Is Boolean)
    End Sub
End Module
"#,
    ["true"]
);

vb_case!(
    typeof_is_does_not_treat_boolean_as_object_type_match,
    r#"
Module M
    Sub Main()
        Dim obj As Object = True
        Console.WriteLine(TypeOf obj Is Object)
    End Sub
End Module
"#,
    ["false"]
);

vb_case!(
    typeof_is_reports_true_for_custom_class_instance,
    r#"
Module M
    Class Greeter
    End Class

    Sub Main()
        Dim obj As Object = New Greeter()
        Console.WriteLine(TypeOf obj Is Greeter)
    End Sub
End Module
"#,
    ["true"]
);

vb_case!(
    typeof_is_reports_false_for_custom_class_against_string,
    r#"
Module M
    Class Greeter
    End Class

    Sub Main()
        Dim obj As Object = New Greeter()
        Console.WriteLine(TypeOf obj Is String)
    End Sub
End Module
"#,
    ["false"]
);
