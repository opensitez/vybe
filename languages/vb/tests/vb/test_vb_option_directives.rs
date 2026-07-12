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
    option_explicit_on_allows_basic_program_structure,
    r#"
Option Explicit On
Module M
    Sub Main()
        Dim total As Integer = 1
        total = total + 1
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["2"]
);

vb_case!(
    option_explicit_off_allows_basic_program_structure,
    r#"
Option Explicit Off
Module M
    Sub Main()
        Dim total As Integer = 2
        total = total + 3
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["5"]
);

vb_case!(
    option_strict_on_allows_typed_arithmetic,
    r#"
Option Strict On
Module M
    Sub Main()
        Dim left As Integer = 4
        Dim right As Integer = 5
        Console.WriteLine(left + right)
    End Sub
End Module
"#,
    ["9"]
);

vb_case!(
    option_strict_off_allows_typed_arithmetic,
    r#"
Option Strict Off
Module M
    Sub Main()
        Dim left As Integer = 7
        Dim right As Integer = 2
        Console.WriteLine(left - right)
    End Sub
End Module
"#,
    ["5"]
);

vb_case!(
    option_infer_on_preserves_local_inference_surface,
    r#"
Option Infer On
Module M
    Sub Main()
        Dim total = 6
        Console.WriteLine(total + 1)
    End Sub
End Module
"#,
    ["7"]
);

vb_case!(
    option_infer_off_with_explicit_type_remains_supported,
    r#"
Option Infer Off
Module M
    Sub Main()
        Dim total As Integer = 8
        Console.WriteLine(total + 2)
    End Sub
End Module
"#,
    ["10"]
);

vb_case!(
    option_directives_can_be_combined_in_pairs,
    r#"
Option Explicit On
Option Strict On
Module M
    Sub Main()
        Dim total As Integer = 3
        Console.WriteLine(total * 4)
    End Sub
End Module
"#,
    ["12"]
);

vb_case!(
    option_directives_can_be_combined_across_all_three_settings,
    r#"
Option Explicit On
Option Strict On
Option Infer On
Module M
    Sub Main()
        Dim total = 9
        Console.WriteLine(total + 3)
    End Sub
End Module
"#,
    ["12"]
);
