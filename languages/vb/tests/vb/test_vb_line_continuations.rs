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
    line_continuation_sums_three_terms_across_lines,
    r#"
Module M
    Sub Main()
        Dim total As Integer = 1 + _
            2 + _
            3
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["6"]
);

vb_case!(
    line_continuation_handles_negative_and_positive_terms,
    r#"
Module M
    Sub Main()
        Dim total As Integer = -5 + _
            7 + _
            4
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["6"]
);

vb_case!(
    line_continuation_can_wrap_string_concatenation,
    r#"
Module M
    Sub Main()
        Dim text As String = "Vy" & _
            "be" & _
            "x"
        Console.WriteLine(text)
    End Sub
End Module
"#,
    ["Vybex"]
);

vb_case!(
    line_continuation_can_split_function_arguments,
    r#"
Module M
    Function Add(a As Integer, b As Integer, c As Integer) As Integer
        Return a + b + c
    End Function

    Sub Main()
        Console.WriteLine(Add(1, _
            3, _
            5))
    End Sub
End Module
"#,
    ["9"]
);

vb_case!(
    line_continuation_can_span_parenthesized_expression,
    r#"
Module M
    Sub Main()
        Dim total As Integer = (2 + _
            3) * _
            4
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["20"]
);

vb_case!(
    line_continuation_can_wrap_comparison_expression,
    r#"
Module M
    Sub Main()
        Dim result As Boolean = 1 + _
            2 = _
            3
        If result Then
            Console.WriteLine("match")
        Else
            Console.WriteLine("miss")
        End If
    End Sub
End Module
"#,
    ["match"]
);

vb_case!(
    line_continuation_can_split_assignment_from_function_call,
    r#"
Module M
    Function BuildText() As String
        Return "core"
    End Function

    Sub Main()
        Dim text As String = BuildText() & _
            "-" & _
            "vb"
        Console.WriteLine(text)
    End Sub
End Module
"#,
    ["core-vb"]
);

vb_case!(
    line_continuation_can_chain_multiple_subtractions,
    r#"
Module M
    Sub Main()
        Dim total As Integer = 20 - _
            3 - _
            4
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["13"]
);
