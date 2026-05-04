use super::helpers::run_vb;

fn assert_vb_output_owned(src: String, expected: Vec<String>) {
    let out = run_vb(&src);
    assert_eq!(out, expected);
}

fn assert_line_continuation(a: i32, b: i32, c: i32) {
    let src = format!(r#"
Module M
    Sub Main()
        Dim total As Integer = {a} + _
            {b} + _
            {c}
        Console.WriteLine(total)
    End Sub
End Module
"#, a = a, b = b, c = c);
    assert_vb_output_owned(src, vec![(a + b + c).to_string()]);
}

macro_rules! line_continuation_cases {
    ($($name:ident => ($a:expr, $b:expr, $c:expr)),* $(,)?) => {
        $(#[test] fn $name() { assert_line_continuation($a, $b, $c); })*
    };
}

line_continuation_cases! {
    line_continuations_001 => (1, 2, 3),
    line_continuations_002 => (2, 3, 4),
    line_continuations_003 => (3, 4, 5),
    line_continuations_004 => (4, 5, 6),
    line_continuations_005 => (5, 6, 7),
    line_continuations_006 => (6, 7, 8),
    line_continuations_007 => (7, 8, 9),
    line_continuations_008 => (8, 9, 10),
    line_continuations_009 => (9, 10, 11),
    line_continuations_010 => (10, 11, 12),
}